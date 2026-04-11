import AppKit
import Foundation
import SwiftUI

struct BridgeLaunchSpec: Equatable {
    let command: String
    let args: [String]
    let currentDirectoryURL: URL?
}

@MainActor
final class BridgeController: NSObject, ObservableObject {
    nonisolated private static let bridgeAutoStartRetryInterval: TimeInterval = 10
    nonisolated private static let taskpaneAutoStartRetryInterval: TimeInterval = 10

    @Published private(set) var snapshot: HelperSnapshot?
    @Published private(set) var setupState = HelperSetupState(
        taskpaneServerReachable: false,
        taskpaneServerProcessRunning: false,
        taskpaneServerStarting: false,
        bridgeReachable: false,
        bridgeProcessRunning: false,
        bridgeStarting: false,
        wordAppRunning: false,
        connectedSessionCount: 0,
        assetAvailability: .unavailable
    )
    @Published private(set) var isLoading = false
    @Published private(set) var isStarting = false
    @Published private(set) var isStopping = false
    @Published private(set) var lastError: String?
    @Published private(set) var isTaskpaneServerStarting = false
    @Published private(set) var taskpaneServerLastError: String?
    @Published private(set) var isWordAddinStarting = false
    @Published private(set) var wordAddinLastError: String?

    private var pollTask: Task<Void, Never>?
    private var bridgeProcess: Process?
    private var taskpaneServerProcess: Process?
    private var wordAddinProcess: Process?
    private var isTaskpaneServerReachable = false
    private var lastBridgeAutoStartAttemptAt: Date?
    private var lastTaskpaneAutoStartAttemptAt: Date?
    private var previousSnapshot: HelperSnapshot?
    private var previousBridgeReachable = false
    private lazy var session: URLSession = {
        URLSession(configuration: .ephemeral, delegate: self, delegateQueue: nil)
    }()

    let baseURL = URL(string: "https://localhost:4017")!
    let taskpaneBaseURL = URL(string: "https://localhost:3014")!

    var isBridgeRunning: Bool {
        setupState.bridgeReachable || setupState.bridgeProcessRunning || isStarting
    }

    var isWordAddinDevRunning: Bool {
        wordAddinProcess?.isRunning == true
    }

    var isWordRunning: Bool {
        !NSRunningApplication.runningApplications(withBundleIdentifier: "com.microsoft.Word").isEmpty
    }

    override init() {
        super.init()
        NotificationCenter.default.addObserver(
            forName: NSApplication.willTerminateNotification,
            object: nil,
            queue: .main
        ) { [weak self] _ in
            Task { @MainActor [weak self] in
                self?.stopTaskpaneServerOnQuit()
                self?.stopWordAddinProcessOnQuit()
            }
        }
    }

    deinit {
        pollTask?.cancel()
    }

    func startMonitoring() {
        guard pollTask == nil else { return }
        pollTask = Task { [weak self] in
            guard let self else { return }
            self.updateSetupState()
            await self.requestNotificationsIfNeeded()
            while !Task.isCancelled {
                self.autoStartTaskpaneServerIfNeeded()
                self.autoStartBridgeIfNeeded()
                await self.refresh(showLoading: false)
                await self.refreshTaskpaneServer()
                try? await Task.sleep(for: .seconds(3))
            }
        }
    }

    func refresh(showLoading: Bool = true) async {
        if showLoading {
            isLoading = true
        }

        defer {
            if showLoading {
                isLoading = false
            }
        }

        do {
            let response: BridgeServerStatusResponse = try await requestJSON(path: "/status")
            previousSnapshot = snapshot
            snapshot = HelperSnapshot(fetchedAt: Date(), status: response.status)
            lastError = nil
            updateSetupState()
            maybeNotifyOnStateTransition()
        } catch {
            previousSnapshot = snapshot
            snapshot = nil
            lastError = describe(error: error)
            updateSetupState()
            maybeNotifyOnStateTransition()
        }
    }

    func startBridge() {
        guard bridgeProcess == nil || bridgeProcess?.isRunning == false else { return }
        guard snapshot == nil else { return }
        let environment = launchEnvironment()
        let repoRoot = resolveRepoRoot()
        guard let launchSpec = Self.makeBridgeLaunchSpec(environment: environment, repoRoot: repoRoot) else {
            lastError =
                "Could not start the bridge. Install `office-bridge` on your PATH, or set WORD_MCP_BRIDGE_REPO_ROOT to a local Word-MCP-Bridge checkout."
            return
        }

        isStarting = true
        lastError = nil

        let process = Process()
        process.currentDirectoryURL = launchSpec.currentDirectoryURL
        process.executableURL = URL(fileURLWithPath: launchSpec.command)
        process.arguments = launchSpec.args
        process.environment = environment
        process.terminationHandler = { [weak self] process in
            Task { @MainActor [weak self] in
                self?.bridgeProcess = nil
                self?.isStarting = false
                if process.terminationStatus != 0 {
                    self?.lastError = self?.friendlyBridgeExitMessage(status: process.terminationStatus)
                }
                self?.updateSetupState()
            }
        }

        let output = Pipe()
        process.standardOutput = output
        process.standardError = output

        do {
            try process.run()
            bridgeProcess = process
            Task { @MainActor [weak self] in
                try? await Task.sleep(for: .seconds(2))
                self?.isStarting = false
                await self?.refresh(showLoading: false)
                if self?.bridgeProcess?.isRunning == true, self?.snapshot == nil, self?.lastError == nil {
                    self?.lastError = "The bridge process started, but the local status endpoint is not reachable yet."
                }
                self?.updateSetupState()
            }
        } catch {
            isStarting = false
            lastError = "Failed to start bridge: \(error.localizedDescription)"
            updateSetupState()
        }
    }

    func stopBridge() {
        guard !isStopping else { return }
        isStopping = true

        Task {
            defer { isStopping = false }
            do {
                var request = URLRequest(url: baseURL.appending(path: "shutdown"))
                request.httpMethod = "POST"
                if let authToken = bridgeAuthToken() {
                    request.setValue(authToken, forHTTPHeaderField: "X-Office-Bridge-Token")
                }
                let (_, response) = try await session.data(for: request)
                if let http = response as? HTTPURLResponse, !(200 ..< 300).contains(http.statusCode) {
                    throw URLError(.badServerResponse)
                }
                try? await Task.sleep(for: .seconds(1))
                await refresh(showLoading: false)
            } catch {
                lastError = "Failed to stop bridge: \(error.localizedDescription)"
            }
        }
    }

    func copyMcpConfig() {
        let block = Self.makeMcpConfigSnippet(
            environment: launchEnvironment(),
            repoRoot: resolveRepoRoot(),
            bridgeURL: "https://localhost:4017"
        )
        NSPasteboard.general.clearContents()
        NSPasteboard.general.setString(block, forType: .string)
    }

    func openRepoReadme() {
        if let setupGuideURL = resolvedAssetAvailability().setupGuideURL {
            NSWorkspace.shared.open(setupGuideURL)
            return
        }
        guard let repoRoot = resolveRepoRoot() else { return }
        NSWorkspace.shared.open(repoRoot.appending(path: "README.md"))
    }

    func openManifestFolder() {
        if let productionManifestURL = resolvedAssetAvailability().productionManifestURL {
            NSWorkspace.shared.activateFileViewerSelecting([productionManifestURL])
            return
        }
        guard let repoRoot = resolveRepoRoot() else { return }
        NSWorkspace.shared.open(repoRoot.appending(path: "packages/word-addin"))
    }

    func openWord() {
        if let wordURL = NSWorkspace.shared.urlForApplication(withBundleIdentifier: "com.microsoft.Word") {
            NSWorkspace.shared.openApplication(at: wordURL, configuration: NSWorkspace.OpenConfiguration())
            return
        }
        lastError = "Microsoft Word is not installed or could not be found."
    }

    func openProductionManifest() {
        guard let productionManifestURL = resolvedAssetAvailability().productionManifestURL else {
            lastError = "The local production manifest is not available in the app bundle or repo."
            return
        }
        NSWorkspace.shared.activateFileViewerSelecting([productionManifestURL])
    }

    func openDevelopmentManifest() {
        guard let developmentManifestURL = resolvedAssetAvailability().developmentManifestURL else {
            lastError = "The local development manifest is not available in the app bundle or repo."
            return
        }
        NSWorkspace.shared.activateFileViewerSelecting([developmentManifestURL])
    }

    func openSetupGuide() {
        guard let setupGuideURL = resolvedAssetAvailability().setupGuideURL else {
            lastError = "The setup guide is not available in the app bundle or repo."
            return
        }
        NSWorkspace.shared.open(setupGuideURL)
    }

    func completeOnboarding() {
        UserDefaults.standard.set(true, forKey: HelperPreferences.hasCompletedOnboardingKey)
        UserDefaults.standard.set(false, forKey: HelperPreferences.dismissSetupBannerKey)
    }

    func resetOnboarding() {
        UserDefaults.standard.set(false, forKey: HelperPreferences.hasCompletedOnboardingKey)
        UserDefaults.standard.set(false, forKey: HelperPreferences.dismissSetupBannerKey)
    }

    func startWordAddinDevSession() {
        guard wordAddinProcess?.isRunning != true else { return }
        guard let repoRoot = resolveRepoRoot() else {
            wordAddinLastError = "Could not locate the Word-MCP-Bridge repo root. Set WORD_MCP_BRIDGE_REPO_ROOT or clone the repo under ~/LocalDev/Word-MCP-Bridge."
            return
        }
        let wordAddinDir = repoRoot.appending(path: "packages/word-addin")
        let manifestPath = wordAddinDir.appending(path: "manifest.xml")
        guard FileManager.default.fileExists(atPath: manifestPath.path()) else {
            wordAddinLastError = "Could not find packages/word-addin/manifest.xml in the repo."
            return
        }

        isWordAddinStarting = true
        wordAddinLastError = nil

        let process = Process()
        process.currentDirectoryURL = wordAddinDir
        process.executableURL = URL(fileURLWithPath: "/usr/bin/env")
        process.arguments = ["pnpm", "exec", "office-addin-debugging", "start", "manifest.xml"]
        process.environment = launchEnvironment()

        let stderrPipe = Pipe()
        if let nullOut = try? FileHandle(forWritingTo: URL(fileURLWithPath: "/dev/null")) {
            process.standardOutput = nullOut
        }
        process.standardError = stderrPipe

        process.terminationHandler = { [weak self] proc in
            Task { @MainActor [weak self] in
                self?.wordAddinProcess = nil
                self?.isWordAddinStarting = false
                if proc.terminationStatus != 0 {
                    let data = stderrPipe.fileHandleForReading.readDataToEndOfFile()
                    let text = String(data: data, encoding: .utf8)?
                        .trimmingCharacters(in: .whitespacesAndNewlines) ?? ""
                    let detail = text.isEmpty ? "exit code \(proc.terminationStatus)" : text
                    self?.wordAddinLastError = "Word dev add-in exited: \(detail)"
                }
                self?.updateSetupState()
            }
        }

        do {
            try process.run()
            wordAddinProcess = process
            Task { @MainActor [weak self] in
                try? await Task.sleep(for: .seconds(1.5))
                self?.isWordAddinStarting = false
                self?.updateSetupState()
            }
        } catch {
            isWordAddinStarting = false
            wordAddinLastError = "Failed to start Word dev add-in: \(error.localizedDescription)"
            updateSetupState()
        }
    }

    func stopWordAddinDevSession() {
        wordAddinLastError = nil
        if let proc = wordAddinProcess, proc.isRunning {
            proc.terminate()
        }
        wordAddinProcess = nil
        isWordAddinStarting = false

        guard let repoRoot = resolveRepoRoot() else {
            updateSetupState()
            return
        }
        let wordAddinDir = repoRoot.appending(path: "packages/word-addin")
        let manifestPath = wordAddinDir.appending(path: "manifest.xml")
        guard FileManager.default.fileExists(atPath: manifestPath.path()) else {
            updateSetupState()
            return
        }

        let stopProcess = Process()
        stopProcess.executableURL = URL(fileURLWithPath: "/usr/bin/env")
        stopProcess.arguments = ["pnpm", "exec", "office-addin-debugging", "stop", "manifest.xml"]
        stopProcess.currentDirectoryURL = wordAddinDir
        stopProcess.environment = launchEnvironment()
        let stderrPipe = Pipe()
        if let nullOut = try? FileHandle(forWritingTo: URL(fileURLWithPath: "/dev/null")) {
            stopProcess.standardOutput = nullOut
        }
        stopProcess.standardError = stderrPipe
        stopProcess.terminationHandler = { [weak self] proc in
            Task { @MainActor [weak self] in
                if proc.terminationStatus != 0 {
                    let data = stderrPipe.fileHandleForReading.readDataToEndOfFile()
                    let text = String(data: data, encoding: .utf8)?
                        .trimmingCharacters(in: .whitespacesAndNewlines) ?? ""
                    let detail = text.isEmpty ? "exit code \(proc.terminationStatus)" : text
                    self?.wordAddinLastError = "Word dev add-in stop exited: \(detail)"
                }
                self?.updateSetupState()
            }
        }

        do {
            try stopProcess.run()
        } catch {
            wordAddinLastError = "Could not run office-addin-debugging stop: \(error.localizedDescription)"
            updateSetupState()
        }
    }

    private func stopWordAddinProcessOnQuit() {
        if let proc = wordAddinProcess, proc.isRunning {
            proc.terminate()
        }
        wordAddinProcess = nil
    }

    func startTaskpaneServer() {
        guard taskpaneServerProcess?.isRunning != true else { return }
        let resourcesRoot = Bundle.main.resourceURL
        guard resolvedTaskpaneAssetRoot(resourcesRoot: resourcesRoot) != nil else {
            taskpaneServerLastError = "Could not find packaged taskpane assets. Rebuild the helper bundle so it includes the local taskpane files."
            isTaskpaneServerReachable = false
            updateSetupState()
            return
        }
        guard let certs = resolvedLocalhostCertificateURLs() else {
            taskpaneServerLastError = "The localhost HTTPS certificate is missing. Install the Office add-in localhost certificate before opening the Word MCP Bridge taskpane."
            isTaskpaneServerReachable = false
            updateSetupState()
            return
        }
        guard let spec = Self.makeTaskpaneServerLaunchSpec(
            environment: launchEnvironment(),
            resourcesRoot: resourcesRoot,
            repoRoot: resolveRepoRoot(),
            port: 3014,
            certURL: certs.certURL,
            keyURL: certs.keyURL
        ) else {
            taskpaneServerLastError = "Could not start the local taskpane server. The helper could not find its bundled server script."
            isTaskpaneServerReachable = false
            updateSetupState()
            return
        }

        isTaskpaneServerStarting = true
        taskpaneServerLastError = nil

        let process = Process()
        process.currentDirectoryURL = spec.currentDirectoryURL
        process.executableURL = URL(fileURLWithPath: spec.command)
        process.arguments = spec.args
        process.environment = launchEnvironment()

        let stderrPipe = Pipe()
        if let nullOut = try? FileHandle(forWritingTo: URL(fileURLWithPath: "/dev/null")) {
            process.standardOutput = nullOut
        }
        process.standardError = stderrPipe

        process.terminationHandler = { [weak self] proc in
            Task { @MainActor [weak self] in
                self?.taskpaneServerProcess = nil
                self?.isTaskpaneServerStarting = false
                self?.isTaskpaneServerReachable = false
                if proc.terminationStatus != 0 {
                    let data = stderrPipe.fileHandleForReading.readDataToEndOfFile()
                    let text = String(data: data, encoding: .utf8)?
                        .trimmingCharacters(in: .whitespacesAndNewlines) ?? ""
                    let detail = text.isEmpty ? "exit code \(proc.terminationStatus)" : text
                    self?.taskpaneServerLastError = "Local taskpane server exited: \(detail)"
                }
                self?.updateSetupState()
            }
        }

        do {
            try process.run()
            taskpaneServerProcess = process
            Task { @MainActor [weak self] in
                try? await Task.sleep(for: .seconds(1))
                self?.isTaskpaneServerStarting = false
                await self?.refreshTaskpaneServer()
                self?.updateSetupState()
            }
        } catch {
            isTaskpaneServerStarting = false
            taskpaneServerLastError = "Failed to start the local taskpane server: \(error.localizedDescription)"
            isTaskpaneServerReachable = false
            updateSetupState()
        }
    }

    func stopTaskpaneServer() {
        taskpaneServerLastError = nil
        if let proc = taskpaneServerProcess, proc.isRunning {
            proc.terminate()
        }
        taskpaneServerProcess = nil
        isTaskpaneServerStarting = false
        isTaskpaneServerReachable = false
        updateSetupState()
    }

    private func stopTaskpaneServerOnQuit() {
        if let proc = taskpaneServerProcess, proc.isRunning {
            proc.terminate()
        }
        taskpaneServerProcess = nil
        isTaskpaneServerReachable = false
    }

    private func refreshTaskpaneServer() async {
        do {
            var request = URLRequest(url: taskpaneBaseURL.appending(path: "healthz"))
            request.timeoutInterval = 2
            let (_, response) = try await session.data(for: request)
            if let http = response as? HTTPURLResponse, (200 ..< 300).contains(http.statusCode) {
                taskpaneServerLastError = nil
                isTaskpaneServerReachable = true
            } else {
                taskpaneServerLastError = "The helper's local taskpane server responded unexpectedly."
                isTaskpaneServerReachable = false
            }
        } catch {
            isTaskpaneServerReachable = false
            if taskpaneServerProcess?.isRunning == true || isTaskpaneServerStarting {
                taskpaneServerLastError = "Could not connect to the helper's local taskpane server."
            }
        }
        updateSetupState()
    }

    private func requestJSON<T: Decodable>(path: String) async throws -> T {
        let url = baseURL.appending(path: path)
        let (data, response) = try await session.data(from: url)
        guard let http = response as? HTTPURLResponse, (200 ..< 300).contains(http.statusCode) else {
            throw URLError(.badServerResponse)
        }
        return try JSONDecoder().decode(T.self, from: data)
    }

    private func resolveRepoRoot() -> URL? {
        if let configured = ProcessInfo.processInfo.environment["WORD_MCP_BRIDGE_REPO_ROOT"] {
            let url = URL(fileURLWithPath: configured, isDirectory: true)
            if isWordMCPBridgeRepo(url) {
                return url
            }
        }

        let home = FileManager.default.homeDirectoryForCurrentUser
        let homeCandidates = [
            Self.appendingRelativePath("LocalDev/Word-MCP-Bridge", to: home),
            Self.appendingRelativePath("Code/Word-MCP-Bridge", to: home),
            Self.appendingRelativePath("Developer/Word-MCP-Bridge", to: home),
            Self.appendingRelativePath("Projects/Word-MCP-Bridge", to: home),
            Self.appendingRelativePath("src/Word-MCP-Bridge", to: home),
        ]
        for candidate in homeCandidates where isWordMCPBridgeRepo(candidate) {
            return candidate
        }

        var current = URL(fileURLWithPath: FileManager.default.currentDirectoryPath, isDirectory: true)
        for _ in 0 ..< 8 {
            if isWordMCPBridgeRepo(current) {
                return current
            }
            current.deleteLastPathComponent()
        }
        return nil
    }

    private func isWordMCPBridgeRepo(_ directory: URL) -> Bool {
        let packageJSON = Self.appendingRelativePath("package.json", to: directory)
        guard let data = try? Data(contentsOf: packageJSON),
              let text = String(data: data, encoding: .utf8) else {
            return false
        }
        return text.contains("\"name\": \"word-mcp-bridge\"")
    }

    private func launchEnvironment() -> [String: String] {
        var env = ProcessInfo.processInfo.environment
        let home = NSHomeDirectory()
        let extraPaths = [
            "/opt/homebrew/bin",
            "/usr/local/bin",
            "\(home)/.volta/bin",
            "\(home)/.local/share/pnpm",
        ]
        let existing = env["PATH"] ?? ""
        var segments: [String] = []
        var seen = Set<String>()
        for segment in extraPaths + existing.split(separator: ":").map(String.init) {
            if !segment.isEmpty, !seen.contains(segment) {
                seen.insert(segment)
                segments.append(segment)
            }
        }
        env["PATH"] = segments.joined(separator: ":")
        return env
    }

    private func autoStartBridgeIfNeeded() {
        guard Self.shouldAutoStartBridge(
            autoStartPreferenceValue: UserDefaults.standard.object(
                forKey: HelperPreferences.autoStartBridgeOnLaunchKey
            ),
            bridgeReachable: setupState.bridgeReachable,
            bridgeProcessRunning: bridgeProcess?.isRunning == true,
            bridgeStarting: isStarting,
            lastAttemptAt: lastBridgeAutoStartAttemptAt,
            now: Date()
        ) else {
            return
        }
        lastBridgeAutoStartAttemptAt = Date()
        startBridge()
    }

    nonisolated static func shouldAutoStartBridge(
        autoStartPreferenceValue: Any?,
        bridgeReachable: Bool,
        bridgeProcessRunning: Bool,
        bridgeStarting: Bool,
        lastAttemptAt: Date?,
        now: Date,
        retryInterval: TimeInterval = bridgeAutoStartRetryInterval
    ) -> Bool {
        let autoStartEnabled: Bool
        if let value = autoStartPreferenceValue as? Bool {
            autoStartEnabled = value
        } else {
            autoStartEnabled = true
        }

        guard autoStartEnabled else { return false }
        guard !bridgeReachable, !bridgeProcessRunning, !bridgeStarting else { return false }

        if let lastAttemptAt, now.timeIntervalSince(lastAttemptAt) < retryInterval {
            return false
        }

        return true
    }

    private func autoStartTaskpaneServerIfNeeded() {
        guard Self.shouldAutoStartTaskpaneServer(
            taskpaneReachable: isTaskpaneServerReachable,
            taskpaneProcessRunning: taskpaneServerProcess?.isRunning == true,
            taskpaneStarting: isTaskpaneServerStarting,
            lastAttemptAt: lastTaskpaneAutoStartAttemptAt,
            now: Date()
        ) else {
            return
        }
        lastTaskpaneAutoStartAttemptAt = Date()
        startTaskpaneServer()
    }

    nonisolated static func shouldAutoStartTaskpaneServer(
        taskpaneReachable: Bool,
        taskpaneProcessRunning: Bool,
        taskpaneStarting: Bool,
        lastAttemptAt: Date?,
        now: Date,
        retryInterval: TimeInterval = taskpaneAutoStartRetryInterval
    ) -> Bool {
        guard !taskpaneReachable, !taskpaneProcessRunning, !taskpaneStarting else { return false }

        if let lastAttemptAt, now.timeIntervalSince(lastAttemptAt) < retryInterval {
            return false
        }

        return true
    }

    private func resolvedAssetAvailability() -> HelperAssetAvailability {
        HelperAssetAvailability(
            productionManifestURL: resolvedAssetURL(
                bundledPath: "setup/manifest.prod.xml",
                repoRelativePath: "packages/word-addin/manifest.prod.xml"
            ),
            developmentManifestURL: resolvedAssetURL(
                bundledPath: "setup/manifest.xml",
                repoRelativePath: "packages/word-addin/manifest.xml"
            ),
            setupGuideURL: resolvedAssetURL(
                bundledPath: "setup/SETUP-GUIDE.md",
                repoRelativePath: "apps/mac-helper/SETUP-GUIDE.md"
            )
        )
    }

    private func resolvedAssetURL(bundledPath: String, repoRelativePath: String) -> URL? {
        if let resourceRoot = Bundle.main.resourceURL {
            let resourceURL = Self.appendingRelativePath(bundledPath, to: resourceRoot)
            if FileManager.default.fileExists(atPath: resourceURL.path()) {
                return resourceURL
            }
        }

        if let repoRoot = resolveRepoRoot() {
            let repoURL = Self.appendingRelativePath(repoRelativePath, to: repoRoot)
            if FileManager.default.fileExists(atPath: repoURL.path()) {
                return repoURL
            }
        }

        return nil
    }

    private func resolvedTaskpaneAssetRoot(resourcesRoot: URL?) -> URL? {
        if let resourcesRoot {
            let bundled = Self.appendingRelativePath("taskpane-dist", to: resourcesRoot)
            if FileManager.default.fileExists(atPath: bundled.appending(path: "taskpane.html").path()) {
                return bundled
            }
        }

        if let repoRoot = resolveRepoRoot() {
            let repoDist = Self.appendingRelativePath("packages/word-addin/dist", to: repoRoot)
            if FileManager.default.fileExists(atPath: repoDist.appending(path: "taskpane.html").path()) {
                return repoDist
            }
        }

        return nil
    }

    private func resolvedLocalhostCertificateURLs() -> (certURL: URL, keyURL: URL)? {
        let home = FileManager.default.homeDirectoryForCurrentUser
        let certURL = home.appending(path: ".office-addin-dev-certs/localhost.crt")
        let keyURL = home.appending(path: ".office-addin-dev-certs/localhost.key")
        guard FileManager.default.fileExists(atPath: certURL.path()),
              FileManager.default.fileExists(atPath: keyURL.path()) else {
            return nil
        }
        return (certURL, keyURL)
    }

    nonisolated static func appendingRelativePath(_ relativePath: String, to baseURL: URL) -> URL {
        relativePath
            .split(separator: "/", omittingEmptySubsequences: true)
            .reduce(baseURL) { partialURL, component in
                partialURL.appendingPathComponent(String(component), isDirectory: false)
            }
    }

    nonisolated static func makeBridgeLaunchSpec(environment: [String: String], repoRoot: URL?) -> BridgeLaunchSpec? {
        if let repoRoot, let bridgeScriptURL = repoBridgeScriptURL(repoRoot: repoRoot) {
            return BridgeLaunchSpec(
                command: "/usr/bin/env",
                args: ["node", bridgeScriptURL.path(), "serve"],
                currentDirectoryURL: repoRoot
            )
        }

        if let officeBridgePath = findExecutable(named: "office-bridge", environment: environment) {
            return BridgeLaunchSpec(
                command: officeBridgePath,
                args: ["serve"],
                currentDirectoryURL: nil
            )
        }

        if let repoRoot {
            return BridgeLaunchSpec(
                command: "/usr/bin/env",
                args: ["pnpm", "bridge:serve"],
                currentDirectoryURL: repoRoot
            )
        }

        return nil
    }

    nonisolated static func makeTaskpaneServerLaunchSpec(
        environment: [String: String],
        resourcesRoot: URL?,
        repoRoot: URL?,
        port: Int,
        certURL: URL,
        keyURL: URL
    ) -> BridgeLaunchSpec? {
        let bundledScript = resourcesRoot.map { appendingRelativePath("taskpane/serve_taskpane.py", to: $0) }
        let bundledAssets = resourcesRoot.map { appendingRelativePath("taskpane-dist", to: $0) }

        if let scriptURL = bundledScript,
           let assetRoot = bundledAssets,
           FileManager.default.fileExists(atPath: scriptURL.path()),
           FileManager.default.fileExists(atPath: assetRoot.appending(path: "taskpane.html").path()) {
            return BridgeLaunchSpec(
                command: "/usr/bin/env",
                args: [
                    "python3",
                    scriptURL.path,
                    "--root",
                    assetRoot.path,
                    "--port",
                    "\(port)",
                    "--cert",
                    certURL.path,
                    "--key",
                    keyURL.path,
                ],
                currentDirectoryURL: resourcesRoot
            )
        }

        if let repoRoot {
            let scriptURL = appendingRelativePath("apps/mac-helper/Scripts/serve_taskpane.py", to: repoRoot)
            let assetRoot = appendingRelativePath("packages/word-addin/dist", to: repoRoot)
            if FileManager.default.fileExists(atPath: scriptURL.path()),
               FileManager.default.fileExists(atPath: assetRoot.appending(path: "taskpane.html").path()) {
                return BridgeLaunchSpec(
                    command: "/usr/bin/env",
                    args: [
                        "python3",
                        scriptURL.path,
                        "--root",
                        assetRoot.path,
                        "--port",
                        "\(port)",
                        "--cert",
                        certURL.path,
                        "--key",
                        keyURL.path,
                    ],
                    currentDirectoryURL: repoRoot
                )
            }
        }

        return nil
    }

    nonisolated static func makeMcpConfigSnippet(environment: [String: String], repoRoot: URL?, bridgeURL: String) -> String {
        let launchSpec = makeMcpLaunchSpec(environment: environment, repoRoot: repoRoot, bridgeURL: bridgeURL)

        let argsJSON = launchSpec.args
            .map { "\"\($0.replacingOccurrences(of: "\\", with: "\\\\").replacingOccurrences(of: "\"", with: "\\\""))\"" }
            .joined(separator: ", ")

        return """
        {
          "mcpServers": {
            "word-mcp-bridge": {
              "command": "\(launchSpec.command)",
              "args": [\(argsJSON)]
            }
          }
        }
        """
    }

    nonisolated static func makeMcpLaunchSpec(environment: [String: String], repoRoot: URL?, bridgeURL: String) -> BridgeLaunchSpec {
        if let repoRoot, let bridgeScriptURL = repoBridgeScriptURL(repoRoot: repoRoot) {
            return BridgeLaunchSpec(
                command: "/usr/bin/env",
                args: ["node", bridgeScriptURL.path(), "mcp-serve", "--url", bridgeURL],
                currentDirectoryURL: repoRoot
            )
        }

        if let officeBridgePath = findExecutable(named: "office-bridge", environment: environment) {
            return BridgeLaunchSpec(
                command: officeBridgePath,
                args: ["mcp-serve", "--url", bridgeURL],
                currentDirectoryURL: nil
            )
        }

        return BridgeLaunchSpec(
            command: "office-bridge",
            args: ["mcp-serve", "--url", bridgeURL],
            currentDirectoryURL: nil
        )
    }

    private nonisolated static func repoBridgeScriptURL(repoRoot: URL) -> URL? {
        let scriptURL = appendingRelativePath("packages/bridge/bin/office-bridge.js", to: repoRoot)
        return FileManager.default.fileExists(atPath: scriptURL.path()) ? scriptURL : nil
    }

    private nonisolated static func findExecutable(named executable: String, environment: [String: String]) -> String? {
        let searchPath = environment["PATH"] ?? ProcessInfo.processInfo.environment["PATH"] ?? ""
        for pathEntry in searchPath.split(separator: ":") {
            let candidate = URL(fileURLWithPath: String(pathEntry), isDirectory: true)
                .appendingPathComponent(executable, isDirectory: false)
            if FileManager.default.isExecutableFile(atPath: candidate.path()) {
                return candidate.path()
            }
        }
        return nil
    }

    private func updateSetupState() {
        let assetAvailability = resolvedAssetAvailability()
        setupState = HelperSetupState(
            taskpaneServerReachable: isTaskpaneServerReachable,
            taskpaneServerProcessRunning: taskpaneServerProcess?.isRunning == true,
            taskpaneServerStarting: isTaskpaneServerStarting,
            bridgeReachable: snapshot != nil,
            bridgeProcessRunning: bridgeProcess?.isRunning == true,
            bridgeStarting: isStarting,
            wordAppRunning: isWordRunning,
            connectedSessionCount: snapshot?.status.sessionCount ?? 0,
            assetAvailability: assetAvailability
        )
    }

    private func bridgeAuthToken() -> String? {
        if let envToken = ProcessInfo.processInfo.environment["OFFICE_BRIDGE_TOKEN"]?
            .trimmingCharacters(in: .whitespacesAndNewlines),
            !envToken.isEmpty {
            return envToken
        }

        let tokenURL = URL(fileURLWithPath: NSTemporaryDirectory(), isDirectory: true)
            .appending(path: "office-agents-bridge", directoryHint: .isDirectory)
            .appending(path: "auth-token", directoryHint: .notDirectory)
        guard let token = try? String(contentsOf: tokenURL, encoding: .utf8)
            .trimmingCharacters(in: .whitespacesAndNewlines),
            !token.isEmpty else {
            return nil
        }
        return token
    }

    private func friendlyBridgeExitMessage(status: Int32) -> String {
        if status != 0 {
            return "Bridge exited with status \(status). If port 4017 is already in use, stop the existing bridge server before starting a new one."
        }
        return "Bridge exited."
    }

    private func describe(error: Error) -> String {
        if error is DecodingError {
            return "Bridge status response was not in the format the helper expected."
        }
        if let urlError = error as? URLError {
            switch urlError.code {
            case .cannotConnectToHost, .cannotFindHost, .timedOut:
                return "Could not connect to the local bridge server."
            case .badServerResponse:
                return "The bridge responded unexpectedly. Make sure the local bridge server is running on port 4017 and using the same auth token."
            default:
                break
            }
        }
        return error.localizedDescription
    }

    private func requestNotificationsIfNeeded() async {
        guard UserDefaults.standard.object(forKey: HelperPreferences.notificationsEnabledKey) == nil
            || UserDefaults.standard.bool(forKey: HelperPreferences.notificationsEnabledKey)
        else {
            return
        }

        await HelperNotifications.requestAuthorizationIfNeeded()
    }

    private func maybeNotifyOnStateTransition() {
        let notificationsEnabled =
            UserDefaults.standard.object(forKey: HelperPreferences.notificationsEnabledKey) == nil
            || UserDefaults.standard.bool(forKey: HelperPreferences.notificationsEnabledKey)
        guard notificationsEnabled else {
            previousBridgeReachable = snapshot != nil
            return
        }

        let currentBridgeReachable = snapshot != nil
        let previousConnected = (previousSnapshot?.status.sessionCount ?? 0) > 0
        let currentConnected = (snapshot?.status.sessionCount ?? 0) > 0

        if !previousConnected && currentConnected {
            HelperNotifications.post(
                title: "Word connected",
                body: "A Word taskpane session is connected to Word MCP Bridge."
            )
        } else if previousConnected && !currentConnected {
            HelperNotifications.post(
                title: "Word disconnected",
                body: "The helper lost its live Word taskpane session. Open Word and the add-in to reconnect."
            )
        }

        if previousBridgeReachable && !currentBridgeReachable && !isStopping {
            HelperNotifications.post(
                title: "Bridge not reachable",
                body: "The local bridge stopped responding. Reopen the helper or restart the bridge."
            )
        }

        previousBridgeReachable = currentBridgeReachable
    }
}

extension BridgeController: URLSessionDelegate {
    nonisolated func urlSession(
        _ session: URLSession,
        didReceive challenge: URLAuthenticationChallenge,
        completionHandler: @escaping (URLSession.AuthChallengeDisposition, URLCredential?) -> Void
    ) {
        guard challenge.protectionSpace.host == "127.0.0.1" || challenge.protectionSpace.host == "localhost",
              let trust = challenge.protectionSpace.serverTrust else {
            completionHandler(.performDefaultHandling, nil)
            return
        }

        completionHandler(.useCredential, URLCredential(trust: trust))
    }
}
