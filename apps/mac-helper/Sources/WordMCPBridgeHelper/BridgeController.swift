import AppKit
import Foundation
import SwiftUI

@MainActor
final class BridgeController: NSObject, ObservableObject {
    @Published private(set) var snapshot: HelperSnapshot?
    @Published private(set) var setupState = HelperSetupState(
        bridgeReachable: false,
        connectedSessionCount: 0,
        assetAvailability: .unavailable
    )
    @Published private(set) var isLoading = false
    @Published private(set) var isStarting = false
    @Published private(set) var isStopping = false
    @Published private(set) var lastError: String?

    private var pollTask: Task<Void, Never>?
    private var bridgeProcess: Process?
    private var hasAttemptedAutoStart = false
    private var previousSnapshot: HelperSnapshot?
    private var previousBridgeReachable = false
    private lazy var session: URLSession = {
        URLSession(configuration: .ephemeral, delegate: self, delegateQueue: nil)
    }()

    let baseURL = URL(string: "https://127.0.0.1:4017")!

    var isBridgeRunning: Bool {
        snapshot?.status.running == true
    }

    deinit {
        pollTask?.cancel()
    }

    func startMonitoring() {
        guard pollTask == nil else { return }
        pollTask = Task { [weak self] in
            guard let self else { return }
            self.updateSetupState()
            self.autoStartBridgeIfNeeded()
            await self.requestNotificationsIfNeeded()
            while !Task.isCancelled {
                await self.refresh()
                try? await Task.sleep(for: .seconds(3))
            }
        }
    }

    func refresh() async {
        isLoading = true

        defer { isLoading = false }

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
            lastError = error.localizedDescription
            updateSetupState()
            maybeNotifyOnStateTransition()
        }
    }

    func startBridge() {
        guard bridgeProcess == nil || bridgeProcess?.isRunning == false else { return }
        guard let repoRoot = resolveRepoRoot() else {
            lastError = "Could not locate the Word-MCP-Bridge repo root."
            return
        }

        isStarting = true
        lastError = nil

        let process = Process()
        process.currentDirectoryURL = repoRoot
        process.executableURL = URL(fileURLWithPath: "/usr/bin/env")
        process.arguments = ["pnpm", "bridge:serve"]
        process.terminationHandler = { [weak self] process in
            Task { @MainActor [weak self] in
                self?.bridgeProcess = nil
                self?.isStarting = false
                if process.terminationStatus != 0 {
                    self?.lastError = "Bridge exited with status \(process.terminationStatus)."
                }
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
                await self?.refresh()
            }
        } catch {
            isStarting = false
            lastError = "Failed to start bridge: \(error.localizedDescription)"
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
                let (_, response) = try await session.data(for: request)
                if let http = response as? HTTPURLResponse, !(200 ..< 300).contains(http.statusCode) {
                    throw URLError(.badServerResponse)
                }
                try? await Task.sleep(for: .seconds(1))
                await refresh()
            } catch {
                lastError = "Failed to stop bridge: \(error.localizedDescription)"
            }
        }
    }

    func copyMcpConfig() {
        let block = """
        {
          "mcpServers": {
            "word-mcp-bridge": {
              "command": "npx",
              "args": ["-y", "@word-mcp-bridge/bridge", "mcp-serve", "--url", "https://localhost:4017"]
            }
          }
        }
        """
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
        if let hostedManifestURL = resolvedAssetAvailability().hostedManifestURL {
            NSWorkspace.shared.activateFileViewerSelecting([hostedManifestURL])
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

    func openHostedManifest() {
        guard let hostedManifestURL = resolvedAssetAvailability().hostedManifestURL else {
            lastError = "The hosted manifest is not available in the app bundle or repo."
            return
        }
        NSWorkspace.shared.activateFileViewerSelecting([hostedManifestURL])
    }

    func openLocalManifest() {
        guard let localManifestURL = resolvedAssetAvailability().localManifestURL else {
            lastError = "The local dev manifest is not available in the app bundle or repo."
            return
        }
        NSWorkspace.shared.activateFileViewerSelecting([localManifestURL])
    }

    func openSetupGuide() {
        guard let setupGuideURL = resolvedAssetAvailability().setupGuideURL else {
            lastError = "The setup guide is not available in the app bundle or repo."
            return
        }
        NSWorkspace.shared.open(setupGuideURL)
    }

    func showOnboarding(force: Bool = false) {
        if force || !UserDefaults.standard.bool(forKey: HelperPreferences.hasCompletedOnboardingKey) {
            OnboardingWindowManager.shared.show(controller: self)
        }
    }

    func completeOnboarding() {
        UserDefaults.standard.set(true, forKey: HelperPreferences.hasCompletedOnboardingKey)
    }

    func resetOnboarding() {
        UserDefaults.standard.set(false, forKey: HelperPreferences.hasCompletedOnboardingKey)
        showOnboarding(force: true)
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
            if FileManager.default.fileExists(atPath: url.appending(path: "package.json").path()) {
                return url
            }
        }

        var current = URL(fileURLWithPath: FileManager.default.currentDirectoryPath, isDirectory: true)
        for _ in 0 ..< 8 {
            let candidate = current.appending(path: "package.json")
            if let data = try? Data(contentsOf: candidate),
               let text = String(data: data, encoding: .utf8),
               text.contains("\"name\": \"word-mcp-bridge\"") {
                return current
            }
            current.deleteLastPathComponent()
        }
        return nil
    }

    private func autoStartBridgeIfNeeded() {
        guard !hasAttemptedAutoStart else { return }
        hasAttemptedAutoStart = true
        guard UserDefaults.standard.object(forKey: HelperPreferences.autoStartBridgeOnLaunchKey) == nil
            || UserDefaults.standard.bool(forKey: HelperPreferences.autoStartBridgeOnLaunchKey)
        else {
            return
        }
        guard !isBridgeRunning, bridgeProcess?.isRunning != true else { return }
        startBridge()
    }

    private func resolvedAssetAvailability() -> HelperAssetAvailability {
        HelperAssetAvailability(
            hostedManifestURL: resolvedAssetURL(
                bundledPath: "setup/manifest.prod.xml",
                repoRelativePath: "packages/word-addin/manifest.prod.xml"
            ),
            localManifestURL: resolvedAssetURL(
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
        if let resourceURL = Bundle.main.resourceURL?.appending(path: bundledPath),
           FileManager.default.fileExists(atPath: resourceURL.path()) {
            return resourceURL
        }

        if let repoRoot = resolveRepoRoot() {
            let repoURL = repoRoot.appending(path: repoRelativePath)
            if FileManager.default.fileExists(atPath: repoURL.path()) {
                return repoURL
            }
        }

        return nil
    }

    private func updateSetupState() {
        let assetAvailability = resolvedAssetAvailability()
        setupState = HelperSetupState(
            bridgeReachable: snapshot != nil,
            connectedSessionCount: snapshot?.status.sessionCount ?? 0,
            assetAvailability: assetAvailability
        )
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
