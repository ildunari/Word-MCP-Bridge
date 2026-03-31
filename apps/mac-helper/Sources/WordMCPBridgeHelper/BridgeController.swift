import AppKit
import Foundation
import SwiftUI

@MainActor
final class BridgeController: NSObject, ObservableObject {
    @Published private(set) var snapshot: HelperSnapshot?
    @Published private(set) var isLoading = false
    @Published private(set) var isStarting = false
    @Published private(set) var isStopping = false
    @Published private(set) var lastError: String?

    private var pollTask: Task<Void, Never>?
    private var bridgeProcess: Process?
    private var hasAttemptedAutoStart = false
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
            self.autoStartBridgeIfNeeded()
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
            snapshot = HelperSnapshot(fetchedAt: Date(), status: response.status)
            lastError = nil
        } catch {
            snapshot = nil
            lastError = error.localizedDescription
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
              "args": ["-y", "@word-mcp-bridge/bridge", "mcp-serve"]
            }
          }
        }
        """
        NSPasteboard.general.clearContents()
        NSPasteboard.general.setString(block, forType: .string)
    }

    func openRepoReadme() {
        guard let repoRoot = resolveRepoRoot() else { return }
        NSWorkspace.shared.open(repoRoot.appending(path: "README.md"))
    }

    func openManifestFolder() {
        guard let repoRoot = resolveRepoRoot() else { return }
        NSWorkspace.shared.open(repoRoot.appending(path: "packages/word-addin"))
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
