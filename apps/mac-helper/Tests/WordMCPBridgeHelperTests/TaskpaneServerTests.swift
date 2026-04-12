import XCTest
@testable import WordMCPBridgeHelper

final class TaskpaneServerTests: XCTestCase {
    func testLaunchWordTaskpaneHelpMentionsHiddenSharedRuntimeSuccess() throws {
        let repoRoot = URL(fileURLWithPath: #filePath)
            .deletingLastPathComponent()
            .deletingLastPathComponent()
            .deletingLastPathComponent()
            .deletingLastPathComponent()
            .deletingLastPathComponent()
        let scriptURL = repoRoot.appendingPathComponent("scripts/bridge/launch-word-taskpane.sh")

        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/bin/bash")
        process.arguments = [scriptURL.path, "--help"]

        let output = Pipe()
        process.standardOutput = output
        process.standardError = output

        try process.run()
        process.waitUntilExit()

        let data = output.fileHandleForReading.readDataToEndOfFile()
        let text = String(decoding: data, as: UTF8.self)

        XCTAssertEqual(process.terminationStatus, 0)
        XCTAssertTrue(text.contains("hidden shared-runtime session"))
    }

    @MainActor
    func testBridgeControllerDefaultsToLocalTaskpaneBaseUrl() {
        let controller = BridgeController()

        XCTAssertEqual(controller.taskpaneBaseURL.absoluteString, "https://localhost:3014")
    }

    func testTaskpaneServerLaunchSpecUsesBundledScriptAndAssets() throws {
        let resourcesRoot = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        let taskpaneRoot = resourcesRoot.appendingPathComponent("taskpane", isDirectory: true)
        let serverScript = taskpaneRoot.appendingPathComponent("serve_taskpane.mjs")
        let distRoot = resourcesRoot.appendingPathComponent("taskpane-dist", isDirectory: true)
        let bundledNodeRoot = resourcesRoot.appendingPathComponent("runtime/node/bin", isDirectory: true)
        let bundledNode = bundledNodeRoot.appendingPathComponent("node")

        try FileManager.default.createDirectory(at: taskpaneRoot, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: distRoot, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: bundledNodeRoot, withIntermediateDirectories: true)
        try Data("# test".utf8).write(to: serverScript)
        try Data("#!/bin/sh\nexit 0\n".utf8).write(to: bundledNode)
        try FileManager.default.setAttributes([.posixPermissions: 0o755], ofItemAtPath: bundledNode.path)
        try Data("<html></html>".utf8).write(to: distRoot.appendingPathComponent("taskpane.html"))

        let spec = BridgeController.makeTaskpaneServerLaunchSpec(
            environment: ["PATH": "/usr/bin:/bin"],
            resourcesRoot: resourcesRoot,
            repoRoot: nil,
            port: 3014,
            certURL: URL(fileURLWithPath: "/tmp/localhost.crt"),
            keyURL: URL(fileURLWithPath: "/tmp/localhost.key")
        )

        XCTAssertEqual(spec?.command, bundledNode.path)
        XCTAssertEqual(spec?.currentDirectoryURL, resourcesRoot)
        XCTAssertEqual(
            spec?.args,
            [
                serverScript.path,
                "--root",
                distRoot.path,
                "--port",
                "3014",
                "--cert",
                "/tmp/localhost.crt",
                "--key",
                "/tmp/localhost.key",
            ]
        )
    }

    func testTaskpaneLauncherSpecUsesBundledCompiledLauncherBinary() throws {
        let resourcesRoot = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        let launcherBinDir = resourcesRoot.appendingPathComponent("runtime/bin", isDirectory: true)
        let launcherURL = launcherBinDir.appendingPathComponent("word-mcp-bridge-word-launcher")
        try FileManager.default.createDirectory(at: launcherBinDir, withIntermediateDirectories: true)
        try Data("#!/bin/sh\nexit 0\n".utf8).write(to: launcherURL)
        try FileManager.default.setAttributes([.posixPermissions: 0o755], ofItemAtPath: launcherURL.path)

        let spec = BridgeController.makeTaskpaneLauncherSpec(
            resourcesRoot: resourcesRoot,
            repoRoot: nil,
            bridgeURL: "https://localhost:4017",
            timeoutSeconds: 20
        )

        XCTAssertEqual(spec?.command, launcherURL.path)
        XCTAssertEqual(spec?.args, ["--mode", "open"])
        XCTAssertEqual(spec?.currentDirectoryURL, resourcesRoot)
    }

    func testWordSessionLaunchProgressRejectsUnchangedExistingSession() {
        let baseline = [makeWordSession(sessionId: "word:existing", connectedAt: 1_000, lastSeenAt: 2_000)]
        let current = [makeWordSession(sessionId: "word:existing", connectedAt: 1_000, lastSeenAt: 2_000)]

        let didSucceed = BridgeController.didLaunchWordSessionSuccessfully(
            baselineWordSessions: baseline,
            currentSessions: current,
            launchStartedAtMs: 3_000,
            launcherOutput: "Triggered the Word MCP Bridge ribbon command."
        )

        XCTAssertFalse(didSucceed)
    }

    func testWordSessionLaunchProgressAcceptsNewSessionId() {
        let baseline = [makeWordSession(sessionId: "word:existing", connectedAt: 1_000, lastSeenAt: 2_000)]
        let current = [
            makeWordSession(sessionId: "word:existing", connectedAt: 1_000, lastSeenAt: 2_000),
            makeWordSession(sessionId: "word:new", connectedAt: 4_000, lastSeenAt: 4_000),
        ]

        let didSucceed = BridgeController.didLaunchWordSessionSuccessfully(
            baselineWordSessions: baseline,
            currentSessions: current,
            launchStartedAtMs: 3_000,
            launcherOutput: "Triggered the Word MCP Bridge ribbon command."
        )

        XCTAssertTrue(didSucceed)
    }

    func testWordSessionLaunchProgressAllowsAlreadyOpenPaneWithoutNewSession() {
        let baseline = [makeWordSession(sessionId: "word:existing", connectedAt: 1_000, lastSeenAt: 2_000)]
        let current = [makeWordSession(sessionId: "word:existing", connectedAt: 1_000, lastSeenAt: 2_000)]

        let didSucceed = BridgeController.didLaunchWordSessionSuccessfully(
            baselineWordSessions: baseline,
            currentSessions: current,
            launchStartedAtMs: 3_000,
            launcherOutput: "Word MCP Bridge pane is already open."
        )

        XCTAssertTrue(didSucceed)
    }

    func testBridgeAutoStartDefaultsToEnabledAndRetriesWhenBridgeIsDown() {
        let shouldAutoStart = BridgeController.shouldAutoStartBridge(
            autoStartPreferenceValue: nil,
            bridgeReachable: false,
            bridgeProcessRunning: false,
            bridgeStarting: false,
            lastAttemptAt: nil,
            now: Date()
        )

        XCTAssertTrue(shouldAutoStart)
    }

    func testBridgeAutoStartHonorsRetryCooldown() {
        let now = Date()

        let shouldAutoStart = BridgeController.shouldAutoStartBridge(
            autoStartPreferenceValue: true,
            bridgeReachable: false,
            bridgeProcessRunning: false,
            bridgeStarting: false,
            lastAttemptAt: now.addingTimeInterval(-2),
            now: now,
            retryInterval: 10
        )

        XCTAssertFalse(shouldAutoStart)
    }

    func testBridgeAutoStartCanRetryAfterCooldownExpires() {
        let now = Date()

        let shouldAutoStart = BridgeController.shouldAutoStartBridge(
            autoStartPreferenceValue: true,
            bridgeReachable: false,
            bridgeProcessRunning: false,
            bridgeStarting: false,
            lastAttemptAt: now.addingTimeInterval(-15),
            now: now,
            retryInterval: 10
        )

        XCTAssertTrue(shouldAutoStart)
    }

    func testBridgeAutoStartStopsWhenPreferenceIsDisabled() {
        let shouldAutoStart = BridgeController.shouldAutoStartBridge(
            autoStartPreferenceValue: false,
            bridgeReachable: false,
            bridgeProcessRunning: false,
            bridgeStarting: false,
            lastAttemptAt: nil,
            now: Date()
        )

        XCTAssertFalse(shouldAutoStart)
    }

    func testTaskpaneAutoStartRetriesWhenServerIsDown() {
        let shouldAutoStart = BridgeController.shouldAutoStartTaskpaneServer(
            taskpaneReachable: false,
            taskpaneProcessRunning: false,
            taskpaneStarting: false,
            lastAttemptAt: nil,
            now: Date()
        )

        XCTAssertTrue(shouldAutoStart)
    }

    func testTaskpaneAutoStartHonorsRetryCooldown() {
        let now = Date()

        let shouldAutoStart = BridgeController.shouldAutoStartTaskpaneServer(
            taskpaneReachable: false,
            taskpaneProcessRunning: false,
            taskpaneStarting: false,
            lastAttemptAt: now.addingTimeInterval(-3),
            now: now,
            retryInterval: 10
        )

        XCTAssertFalse(shouldAutoStart)
    }

    func testTaskpaneAutoStartStopsWhenServerIsAlreadyHealthy() {
        let shouldAutoStart = BridgeController.shouldAutoStartTaskpaneServer(
            taskpaneReachable: true,
            taskpaneProcessRunning: true,
            taskpaneStarting: false,
            lastAttemptAt: nil,
            now: Date()
        )

        XCTAssertFalse(shouldAutoStart)
    }

    func testTaskpaneAutoOpenRunsWhenWordLaunchesIntoHealthyInstalledSetup() {
        let shouldAutoOpen = BridgeController.shouldAutoOpenTaskpane(
            autoOpenPreferenceValue: true,
            installReady: true,
            certificateReady: true,
            taskpaneLauncherReady: true,
            taskpaneReachable: true,
            bridgeReachable: true,
            wordRunning: true,
            hasWordSession: false,
            isOpeningTaskpane: false,
            restartRequired: false,
            lastAttemptAt: nil,
            now: Date()
        )

        XCTAssertTrue(shouldAutoOpen)
    }

    func testTaskpaneAutoOpenStopsWhenRestartIsRequired() {
        let shouldAutoOpen = BridgeController.shouldAutoOpenTaskpane(
            autoOpenPreferenceValue: true,
            installReady: true,
            certificateReady: true,
            taskpaneLauncherReady: true,
            taskpaneReachable: true,
            bridgeReachable: true,
            wordRunning: true,
            hasWordSession: false,
            isOpeningTaskpane: false,
            restartRequired: true,
            lastAttemptAt: nil,
            now: Date()
        )

        XCTAssertFalse(shouldAutoOpen)
    }

    func testTaskpaneAutoOpenStopsWhenCertificateIsNotReady() {
        let shouldAutoOpen = BridgeController.shouldAutoOpenTaskpane(
            autoOpenPreferenceValue: true,
            installReady: true,
            certificateReady: false,
            taskpaneLauncherReady: true,
            taskpaneReachable: true,
            bridgeReachable: true,
            wordRunning: true,
            hasWordSession: false,
            isOpeningTaskpane: false,
            restartRequired: false,
            lastAttemptAt: nil,
            now: Date()
        )

        XCTAssertFalse(shouldAutoOpen)
    }

    func testTaskpaneAutoOpenStopsWhenAccessibilityAutomationIsNotReady() {
        let shouldAutoOpen = BridgeController.shouldAutoOpenTaskpane(
            autoOpenPreferenceValue: true,
            installReady: true,
            certificateReady: true,
            taskpaneLauncherReady: false,
            taskpaneReachable: true,
            bridgeReachable: true,
            wordRunning: true,
            hasWordSession: false,
            isOpeningTaskpane: false,
            restartRequired: false,
            lastAttemptAt: nil,
            now: Date()
        )

        XCTAssertFalse(shouldAutoOpen)
    }

    private func makeWordSession(sessionId: String, connectedAt: Int64, lastSeenAt: Int64) -> BridgeSessionPayload {
        BridgeSessionPayload(
            snapshot: BridgeSessionSnapshotPayload(
                sessionId: sessionId,
                instanceId: nil,
                app: "word",
                appName: "Microsoft Word",
                documentId: "word-local:\(sessionId)",
                documentMetadata: BridgeMetadataPayload(documentId: nil, title: "Doc", url: nil),
                runtimeState: nil,
                gateway: nil
            ),
            connectedAt: connectedAt,
            lastSeenAt: lastSeenAt,
            pendingCount: 0,
            metrics: BridgeMetricsPayload(
                eventCount: 0,
                toolCallCount: 0,
                bridgeErrorCount: 0,
                requestTimeoutCount: 0,
                connectionDropCount: 0,
                disconnectedSessionCount: nil,
                connectedSessionCount: nil,
                pendingCount: nil
            )
        )
    }
}
