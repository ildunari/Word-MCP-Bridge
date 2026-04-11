import XCTest
@testable import WordMCPBridgeHelper

final class TaskpaneServerTests: XCTestCase {
    @MainActor
    func testBridgeControllerDefaultsToLocalTaskpaneBaseUrl() {
        let controller = BridgeController()

        XCTAssertEqual(controller.taskpaneBaseURL.absoluteString, "https://localhost:3014")
    }

    func testTaskpaneServerLaunchSpecUsesBundledScriptAndAssets() throws {
        let resourcesRoot = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        let taskpaneRoot = resourcesRoot.appendingPathComponent("taskpane", isDirectory: true)
        let serverScript = taskpaneRoot.appendingPathComponent("serve_taskpane.py")
        let distRoot = resourcesRoot.appendingPathComponent("taskpane-dist", isDirectory: true)

        try FileManager.default.createDirectory(at: taskpaneRoot, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: distRoot, withIntermediateDirectories: true)
        try Data("# test".utf8).write(to: serverScript)
        try Data("<html></html>".utf8).write(to: distRoot.appendingPathComponent("taskpane.html"))

        let spec = BridgeController.makeTaskpaneServerLaunchSpec(
            environment: ["PATH": "/usr/bin:/bin"],
            resourcesRoot: resourcesRoot,
            repoRoot: nil,
            port: 3014,
            certURL: URL(fileURLWithPath: "/tmp/localhost.crt"),
            keyURL: URL(fileURLWithPath: "/tmp/localhost.key")
        )

        XCTAssertEqual(spec?.command, "/usr/bin/env")
        XCTAssertEqual(spec?.currentDirectoryURL, resourcesRoot)
        XCTAssertEqual(
            spec?.args,
            [
                "python3",
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
}
