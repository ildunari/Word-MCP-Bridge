import XCTest
@testable import WordMCPBridgeHelper

final class BridgeControllerTests: XCTestCase {
    @MainActor
    func testBridgeControllerDefaultsToLocalhostBaseUrl() {
        let controller = BridgeController()

        XCTAssertEqual(controller.baseURL.absoluteString, "https://localhost:4017")
    }

    func testAppendsRelativePathInsideDirectoriesWithSpaces() throws {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        let baseDir = tempDir.appendingPathComponent("Word MCP Bridge Helper.app", isDirectory: true)
        let resourcesDir = baseDir.appendingPathComponent("Contents/Resources", isDirectory: true)
        let setupDir = resourcesDir.appendingPathComponent("setup", isDirectory: true)
        try FileManager.default.createDirectory(at: setupDir, withIntermediateDirectories: true)

        let manifestURL = setupDir.appendingPathComponent("manifest.prod.xml")
        try Data("<manifest/>".utf8).write(to: manifestURL)

        let resolved = BridgeController.appendingRelativePath("setup/manifest.prod.xml", to: resourcesDir)

        XCTAssertEqual(resolved.path, manifestURL.path)
        XCTAssertTrue(FileManager.default.fileExists(atPath: resolved.path))
    }

    func testMcpConfigUsesRepoBridgeScriptWhenRepoIsAvailable() throws {
        let repoRoot = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        let bridgeBinDir = repoRoot.appendingPathComponent("packages/bridge/bin", isDirectory: true)
        try FileManager.default.createDirectory(at: bridgeBinDir, withIntermediateDirectories: true)
        let bridgeScriptURL = bridgeBinDir.appendingPathComponent("office-bridge.js")
        try Data("console.log('bridge')".utf8).write(to: bridgeScriptURL)

        let snippet = BridgeController.makeMcpConfigSnippet(
            environment: ["PATH": "/usr/bin:/bin"],
            repoRoot: repoRoot,
            bridgeURL: "https://localhost:4017"
        )

        XCTAssertTrue(snippet.contains("\"command\": \"/usr/bin/env\""))
        XCTAssertTrue(snippet.contains(bridgeScriptURL.path))
        XCTAssertTrue(snippet.contains("\"mcp-serve\""))
        XCTAssertFalse(snippet.contains("\"command\": \"npx\""))
    }

    func testBridgeLaunchSpecFallsBackToOfficeBridgeExecutableFromPath() throws {
        let binDir = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: binDir, withIntermediateDirectories: true)

        let officeBridgeURL = binDir.appendingPathComponent("office-bridge")
        try Data("#!/bin/sh\nexit 0\n".utf8).write(to: officeBridgeURL)
        try FileManager.default.setAttributes(
            [.posixPermissions: 0o755],
            ofItemAtPath: officeBridgeURL.path
        )

        let spec = BridgeController.makeBridgeLaunchSpec(
            environment: ["PATH": binDir.path],
            repoRoot: nil
        )

        XCTAssertEqual(spec?.command, officeBridgeURL.path)
        XCTAssertEqual(spec?.args, ["serve"])
        XCTAssertNil(spec?.currentDirectoryURL)
    }
}
