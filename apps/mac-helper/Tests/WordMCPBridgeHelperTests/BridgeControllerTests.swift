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
            resourcesRoot: nil,
            repoRoot: repoRoot,
            bridgeURL: "https://localhost:4017"
        )

        XCTAssertTrue(snippet.contains("\"command\": \"/usr/bin/env\""))
        XCTAssertTrue(snippet.contains(bridgeScriptURL.path))
        XCTAssertTrue(snippet.contains("\"mcp-serve\""))
        XCTAssertFalse(snippet.contains("\"command\": \"npx\""))
    }

    func testMcpConfigUsesBundledRuntimeWrapperWhenAvailable() throws {
        let resourcesRoot = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        let runtimeBinDir = resourcesRoot.appendingPathComponent("runtime/bin", isDirectory: true)
        try FileManager.default.createDirectory(at: runtimeBinDir, withIntermediateDirectories: true)
        let wrapperURL = runtimeBinDir.appendingPathComponent("office-bridge")
        try Data("#!/bin/sh\nexit 0\n".utf8).write(to: wrapperURL)
        try FileManager.default.setAttributes([.posixPermissions: 0o755], ofItemAtPath: wrapperURL.path)

        let snippet = BridgeController.makeMcpConfigSnippet(
            environment: ["PATH": "/usr/bin:/bin"],
            resourcesRoot: resourcesRoot,
            repoRoot: nil,
            bridgeURL: "https://localhost:4017"
        )

        XCTAssertTrue(snippet.contains("\"command\": \"\(wrapperURL.path)\""))
        XCTAssertFalse(snippet.contains("\"command\": \"office-bridge\""))
        XCTAssertTrue(snippet.contains("\"mcp-serve\""))
    }

    func testBridgeLaunchSpecPrefersBundledRuntimeOverPathExecutable() throws {
        let resourcesRoot = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        let runtimeBinDir = resourcesRoot.appendingPathComponent("runtime/bin", isDirectory: true)
        try FileManager.default.createDirectory(at: runtimeBinDir, withIntermediateDirectories: true)
        let bundledExecutable = runtimeBinDir.appendingPathComponent("office-bridge")
        try Data("#!/bin/sh\nexit 0\n".utf8).write(to: bundledExecutable)
        try FileManager.default.setAttributes([.posixPermissions: 0o755], ofItemAtPath: bundledExecutable.path)

        let pathBinDir = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: pathBinDir, withIntermediateDirectories: true)
        let pathExecutable = pathBinDir.appendingPathComponent("office-bridge")
        try Data("#!/bin/sh\nexit 0\n".utf8).write(to: pathExecutable)
        try FileManager.default.setAttributes([.posixPermissions: 0o755], ofItemAtPath: pathExecutable.path)

        let spec = BridgeController.makeBridgeLaunchSpec(
            environment: ["PATH": pathBinDir.path],
            resourcesRoot: resourcesRoot,
            repoRoot: nil
        )

        XCTAssertEqual(spec?.command, bundledExecutable.path)
        XCTAssertEqual(spec?.args, ["serve"])
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
            resourcesRoot: nil,
            repoRoot: nil
        )

        XCTAssertEqual(spec?.command, officeBridgeURL.path)
        XCTAssertEqual(spec?.args, ["serve"])
        XCTAssertNil(spec?.currentDirectoryURL)
    }
}
