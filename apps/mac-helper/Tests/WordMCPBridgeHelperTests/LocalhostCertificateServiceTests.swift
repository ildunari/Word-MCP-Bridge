import XCTest
@testable import WordMCPBridgeHelper

final class LocalhostCertificateServiceTests: XCTestCase {
    func testResolveExistingCertificateURLsPrefersHelperOwnedCertificate() throws {
        let tempHome = makeTempHome()
        let fileManager = FileManager.default
        let helperURLs = LocalhostCertificateService.helperCertificateURLs(
            fileManager: fileManager,
            homeDirectory: tempHome
        )
        let legacyDir = tempHome
            .appendingPathComponent(".office-addin-dev-certs", isDirectory: true)

        try fileManager.createDirectory(
            at: helperURLs.certificateURL.deletingLastPathComponent(),
            withIntermediateDirectories: true
        )
        try fileManager.createDirectory(at: legacyDir, withIntermediateDirectories: true)
        try Data("helper-cert".utf8).write(to: helperURLs.certificateURL)
        try Data("helper-key".utf8).write(to: helperURLs.keyURL)
        try Data("legacy-cert".utf8).write(to: legacyDir.appendingPathComponent("localhost.crt"))
        try Data("legacy-key".utf8).write(to: legacyDir.appendingPathComponent("localhost.key"))

        let resolved = LocalhostCertificateService.resolveExistingCertificateURLs(
            fileManager: fileManager,
            homeDirectory: tempHome
        )

        XCTAssertEqual(resolved?.certURL.path().removingPercentEncoding ?? resolved?.certURL.path(), helperURLs.certificateURL.path().removingPercentEncoding ?? helperURLs.certificateURL.path())
        XCTAssertEqual(resolved?.keyURL.path().removingPercentEncoding ?? resolved?.keyURL.path(), helperURLs.keyURL.path().removingPercentEncoding ?? helperURLs.keyURL.path())
    }

    func testResolveExistingCertificateURLsFallsBackToLegacyDevCerts() throws {
        let tempHome = makeTempHome()
        let legacyDir = tempHome
            .appendingPathComponent(".office-addin-dev-certs", isDirectory: true)
        try FileManager.default.createDirectory(at: legacyDir, withIntermediateDirectories: true)
        let legacyCert = legacyDir.appendingPathComponent("localhost.crt")
        let legacyKey = legacyDir.appendingPathComponent("localhost.key")
        try Data("legacy-cert".utf8).write(to: legacyCert)
        try Data("legacy-key".utf8).write(to: legacyKey)

        let resolved = LocalhostCertificateService.resolveExistingCertificateURLs(
            fileManager: .default,
            homeDirectory: tempHome
        )

        XCTAssertEqual(resolved?.certURL.path().removingPercentEncoding ?? resolved?.certURL.path(), legacyCert.path().removingPercentEncoding ?? legacyCert.path())
        XCTAssertEqual(resolved?.keyURL.path().removingPercentEncoding ?? resolved?.keyURL.path(), legacyKey.path().removingPercentEncoding ?? legacyKey.path())
    }

    func testHelperCertificateURLsUseApplicationSupportDirectory() {
        let tempHome = makeTempHome()
        let urls = LocalhostCertificateService.helperCertificateURLs(
            fileManager: .default,
            homeDirectory: tempHome
        )

        XCTAssertTrue((urls.certificateURL.path().removingPercentEncoding ?? urls.certificateURL.path()).contains("Library/Application Support/Word MCP Bridge/certs"))
        XCTAssertTrue(urls.opensslConfigURL.lastPathComponent == "openssl-localhost.cnf")
    }

    func testEnsureReadyRetriesTrustWhenHelperCertificateAlreadyExists() throws {
        let tempHome = makeTempHome()
        let fileManager = FileManager.default
        let helperURLs = LocalhostCertificateService.helperCertificateURLs(
            fileManager: fileManager,
            homeDirectory: tempHome
        )
        try fileManager.createDirectory(
            at: helperURLs.certificateURL.deletingLastPathComponent(),
            withIntermediateDirectories: true
        )
        try Data("helper-cert".utf8).write(to: helperURLs.certificateURL)
        try Data("helper-key".utf8).write(to: helperURLs.keyURL)

        var trustCallCount = 0
        let resolved = try LocalhostCertificateService.ensureReady(
            fileManager: fileManager,
            homeDirectory: tempHome,
            trustCertificateHandler: { certURL, _ in
                trustCallCount += 1
                XCTAssertEqual(certURL.path(), helperURLs.certificateURL.path())
            }
        )

        XCTAssertEqual(trustCallCount, 1)
        XCTAssertEqual(resolved.certURL.path(), helperURLs.certificateURL.path())
        XCTAssertEqual(resolved.keyURL.path(), helperURLs.keyURL.path())
    }

    private func makeTempHome() -> URL {
        let homeDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try? FileManager.default.createDirectory(at: homeDirectory, withIntermediateDirectories: true)
        return homeDirectory
    }
}
