import XCTest
@testable import WordMCPBridgeHelper

final class WordInstallServiceTests: XCTestCase {
    func testEvaluateInstallStatusReportsNotInstalledWhenManifestMissing() throws {
        let tempDir = makeTempDirectory()
        let bundledManifestURL = tempDir.appendingPathComponent("manifest.prod.xml")
        try manifestXML(
            version: "1.0.0.0",
            sourceLocation: "https://localhost:3014/taskpane.html"
        ).write(to: bundledManifestURL, atomically: true, encoding: .utf8)

        let status = WordInstallService.evaluateInstallStatus(
            bundledManifestURL: bundledManifestURL,
            manifestDirectoryURL: tempDir.appendingPathComponent("wef", isDirectory: true),
            wordIsRunning: false
        )

        XCTAssertEqual(status.state, .notInstalled)
        XCTAssertFalse(status.requiresWordRestart)
    }

    func testEvaluateInstallStatusReportsInstalledCurrentForMatchingManifest() throws {
        let tempDir = makeTempDirectory()
        let bundledManifestURL = tempDir.appendingPathComponent("manifest.prod.xml")
        let manifestXML = manifestXML(
            version: "1.0.0.0",
            sourceLocation: "https://localhost:3014/taskpane.html"
        )
        try manifestXML.write(to: bundledManifestURL, atomically: true, encoding: .utf8)

        let wefDir = tempDir.appendingPathComponent("wef", isDirectory: true)
        try FileManager.default.createDirectory(at: wefDir, withIntermediateDirectories: true)
        try manifestXML.write(
            to: wefDir.appendingPathComponent("A89087D2-08B4-481F-8E40-7EB65D0966F9.manifest.xml"),
            atomically: true,
            encoding: .utf8
        )

        let status = WordInstallService.evaluateInstallStatus(
            bundledManifestURL: bundledManifestURL,
            manifestDirectoryURL: wefDir,
            wordIsRunning: false
        )

        XCTAssertEqual(status.state, .installedCurrent)
        XCTAssertTrue(status.isInstalledCurrent)
        XCTAssertFalse(status.requiresWordRestart)
    }

    func testEvaluateInstallStatusReportsInstalledOutdatedForStaleSourceLocation() throws {
        let tempDir = makeTempDirectory()
        let bundledManifestURL = tempDir.appendingPathComponent("manifest.prod.xml")
        try manifestXML(
            version: "1.0.0.0",
            sourceLocation: "https://localhost:3014/taskpane.html"
        ).write(to: bundledManifestURL, atomically: true, encoding: .utf8)

        let wefDir = tempDir.appendingPathComponent("wef", isDirectory: true)
        try FileManager.default.createDirectory(at: wefDir, withIntermediateDirectories: true)
        try manifestXML(
            version: "1.0.0.0",
            sourceLocation: "https://localhost:3013/taskpane.html"
        ).write(
            to: wefDir.appendingPathComponent("A89087D2-08B4-481F-8E40-7EB65D0966F9.manifest.xml"),
            atomically: true,
            encoding: .utf8
        )

        let status = WordInstallService.evaluateInstallStatus(
            bundledManifestURL: bundledManifestURL,
            manifestDirectoryURL: wefDir,
            wordIsRunning: true
        )

        XCTAssertEqual(status.state, .installedOutdated)
        XCTAssertTrue(status.requiresWordRestart)
    }

    func testEvaluateInstallStatusReportsInstallFailedForBrokenInstalledManifest() throws {
        let tempDir = makeTempDirectory()
        let bundledManifestURL = tempDir.appendingPathComponent("manifest.prod.xml")
        try manifestXML(
            version: "1.0.0.0",
            sourceLocation: "https://localhost:3014/taskpane.html"
        ).write(to: bundledManifestURL, atomically: true, encoding: .utf8)

        let wefDir = tempDir.appendingPathComponent("wef", isDirectory: true)
        try FileManager.default.createDirectory(at: wefDir, withIntermediateDirectories: true)
        try "<OfficeApp></OfficeApp>".write(
            to: wefDir.appendingPathComponent("A89087D2-08B4-481F-8E40-7EB65D0966F9.manifest.xml"),
            atomically: true,
            encoding: .utf8
        )

        let status = WordInstallService.evaluateInstallStatus(
            bundledManifestURL: bundledManifestURL,
            manifestDirectoryURL: wefDir,
            wordIsRunning: false
        )

        XCTAssertEqual(status.state, .installFailed)
    }

    func testInstallBundledManifestReplacesStaleCopyAndRemovesDuplicate() throws {
        let tempDir = makeTempDirectory()
        let bundledManifestURL = tempDir.appendingPathComponent("manifest.prod.xml")
        try manifestXML(
            version: "1.0.1.0",
            sourceLocation: "https://localhost:3014/taskpane.html"
        ).write(to: bundledManifestURL, atomically: true, encoding: .utf8)

        let wefDir = tempDir.appendingPathComponent("wef", isDirectory: true)
        try FileManager.default.createDirectory(at: wefDir, withIntermediateDirectories: true)
        try manifestXML(
            version: "1.0.0.0",
            sourceLocation: "https://localhost:3013/taskpane.html"
        ).write(
            to: wefDir.appendingPathComponent("A89087D2-08B4-481F-8E40-7EB65D0966F9.manifest.xml"),
            atomically: true,
            encoding: .utf8
        )
        try manifestXML(
            version: "1.0.0.0",
            sourceLocation: "https://localhost:3013/taskpane.html"
        ).write(
            to: wefDir.appendingPathComponent("duplicate.manifest.xml"),
            atomically: true,
            encoding: .utf8
        )

        let result = try WordInstallService.installBundledManifest(
            bundledManifestURL: bundledManifestURL,
            manifestDirectoryURL: wefDir,
            wordIsRunning: false
        )

        XCTAssertEqual(result.status.state, .installedCurrent)
        XCTAssertTrue(result.changedInstall)
        XCTAssertGreaterThanOrEqual(result.removedManifestCount, 1)
    }

    func testRecommendedActionChoosesRestartWordWhenInstallChangedWhileRunning() {
        let status = WordInstallStatus(
            state: .installedOutdated,
            bundledManifestURL: nil,
            installedManifestURL: nil,
            bundledMetadata: nil,
            installedMetadata: nil,
            requiresWordRestart: true,
            lastError: nil
        )

        XCTAssertEqual(
            WordInstallService.recommendedAction(for: status, wordIsRunning: true),
            .restartWordThenOpen
        )
    }

    private func makeTempDirectory() -> URL {
        let url = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try? FileManager.default.createDirectory(at: url, withIntermediateDirectories: true)
        return url
    }

    private func manifestXML(version: String, sourceLocation: String) -> String {
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xsi:type="TaskPaneApp">
          <Id>A89087D2-08B4-481F-8E40-7EB65D0966F9</Id>
          <Version>\(version)</Version>
          <DefaultSettings>
            <SourceLocation DefaultValue="\(sourceLocation)"/>
          </DefaultSettings>
          <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
            <Resources>
              <bt:Urls>
                <bt:Url id="Taskpane.Url" DefaultValue="\(sourceLocation)"/>
              </bt:Urls>
            </Resources>
          </VersionOverrides>
        </OfficeApp>
        """
    }
}
