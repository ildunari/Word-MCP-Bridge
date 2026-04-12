import AppKit
import CryptoKit
import Foundation

enum WordInstallState: String {
    case unavailable
    case notInstalled
    case installedCurrent
    case installedOutdated
    case installFailed

    var isReady: Bool {
        self == .installedCurrent
    }
}

enum WordInstallRecommendation: Equatable {
    case installThenOpen
    case reinstallThenOpen
    case restartWordThenOpen
    case openOnly
}

struct WordManifestMetadata: Equatable {
    let addinID: String
    let version: String?
    let sourceLocation: String?
    let rawHash: String
}

struct WordInstallStatus: Equatable {
    let state: WordInstallState
    let bundledManifestURL: URL?
    let installedManifestURL: URL?
    let bundledMetadata: WordManifestMetadata?
    let installedMetadata: WordManifestMetadata?
    let requiresWordRestart: Bool
    let lastError: String?

    static let unavailable = WordInstallStatus(
        state: .unavailable,
        bundledManifestURL: nil,
        installedManifestURL: nil,
        bundledMetadata: nil,
        installedMetadata: nil,
        requiresWordRestart: false,
        lastError: nil
    )

    var isInstalledCurrent: Bool {
        state == .installedCurrent
    }

    var statusLabel: String {
        switch state {
        case .unavailable:
            return "Install unavailable"
        case .notInstalled:
            return "Not installed in Word"
        case .installedCurrent:
            return requiresWordRestart ? "Installed, restart Word" : "Installed in Word"
        case .installedOutdated:
            return requiresWordRestart ? "Installed, restart required" : "Installed, repair recommended"
        case .installFailed:
            return "Install needs attention"
        }
    }

    var statusSummary: String {
        switch state {
        case .unavailable:
            return "The helper could not find the bundled production manifest or the Word sideload folder."
        case .notInstalled:
            return "The production add-in is not installed in Word yet."
        case .installedCurrent:
            return requiresWordRestart
                ? "The production add-in is installed, but Word should be restarted before the helper reopens the taskpane."
                : "The bundled production add-in is installed and ready for Word to open."
        case .installedOutdated:
            return requiresWordRestart
                ? "Word still has an older add-in manifest. Reinstall, then restart Word before reopening the taskpane."
                : "Word has an older add-in manifest. Reinstall to bring it back to the current bundled production manifest."
        case .installFailed:
            return lastError ?? "The helper could not verify or repair the production add-in manifest."
        }
    }
}

struct WordInstallOperationResult: Equatable {
    let status: WordInstallStatus
    let changedInstall: Bool
    let removedManifestCount: Int
}

enum WordInstallService {
    static let wordBundleIdentifier = "com.microsoft.Word"
    static let defaultManifestDirectory = FileManager.default.homeDirectoryForCurrentUser
        .appending(path: "Library/Containers/com.microsoft.Word/Data/Documents/wef", directoryHint: .isDirectory)

    static func isWordInstalled(workspace: NSWorkspace = .shared) -> Bool {
        workspace.urlForApplication(withBundleIdentifier: wordBundleIdentifier) != nil
    }

    static func recommendedAction(
        for status: WordInstallStatus,
        wordIsRunning: Bool
    ) -> WordInstallRecommendation {
        switch status.state {
        case .installedCurrent:
            return status.requiresWordRestart ? .restartWordThenOpen : .openOnly
        case .notInstalled:
            return wordIsRunning ? .restartWordThenOpen : .installThenOpen
        case .installedOutdated, .installFailed:
            return wordIsRunning ? .restartWordThenOpen : .reinstallThenOpen
        case .unavailable:
            return .installThenOpen
        }
    }

    static func evaluateInstallStatus(
        bundledManifestURL: URL?,
        manifestDirectoryURL: URL = defaultManifestDirectory,
        wordIsRunning: Bool
    ) -> WordInstallStatus {
        guard let bundledManifestURL else {
            return WordInstallStatus(
                state: .unavailable,
                bundledManifestURL: nil,
                installedManifestURL: nil,
                bundledMetadata: nil,
                installedMetadata: nil,
                requiresWordRestart: false,
                lastError: "The bundled production manifest is missing."
            )
        }

        do {
            let bundledMetadata = try parseManifestMetadata(from: bundledManifestURL)
            let installedManifestURL = manifestDirectoryURL
                .appendingPathComponent("\(bundledMetadata.addinID).manifest.xml", isDirectory: false)

            guard FileManager.default.fileExists(atPath: installedManifestURL.path) else {
                return WordInstallStatus(
                    state: .notInstalled,
                    bundledManifestURL: bundledManifestURL,
                    installedManifestURL: installedManifestURL,
                    bundledMetadata: bundledMetadata,
                    installedMetadata: nil,
                    requiresWordRestart: false,
                    lastError: nil
                )
            }

            let installedMetadata = try parseManifestMetadata(from: installedManifestURL)
            let matchesCurrent =
                installedMetadata.addinID == bundledMetadata.addinID &&
                installedMetadata.rawHash == bundledMetadata.rawHash &&
                installedMetadata.sourceLocation == bundledMetadata.sourceLocation

            return WordInstallStatus(
                state: matchesCurrent ? .installedCurrent : .installedOutdated,
                bundledManifestURL: bundledManifestURL,
                installedManifestURL: installedManifestURL,
                bundledMetadata: bundledMetadata,
                installedMetadata: installedMetadata,
                requiresWordRestart: !matchesCurrent && wordIsRunning,
                lastError: nil
            )
        } catch {
            return WordInstallStatus(
                state: .installFailed,
                bundledManifestURL: bundledManifestURL,
                installedManifestURL: nil,
                bundledMetadata: nil,
                installedMetadata: nil,
                requiresWordRestart: false,
                lastError: error.localizedDescription
            )
        }
    }

    static func installBundledManifest(
        bundledManifestURL: URL?,
        manifestDirectoryURL: URL = defaultManifestDirectory,
        wordIsRunning: Bool,
        fileManager: FileManager = .default
    ) throws -> WordInstallOperationResult {
        guard let bundledManifestURL else {
            throw NSError(
                domain: "WordInstallService",
                code: 1,
                userInfo: [NSLocalizedDescriptionKey: "The bundled production manifest is missing."]
            )
        }

        let bundledMetadata = try parseManifestMetadata(from: bundledManifestURL)
        let targetURL = manifestDirectoryURL
            .appendingPathComponent("\(bundledMetadata.addinID).manifest.xml", isDirectory: false)

        try fileManager.createDirectory(at: manifestDirectoryURL, withIntermediateDirectories: true)

        var removedCount = 0
        let existingManifestURLs = try fileManager.contentsOfDirectory(
            at: manifestDirectoryURL,
            includingPropertiesForKeys: nil
        ).filter { $0.lastPathComponent.hasSuffix(".manifest.xml") }

        for manifestURL in existingManifestURLs {
            if manifestURL.path == targetURL.path {
                continue
            }
            guard let existingMetadata = try? parseManifestMetadata(from: manifestURL),
                  existingMetadata.addinID == bundledMetadata.addinID else {
                continue
            }
            try fileManager.removeItem(at: manifestURL)
            removedCount += 1
        }

        let previousStatus = evaluateInstallStatus(
            bundledManifestURL: bundledManifestURL,
            manifestDirectoryURL: manifestDirectoryURL,
            wordIsRunning: wordIsRunning
        )
        let changedInstall = previousStatus.state != .installedCurrent

        if fileManager.fileExists(atPath: targetURL.path) {
            try fileManager.removeItem(at: targetURL)
        }
        try fileManager.copyItem(at: bundledManifestURL, to: targetURL)

        let installedMetadata = try parseManifestMetadata(from: targetURL)
        let status = WordInstallStatus(
            state: .installedCurrent,
            bundledManifestURL: bundledManifestURL,
            installedManifestURL: targetURL,
            bundledMetadata: bundledMetadata,
            installedMetadata: installedMetadata,
            requiresWordRestart: changedInstall && wordIsRunning,
            lastError: nil
        )

        return WordInstallOperationResult(
            status: status,
            changedInstall: changedInstall,
            removedManifestCount: removedCount
        )
    }

    static func parseManifestMetadata(from url: URL) throws -> WordManifestMetadata {
        let data = try Data(contentsOf: url)
        return try parseManifestMetadata(from: data)
    }

    static func parseManifestMetadata(from data: Data) throws -> WordManifestMetadata {
        guard let xml = String(data: data, encoding: .utf8) else {
            throw NSError(
                domain: "WordInstallService",
                code: 2,
                userInfo: [NSLocalizedDescriptionKey: "Manifest file is not valid UTF-8."]
            )
        }

        guard let addinID = firstMatch(in: xml, pattern: #"<Id>\s*([^<]+)\s*</Id>"#)?
            .trimmingCharacters(in: .whitespacesAndNewlines),
            !addinID.isEmpty else {
            throw NSError(
                domain: "WordInstallService",
                code: 3,
                userInfo: [NSLocalizedDescriptionKey: "Manifest is missing an add-in ID."]
            )
        }

        let version = firstMatch(in: xml, pattern: #"<Version>\s*([^<]+)\s*</Version>"#)?
            .trimmingCharacters(in: .whitespacesAndNewlines)

        let sourceLocation =
            firstMatch(
                in: xml,
                pattern: #"<bt:Url[^>]*id="Taskpane\.Url"[^>]*DefaultValue="([^"]+)""#
            ) ??
            firstMatch(
                in: xml,
                pattern: #"<SourceLocation[^>]*DefaultValue="([^"]+)""#
            )

        let hash = SHA256.hash(data: data)
            .compactMap { String(format: "%02x", $0) }
            .joined()

        return WordManifestMetadata(
            addinID: addinID,
            version: version,
            sourceLocation: sourceLocation,
            rawHash: hash
        )
    }

    private static func firstMatch(in text: String, pattern: String) -> String? {
        guard let regex = try? NSRegularExpression(pattern: pattern, options: [.caseInsensitive]) else {
            return nil
        }
        let range = NSRange(text.startIndex..<text.endIndex, in: text)
        guard let match = regex.firstMatch(in: text, options: [], range: range),
              match.numberOfRanges > 1,
              let matchRange = Range(match.range(at: 1), in: text) else {
            return nil
        }
        return String(text[matchRange])
    }
}
