import Foundation

enum LocalhostCertificateService {
    private static let bundleDirectoryName = "Word MCP Bridge"
    private static let certificateCommonName = "Word MCP Bridge Localhost"

    struct CertificateURLs: Equatable {
        let certificateURL: URL
        let keyURL: URL
        let opensslConfigURL: URL
    }

    static func resolveExistingCertificateURLs(
        fileManager: FileManager = .default,
        homeDirectory: URL? = nil
    ) -> (certURL: URL, keyURL: URL)? {
        let helperURLs = helperCertificateURLs(fileManager: fileManager, homeDirectory: homeDirectory)
        if fileManager.fileExists(atPath: filesystemPath(for: helperURLs.certificateURL)),
           fileManager.fileExists(atPath: filesystemPath(for: helperURLs.keyURL)) {
            return (helperURLs.certificateURL, helperURLs.keyURL)
        }

        let legacyURLs = legacyCertificateURLs(fileManager: fileManager, homeDirectory: homeDirectory)
        if fileManager.fileExists(atPath: filesystemPath(for: legacyURLs.certURL)),
           fileManager.fileExists(atPath: filesystemPath(for: legacyURLs.keyURL)) {
            return legacyURLs
        }

        return nil
    }

    static func ensureReady(
        fileManager: FileManager = .default,
        homeDirectory: URL? = nil,
        trustCertificateHandler: ((URL, URL?) throws -> Void)? = nil
    ) throws -> (certURL: URL, keyURL: URL) {
        let helperURLs = helperCertificateURLs(fileManager: fileManager, homeDirectory: homeDirectory)
        let trustHandler = trustCertificateHandler ?? trustCertificate
        if fileManager.fileExists(atPath: filesystemPath(for: helperURLs.certificateURL)),
           fileManager.fileExists(atPath: filesystemPath(for: helperURLs.keyURL)) {
            try trustHandler(helperURLs.certificateURL, homeDirectory)
            return (helperURLs.certificateURL, helperURLs.keyURL)
        }

        try fileManager.createDirectory(
            at: helperURLs.certificateURL.deletingLastPathComponent(),
            withIntermediateDirectories: true
        )

        let legacyURLs = legacyCertificateURLs(fileManager: fileManager, homeDirectory: homeDirectory)
        if fileManager.fileExists(atPath: filesystemPath(for: legacyURLs.certURL)),
           fileManager.fileExists(atPath: filesystemPath(for: legacyURLs.keyURL)) {
            try copyLegacyCertificate(legacyURLs: legacyURLs, helperURLs: helperURLs, fileManager: fileManager)
            try trustHandler(helperURLs.certificateURL, homeDirectory)
            return (helperURLs.certificateURL, helperURLs.keyURL)
        }

        try writeOpenSSLConfig(to: helperURLs.opensslConfigURL)
        defer { try? fileManager.removeItem(at: helperURLs.opensslConfigURL) }

        try runProcess(
            command: "/usr/bin/openssl",
            arguments: [
                "req",
                "-x509",
                "-nodes",
                "-newkey",
                "rsa:2048",
                "-sha256",
                "-days",
                "825",
                "-subj",
                "/CN=\(certificateCommonName)",
                "-keyout",
                filesystemPath(for: helperURLs.keyURL),
                "-out",
                filesystemPath(for: helperURLs.certificateURL),
                "-config",
                filesystemPath(for: helperURLs.opensslConfigURL),
                "-extensions",
                "v3_req",
            ]
        )
        try trustHandler(helperURLs.certificateURL, homeDirectory)
        return (helperURLs.certificateURL, helperURLs.keyURL)
    }

    static func helperCertificateURLs(fileManager: FileManager = .default, homeDirectory: URL? = nil) -> CertificateURLs {
        let directory = (homeDirectory ?? fileManager.homeDirectoryForCurrentUser)
            .appendingPathComponent("Library", isDirectory: true)
            .appendingPathComponent("Application Support", isDirectory: true)
            .appendingPathComponent(bundleDirectoryName, isDirectory: true)
            .appendingPathComponent("certs", isDirectory: true)
        return CertificateURLs(
            certificateURL: directory.appendingPathComponent("localhost.crt", isDirectory: false),
            keyURL: directory.appendingPathComponent("localhost.key", isDirectory: false),
            opensslConfigURL: directory.appendingPathComponent("openssl-localhost.cnf", isDirectory: false)
        )
    }

    static func loginKeychainURL(fileManager: FileManager = .default, homeDirectory: URL? = nil) -> URL {
        (homeDirectory ?? fileManager.homeDirectoryForCurrentUser)
            .appendingPathComponent("Library", isDirectory: true)
            .appendingPathComponent("Keychains", isDirectory: true)
            .appendingPathComponent("login.keychain-db", isDirectory: false)
    }

    private static func legacyCertificateURLs(
        fileManager: FileManager,
        homeDirectory: URL?
    ) -> (certURL: URL, keyURL: URL) {
        let home = homeDirectory ?? fileManager.homeDirectoryForCurrentUser
        return (
            home.appendingPathComponent(".office-addin-dev-certs", isDirectory: true)
                .appendingPathComponent("localhost.crt", isDirectory: false),
            home.appendingPathComponent(".office-addin-dev-certs", isDirectory: true)
                .appendingPathComponent("localhost.key", isDirectory: false)
        )
    }

    private static func copyLegacyCertificate(
        legacyURLs: (certURL: URL, keyURL: URL),
        helperURLs: CertificateURLs,
        fileManager: FileManager
    ) throws {
        if fileManager.fileExists(atPath: filesystemPath(for: helperURLs.certificateURL)) {
            try fileManager.removeItem(at: helperURLs.certificateURL)
        }
        if fileManager.fileExists(atPath: filesystemPath(for: helperURLs.keyURL)) {
            try fileManager.removeItem(at: helperURLs.keyURL)
        }
        try fileManager.copyItem(at: legacyURLs.certURL, to: helperURLs.certificateURL)
        try fileManager.copyItem(at: legacyURLs.keyURL, to: helperURLs.keyURL)
    }

    private static func writeOpenSSLConfig(to url: URL) throws {
        let config = """
        [req]
        distinguished_name = req_distinguished_name
        x509_extensions = v3_req
        prompt = no

        [req_distinguished_name]
        CN = localhost
        O = Word MCP Bridge

        [v3_req]
        keyUsage = critical, digitalSignature, keyEncipherment
        extendedKeyUsage = serverAuth
        subjectAltName = @alt_names

        [alt_names]
        DNS.1 = localhost
        IP.1 = 127.0.0.1
        """
        try config.write(to: url, atomically: true, encoding: .utf8)
    }

    private static func trustCertificate(at certURL: URL, homeDirectory: URL?) throws {
        try runProcess(
            command: "/usr/bin/security",
            arguments: [
                "add-trusted-cert",
                "-d",
                "-r",
                "trustRoot",
                "-k",
                filesystemPath(for: loginKeychainURL(homeDirectory: homeDirectory)),
                filesystemPath(for: certURL),
            ]
        )
    }

    private static func runProcess(command: String, arguments: [String]) throws {
        let process = Process()
        process.executableURL = URL(fileURLWithPath: command)
        process.arguments = arguments

        let output = Pipe()
        process.standardOutput = output
        process.standardError = output

        try process.run()
        process.waitUntilExit()

        guard process.terminationStatus == 0 else {
            let data = output.fileHandleForReading.readDataToEndOfFile()
            let text = String(decoding: data, as: UTF8.self)
                .trimmingCharacters(in: .whitespacesAndNewlines)
            throw NSError(
                domain: "LocalhostCertificateService",
                code: Int(process.terminationStatus),
                userInfo: [
                    NSLocalizedDescriptionKey: text.isEmpty ? "Command failed with exit code \(process.terminationStatus)." : text,
                ]
            )
        }
    }

    private static func filesystemPath(for url: URL) -> String {
        url.path.removingPercentEncoding ?? url.path
    }
}
