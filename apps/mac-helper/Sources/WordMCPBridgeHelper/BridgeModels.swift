import Foundation

struct BridgeServerStatusResponse: Decodable {
    let ok: Bool
    let status: BridgeServerStatusPayload
}

struct BridgeServerStatusPayload: Decodable {
    let running: Bool
    let startedAt: String
    let uptimeMs: Int
    let sessionCount: Int
    let sessions: [BridgeSessionPayload]
    let totals: BridgeMetricsPayload
}

struct BridgeSessionPayload: Decodable, Identifiable {
    let sessionId: String
    let app: String
    let connectedAt: String
    let lastSeenAt: String
    let capabilities: [String]
    let metadata: BridgeMetadataPayload?
    let metrics: BridgeMetricsPayload

    var id: String { sessionId }
}

struct BridgeMetadataPayload: Decodable {
    let documentId: String?
    let title: String?
}

struct BridgeMetricsPayload: Decodable {
    let eventCount: Int
    let toolCallCount: Int
    let bridgeErrorCount: Int
    let requestTimeoutCount: Int
    let connectionDropCount: Int
    let disconnectedSessionCount: Int
    let pendingCount: Int
}

struct HelperSnapshot {
    let fetchedAt: Date
    let status: BridgeServerStatusPayload
}

enum HelperInstallFlow: String, CaseIterable, Identifiable {
    case hosted
    case localDev

    var id: String { rawValue }

    var title: String {
        switch self {
        case .hosted:
            return "Hosted add-in"
        case .localDev:
            return "Local dev"
        }
    }

    var summary: String {
        switch self {
        case .hosted:
            return "Recommended for normal use. Install the hosted Word add-in once, then mostly just open the helper app."
        case .localDev:
            return "For developers working from this repo with the local manifest and dev server."
        }
    }
}

struct HelperAssetAvailability {
    let hostedManifestURL: URL?
    let localManifestURL: URL?
    let setupGuideURL: URL?

    static let unavailable = HelperAssetAvailability(
        hostedManifestURL: nil,
        localManifestURL: nil,
        setupGuideURL: nil
    )

    var hasHostedManifest: Bool { hostedManifestURL != nil }
    var hasLocalManifest: Bool { localManifestURL != nil }
    var hasSetupGuide: Bool { setupGuideURL != nil }
}

struct HelperSetupState {
    let bridgeReachable: Bool
    let connectedSessionCount: Int
    let assetAvailability: HelperAssetAvailability

    var hasWordSession: Bool {
        connectedSessionCount > 0
    }

    var isReady: Bool {
        bridgeReachable && hasWordSession
    }

    var currentStepLabel: String {
        if !assetAvailability.hasHostedManifest {
            return "Install assets missing"
        }
        if !bridgeReachable {
            return "Start the bridge"
        }
        if !hasWordSession {
            return "Open Word and the taskpane"
        }
        return "Ready to use"
    }

    var currentStepSummary: String {
        if !assetAvailability.hasHostedManifest {
            return "The helper could not find the hosted add-in manifest yet."
        }
        if !bridgeReachable {
            return "Start the local bridge so Word can connect to it."
        }
        if !hasWordSession {
            return "The bridge is up. Next, open Word and launch the Word MCP Bridge taskpane."
        }
        return "A live Word session is connected. Your CLI and MCP hosts can attach now."
    }
}
