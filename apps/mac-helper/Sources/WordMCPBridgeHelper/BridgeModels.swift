import Foundation

struct BridgeServerStatusResponse: Decodable {
    let ok: Bool
    let status: BridgeServerStatusPayload
}

struct BridgeServerStatusPayload: Decodable {
    let startedAt: Int
    let uptimeMs: Int
    let host: String
    let port: Int
    let httpUrl: String
    let wsUrl: String
    let sessionCount: Int
    let sessions: [BridgeSessionPayload]
    let totals: BridgeMetricsPayload

    var running: Bool { true }
}

struct BridgeSessionPayload: Decodable, Identifiable {
    let snapshot: BridgeSessionSnapshotPayload
    let connectedAt: Int
    let lastSeenAt: Int
    let pendingCount: Int
    let metrics: BridgeMetricsPayload

    var sessionId: String { snapshot.sessionId }
    var app: String { snapshot.app }
    var capabilities: [String] { snapshot.gateway?.capabilities ?? [] }
    var metadata: BridgeMetadataPayload? {
        snapshot.documentMetadata ?? BridgeMetadataPayload(
            documentId: snapshot.documentId,
            title: nil
        )
    }
    var id: String { sessionId }
}

struct BridgeMetadataPayload: Decodable {
    let documentId: String?
    let title: String?
}

struct BridgeSessionSnapshotPayload: Decodable {
    let sessionId: String
    let app: String
    let documentId: String?
    let documentMetadata: BridgeMetadataPayload?
    let gateway: BridgeGatewayPayload?
}

struct BridgeGatewayPayload: Decodable {
    let capabilities: [String]
}

struct BridgeMetricsPayload: Decodable {
    let eventCount: Int
    let toolCallCount: Int
    let bridgeErrorCount: Int
    let requestTimeoutCount: Int
    let connectionDropCount: Int
    let disconnectedSessionCount: Int?
    let connectedSessionCount: Int?
    let pendingCount: Int?
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
    let bridgeProcessRunning: Bool
    let bridgeStarting: Bool
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
        if bridgeStarting {
            return "Starting the bridge"
        }
        if !bridgeReachable {
            return bridgeProcessRunning ? "Waiting for the bridge" : "Start the bridge"
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
        if bridgeStarting {
            return "The helper launched the local bridge and is waiting for it to become reachable."
        }
        if !bridgeReachable {
            return bridgeProcessRunning
                ? "The bridge process is running, but the local status endpoint is not reachable yet."
                : "Start the local bridge so Word can connect to it."
        }
        if !hasWordSession {
            return "The bridge is up. Next, open Word and launch the Word MCP Bridge taskpane."
        }
        return "A live Word session is connected. Your CLI and MCP hosts can attach now."
    }
}
