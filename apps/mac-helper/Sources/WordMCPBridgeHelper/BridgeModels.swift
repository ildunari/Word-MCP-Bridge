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
