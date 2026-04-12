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
    let connectedAt: Int64
    let lastSeenAt: Int64
    let pendingCount: Int
    let metrics: BridgeMetricsPayload

    var sessionId: String { snapshot.sessionId }
    var app: String { snapshot.appName ?? snapshot.app }
    var capabilities: [String] { snapshot.gateway?.capabilities ?? [] }
    var metadata: BridgeMetadataPayload {
        snapshot.documentMetadata ?? BridgeMetadataPayload(
            documentId: snapshot.documentId,
            title: snapshot.appName,
            url: nil,
            visibilityMode: nil,
            paneVisibility: nil,
            taskpaneVisibility: nil,
            startupBehavior: nil,
            startupBehaviorEnabled: nil
        )
    }
    var id: String { sessionId }
    var runtimeVisibilityMode: String? {
        normalizedVisibilityMode(
            snapshot.runtimeState?.visibilityMode,
            snapshot.runtimeState?.paneVisibility,
            snapshot.runtimeState?.taskpaneVisibility,
            metadata.visibilityMode,
            metadata.paneVisibility,
            metadata.taskpaneVisibility
        )
    }
    var isHiddenSharedRuntimeSession: Bool {
        guard let runtimeVisibilityMode else { return false }
        return ["hidden", "background", "background_active", "collapsed"]
            .contains(runtimeVisibilityMode)
    }
    var startsAutomatically: Bool {
        snapshot.runtimeState?.startupBehaviorEnabled == true ||
            metadata.startupBehaviorEnabled == true ||
            snapshot.runtimeState?.startupBehavior?.lowercased() == "load" ||
            metadata.startupBehavior?.lowercased() == "load"
    }

    var documentLabel: String {
        if let title = normalizedTitle(metadata.title) {
            return title
        }
        if isUnsavedWordDocument {
            return "Untitled Word document"
        }
        if let documentId = normalizedText(snapshot.documentId),
           looksLikeFilename(documentId) {
            return basename(documentId)
        }
        return "\(app) document"
    }

    var documentSummary: String {
        if isUnsavedWordDocument {
            return "Unsaved local document"
        }
        if let documentId = normalizedText(snapshot.documentId),
           looksLikeFilename(documentId) {
            return "Saved document"
        }
        if let documentId = normalizedText(snapshot.documentId) {
            return "Document ID \(documentId)"
        }
        return "Live document"
    }

    private var isUnsavedWordDocument: Bool {
        snapshot.app == "word" && (snapshot.documentId?.hasPrefix("word-local:") ?? false)
    }
}

struct BridgeMetadataPayload: Decodable {
    let documentId: String?
    let title: String?
    let url: String?
    let visibilityMode: String?
    let paneVisibility: String?
    let taskpaneVisibility: String?
    let startupBehavior: String?
    let startupBehaviorEnabled: Bool?

    init(
        documentId: String?,
        title: String?,
        url: String?,
        visibilityMode: String? = nil,
        paneVisibility: String? = nil,
        taskpaneVisibility: String? = nil,
        startupBehavior: String? = nil,
        startupBehaviorEnabled: Bool? = nil
    ) {
        self.documentId = documentId
        self.title = title
        self.url = url
        self.visibilityMode = visibilityMode
        self.paneVisibility = paneVisibility
        self.taskpaneVisibility = taskpaneVisibility
        self.startupBehavior = startupBehavior
        self.startupBehaviorEnabled = startupBehaviorEnabled
    }
}

struct BridgeSessionSnapshotPayload: Decodable {
    let sessionId: String
    let instanceId: String?
    let app: String
    let appName: String?
    let documentId: String?
    let documentMetadata: BridgeMetadataPayload?
    let runtimeState: BridgeRuntimeStatePayload?
    let gateway: BridgeGatewayPayload?

    init(
        sessionId: String,
        instanceId: String?,
        app: String,
        appName: String?,
        documentId: String?,
        documentMetadata: BridgeMetadataPayload?,
        runtimeState: BridgeRuntimeStatePayload? = nil,
        gateway: BridgeGatewayPayload?
    ) {
        self.sessionId = sessionId
        self.instanceId = instanceId
        self.app = app
        self.appName = appName
        self.documentId = documentId
        self.documentMetadata = documentMetadata
        self.runtimeState = runtimeState
        self.gateway = gateway
    }
}

struct BridgeRuntimeStatePayload: Decodable {
    let visibilityMode: String?
    let paneVisibility: String?
    let taskpaneVisibility: String?
    let startupBehavior: String?
    let startupBehaviorEnabled: Bool?

    init(
        visibilityMode: String? = nil,
        paneVisibility: String? = nil,
        taskpaneVisibility: String? = nil,
        startupBehavior: String? = nil,
        startupBehaviorEnabled: Bool? = nil
    ) {
        self.visibilityMode = visibilityMode
        self.paneVisibility = paneVisibility
        self.taskpaneVisibility = taskpaneVisibility
        self.startupBehavior = startupBehavior
        self.startupBehaviorEnabled = startupBehaviorEnabled
    }
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

private func normalizedText(_ value: String?) -> String? {
    guard let value else { return nil }
    let trimmed = value.trimmingCharacters(in: .whitespacesAndNewlines)
    return trimmed.isEmpty ? nil : trimmed
}

private func normalizedTitle(_ value: String?) -> String? {
    guard let title = normalizedText(value) else { return nil }
    if ["word mcp bridge", "microsoft word", "word"].contains(title.lowercased()) {
        return nil
    }
    return title
}

private func looksLikeFilename(_ value: String) -> Bool {
    value.contains("/") || value.contains("\\") || value.contains(".")
}

private func basename(_ value: String) -> String {
    value.split(whereSeparator: { $0 == "/" || $0 == "\\" }).last.map(String.init) ?? value
}

struct HelperAssetAvailability {
    let productionManifestURL: URL?
    let developmentManifestURL: URL?
    let setupGuideURL: URL?

    static let unavailable = HelperAssetAvailability(
        productionManifestURL: nil,
        developmentManifestURL: nil,
        setupGuideURL: nil
    )

    var hasProductionManifest: Bool { productionManifestURL != nil }
    var hasDevelopmentManifest: Bool { developmentManifestURL != nil }
    var hasSetupGuide: Bool { setupGuideURL != nil }
}

struct HelperSetupState {
    let wordInstallStatus: WordInstallStatus
    let taskpaneServerReachable: Bool
    let taskpaneServerProcessRunning: Bool
    let taskpaneServerStarting: Bool
    let bridgeReachable: Bool
    let bridgeProcessRunning: Bool
    let bridgeStarting: Bool
    let wordAppRunning: Bool
    let connectedSessionCount: Int
    let assetAvailability: HelperAssetAvailability

    var hasWordSession: Bool {
        connectedSessionCount > 0
    }

    var isReady: Bool {
        wordInstallStatus.isInstalledCurrent && taskpaneServerReachable && bridgeReachable && hasWordSession
    }

    var installReady: Bool {
        wordInstallStatus.isInstalledCurrent
    }

    var whyNotReadyExplanation: String {
        if !wordInstallStatus.isInstalledCurrent {
            return wordInstallStatus.statusSummary
        }
        if !taskpaneServerReachable {
            return "The local taskpane page is not reachable, so Word has nothing valid to load in the sidebar yet."
        }
        if !bridgeReachable {
            return "The local bridge is not reachable, so the taskpane cannot attach back to the helper."
        }
        if !wordAppRunning {
            return "Microsoft Word is not open yet, so the add-in cannot connect."
        }
        if !hasWordSession {
            return "The add-in is installed and the local services are healthy, but the Word MCP Bridge panel is not open in Word yet."
        }
        return "Everything the helper needs is up and connected."
    }

    var currentStepLabel: String {
        if !assetAvailability.hasProductionManifest {
            return "Install assets missing"
        }
        if !wordInstallStatus.isInstalledCurrent {
            return wordInstallStatus.requiresWordRestart ? "Restart Word" : "Install in Word"
        }
        if taskpaneServerStarting {
            return "Starting local panel"
        }
        if !taskpaneServerReachable {
            return taskpaneServerProcessRunning ? "Waiting for local panel" : "Start the helper"
        }
        if bridgeStarting {
            return "Starting the bridge"
        }
        if !bridgeReachable {
            return bridgeProcessRunning ? "Waiting for the bridge" : "Start the bridge"
        }
        if !hasWordSession {
            return wordAppRunning ? "Open the taskpane" : "Open Word"
        }
        return "Ready to use"
    }

    var currentStepSummary: String {
        if !assetAvailability.hasProductionManifest {
            return "The helper could not find the local production add-in manifest yet."
        }
        if !wordInstallStatus.isInstalledCurrent {
            return wordInstallStatus.statusSummary
        }
        if taskpaneServerStarting {
            return "The helper is starting its local taskpane server so Word can load the side panel from this machine."
        }
        if !taskpaneServerReachable {
            return taskpaneServerProcessRunning
                ? "The helper started the local taskpane server, but Word cannot reach it yet."
                : "The helper's local taskpane server is not reachable yet. Reopen the helper before opening the Word MCP Bridge taskpane."
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
            return wordAppRunning
                ? "Word is open, but the Word MCP Bridge taskpane is not connected yet."
                : "The bridge is up. Open Microsoft Word, then open the Word MCP Bridge taskpane."
        }
        return "A live Word session is connected. If shared runtime has hidden the panel, your CLI and MCP hosts can still attach now."
    }
}

private func normalizedVisibilityMode(_ values: String?...) -> String? {
    values
        .lazy
        .compactMap(normalizedText)
        .map { $0.lowercased() }
        .first
}
