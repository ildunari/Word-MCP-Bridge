import XCTest
@testable import WordMCPBridgeHelper

final class BridgeModelsTests: XCTestCase {
    func testDecodesCurrentBridgeStatusPayloadShape() throws {
        let json = """
        {
          "ok": true,
          "status": {
            "startedAt": 1710000000000,
            "uptimeMs": 42000,
            "host": "localhost",
            "port": 4017,
            "httpUrl": "https://localhost:4017",
            "wsUrl": "wss://localhost:4017/ws",
            "sessionCount": 1,
            "sessions": [
              {
                "snapshot": {
                  "sessionId": "word:doc-1",
                  "app": "word",
                  "documentId": "doc-1",
                  "documentMetadata": {
                    "title": "Draft.docx"
                  },
                  "gateway": {
                    "capabilities": ["observe", "unsafe_office_js"]
                  }
                },
                "connectedAt": 1710000000000,
                "lastSeenAt": 1710000001000,
                "recentEvents": [],
                "pendingCount": 0,
                "metrics": {
                  "eventCount": 3,
                  "toolCallCount": 1,
                  "bridgeErrorCount": 0,
                  "requestTimeoutCount": 0,
                  "connectionDropCount": 0
                }
              }
            ],
            "totals": {
              "eventCount": 3,
              "toolCallCount": 1,
              "bridgeErrorCount": 0,
              "requestTimeoutCount": 0,
              "connectionDropCount": 0,
              "disconnectedSessionCount": 0,
              "pendingCount": 0
            }
          }
        }
        """

        let payload = try JSONDecoder().decode(
            BridgeServerStatusResponse.self,
            from: Data(json.utf8)
        )

        XCTAssertEqual(payload.status.sessionCount, 1)
        XCTAssertEqual(payload.status.sessions.first?.sessionId, "word:doc-1")
        XCTAssertEqual(payload.status.sessions.first?.metadata.title, "Draft.docx")
        XCTAssertEqual(payload.status.sessions.first?.capabilities, ["observe", "unsafe_office_js"])
        XCTAssertEqual(payload.status.sessions.first?.documentLabel, "Draft.docx")
        XCTAssertEqual(payload.status.sessions.first?.documentSummary, "Document ID doc-1")
    }

    func testSetupStateSeparatesBridgeStartupFromWordSessionReadiness() {
        let state = HelperSetupState(
            wordInstallStatus: .installedCurrentForTests,
            taskpaneServerReachable: true,
            taskpaneServerProcessRunning: true,
            taskpaneServerStarting: false,
            bridgeReachable: false,
            bridgeProcessRunning: true,
            bridgeStarting: true,
            wordAppRunning: false,
            connectedSessionCount: 0,
            assetAvailability: .init(
                productionManifestURL: URL(fileURLWithPath: "/tmp/manifest.prod.xml"),
                developmentManifestURL: nil,
                setupGuideURL: nil
            )
        )

        XCTAssertEqual(state.currentStepLabel, "Starting the bridge")
        XCTAssertEqual(state.currentStepSummary, "The helper launched the local bridge and is waiting for it to become reachable.")
    }

    func testSetupStateCallsForWordInstallBeforeAnythingElse() {
        let state = HelperSetupState(
            wordInstallStatus: WordInstallStatus(
                state: .notInstalled,
                bundledManifestURL: URL(fileURLWithPath: "/tmp/manifest.prod.xml"),
                installedManifestURL: nil,
                bundledMetadata: nil,
                installedMetadata: nil,
                requiresWordRestart: false,
                lastError: nil
            ),
            taskpaneServerReachable: true,
            taskpaneServerProcessRunning: true,
            taskpaneServerStarting: false,
            bridgeReachable: true,
            bridgeProcessRunning: true,
            bridgeStarting: false,
            wordAppRunning: true,
            connectedSessionCount: 0,
            assetAvailability: .init(
                productionManifestURL: URL(fileURLWithPath: "/tmp/manifest.prod.xml"),
                developmentManifestURL: nil,
                setupGuideURL: nil
            )
        )

        XCTAssertEqual(state.currentStepLabel, "Install in Word")
        XCTAssertEqual(state.currentStepSummary, "The production add-in is not installed in Word yet.")
    }

    func testSetupStatePromptsToOpenWordWhenBridgeIsReadyButWordIsClosed() {
        let state = HelperSetupState(
            wordInstallStatus: .installedCurrentForTests,
            taskpaneServerReachable: true,
            taskpaneServerProcessRunning: true,
            taskpaneServerStarting: false,
            bridgeReachable: true,
            bridgeProcessRunning: true,
            bridgeStarting: false,
            wordAppRunning: false,
            connectedSessionCount: 0,
            assetAvailability: .init(
                productionManifestURL: URL(fileURLWithPath: "/tmp/manifest.prod.xml"),
                developmentManifestURL: nil,
                setupGuideURL: nil
            )
        )

        XCTAssertEqual(state.currentStepLabel, "Open Word")
        XCTAssertEqual(state.currentStepSummary, "The bridge is up. Open Microsoft Word, then open the Word MCP Bridge taskpane.")
    }

    func testSetupStatePromptsToOpenTaskpaneWhenWordIsRunningWithoutSession() {
        let state = HelperSetupState(
            wordInstallStatus: .installedCurrentForTests,
            taskpaneServerReachable: true,
            taskpaneServerProcessRunning: true,
            taskpaneServerStarting: false,
            bridgeReachable: true,
            bridgeProcessRunning: true,
            bridgeStarting: false,
            wordAppRunning: true,
            connectedSessionCount: 0,
            assetAvailability: .init(
                productionManifestURL: URL(fileURLWithPath: "/tmp/manifest.prod.xml"),
                developmentManifestURL: nil,
                setupGuideURL: nil
            )
        )

        XCTAssertEqual(state.currentStepLabel, "Open the taskpane")
        XCTAssertEqual(state.currentStepSummary, "Word is open, but the Word MCP Bridge taskpane is not connected yet.")
    }

    func testSetupStateConnectedSummaryAllowsHiddenSharedRuntimeSessions() {
        let state = HelperSetupState(
            wordInstallStatus: .installedCurrentForTests,
            taskpaneServerReachable: true,
            taskpaneServerProcessRunning: true,
            taskpaneServerStarting: false,
            bridgeReachable: true,
            bridgeProcessRunning: true,
            bridgeStarting: false,
            wordAppRunning: true,
            connectedSessionCount: 1,
            assetAvailability: .init(
                productionManifestURL: URL(fileURLWithPath: "/tmp/manifest.prod.xml"),
                developmentManifestURL: nil,
                setupGuideURL: nil
            )
        )

        XCTAssertEqual(state.currentStepLabel, "Ready to use")
        XCTAssertEqual(
            state.currentStepSummary,
            "A live Word session is connected. If shared runtime has hidden the panel, your CLI and MCP hosts can still attach now."
        )
    }

    func testSetupStateCallsOutLocalTaskpaneServerBeforeBridge() {
        let state = HelperSetupState(
            wordInstallStatus: .installedCurrentForTests,
            taskpaneServerReachable: false,
            taskpaneServerProcessRunning: false,
            taskpaneServerStarting: false,
            bridgeReachable: true,
            bridgeProcessRunning: true,
            bridgeStarting: false,
            wordAppRunning: true,
            connectedSessionCount: 0,
            assetAvailability: .init(
                productionManifestURL: URL(fileURLWithPath: "/tmp/manifest.prod.xml"),
                developmentManifestURL: nil,
                setupGuideURL: nil
            )
        )

        XCTAssertEqual(state.currentStepLabel, "Start the helper")
        XCTAssertEqual(
            state.currentStepSummary,
            "The helper's local taskpane server is not reachable yet. Reopen the helper before opening the Word MCP Bridge taskpane."
        )
    }

    func testSessionUsesFriendlyLabelForUnsavedWordDocument() {
        let payload = BridgeSessionPayload(
            snapshot: BridgeSessionSnapshotPayload(
                sessionId: "word:local-1",
                instanceId: "instance-1",
                app: "word",
                appName: "Microsoft Word",
                documentId: "word-local:local-1",
                documentMetadata: BridgeMetadataPayload(
                    documentId: "word-local:local-1",
                    title: "Word MCP Bridge",
                    url: nil
                ),
                gateway: nil
            ),
            connectedAt: 1,
            lastSeenAt: 1,
            pendingCount: 0,
            metrics: BridgeMetricsPayload(
                eventCount: 0,
                toolCallCount: 0,
                bridgeErrorCount: 0,
                requestTimeoutCount: 0,
                connectionDropCount: 0,
                disconnectedSessionCount: nil,
                connectedSessionCount: nil,
                pendingCount: nil
            )
        )

        XCTAssertEqual(payload.documentLabel, "Untitled Word document")
        XCTAssertEqual(payload.documentSummary, "Unsaved local document")
    }

    func testSessionRecognizesHiddenSharedRuntimeFromRuntimeState() {
        let payload = BridgeSessionPayload(
            snapshot: BridgeSessionSnapshotPayload(
                sessionId: "word:hidden-1",
                instanceId: "instance-1",
                app: "word",
                appName: "Microsoft Word",
                documentId: "doc-1",
                documentMetadata: BridgeMetadataPayload(
                    documentId: "doc-1",
                    title: "Draft.docx",
                    url: nil
                ),
                runtimeState: BridgeRuntimeStatePayload(
                    visibilityMode: "hidden",
                    paneVisibility: nil,
                    taskpaneVisibility: nil,
                    startupBehavior: "load",
                    startupBehaviorEnabled: true
                ),
                gateway: nil
            ),
            connectedAt: 1,
            lastSeenAt: 1,
            pendingCount: 0,
            metrics: BridgeMetricsPayload(
                eventCount: 0,
                toolCallCount: 0,
                bridgeErrorCount: 0,
                requestTimeoutCount: 0,
                connectionDropCount: 0,
                disconnectedSessionCount: nil,
                connectedSessionCount: nil,
                pendingCount: nil
            )
        )

        XCTAssertEqual(payload.runtimeVisibilityMode, "hidden")
        XCTAssertTrue(payload.isHiddenSharedRuntimeSession)
        XCTAssertTrue(payload.startsAutomatically)
    }
}

private extension WordInstallStatus {
    static let installedCurrentForTests = WordInstallStatus(
        state: .installedCurrent,
        bundledManifestURL: URL(fileURLWithPath: "/tmp/manifest.prod.xml"),
        installedManifestURL: URL(fileURLWithPath: "/tmp/A89087D2-08B4-481F-8E40-7EB65D0966F9.manifest.xml"),
        bundledMetadata: nil,
        installedMetadata: nil,
        requiresWordRestart: false,
        lastError: nil
    )
}
