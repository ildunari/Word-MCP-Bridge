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
            "host": "127.0.0.1",
            "port": 4017,
            "httpUrl": "https://127.0.0.1:4017",
            "wsUrl": "wss://127.0.0.1:4017/ws",
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
        XCTAssertEqual(payload.status.sessions.first?.metadata?.title, "Draft.docx")
        XCTAssertEqual(payload.status.sessions.first?.capabilities, ["observe", "unsafe_office_js"])
    }

    func testSetupStateSeparatesBridgeStartupFromWordSessionReadiness() {
        let state = HelperSetupState(
            bridgeReachable: false,
            bridgeProcessRunning: true,
            bridgeStarting: true,
            connectedSessionCount: 0,
            assetAvailability: .init(
                hostedManifestURL: URL(fileURLWithPath: "/tmp/manifest.xml"),
                localManifestURL: nil,
                setupGuideURL: nil
            )
        )

        XCTAssertEqual(state.currentStepLabel, "Starting the bridge")
        XCTAssertEqual(state.currentStepSummary, "The helper launched the local bridge and is waiting for it to become reachable.")
    }
}
