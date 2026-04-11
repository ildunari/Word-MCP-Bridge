import XCTest
@testable import WordMCPBridgeHelper

final class BridgeModelsDecodingTests: XCTestCase {
    func testDecodesBridgeStatusPayloadFromLiveShape() throws {
        let json = """
        {
          "ok": true,
          "status": {
            "startedAt": 1774997880076,
            "uptimeMs": 2872,
            "host": "localhost",
            "port": 4017,
            "httpUrl": "https://localhost:4017",
            "wsUrl": "wss://localhost:4017/ws",
            "sessionCount": 1,
            "sessions": [
              {
                "snapshot": {
                  "sessionId": "session-1",
                  "instanceId": "instance-1",
                  "app": "word",
                  "appName": "Microsoft Word",
                  "documentId": "doc-123",
                  "documentMetadata": {
                    "title": "Quarterly Report",
                    "documentId": "doc-123"
                  },
                  "gateway": {
                    "capabilities": ["observe", "tool_call"]
                  }
                },
                "connectedAt": 1774997880000,
                "lastSeenAt": 1774997881000,
                "recentEvents": [],
                "pendingCount": 0,
                "metrics": {
                  "eventCount": 3,
                  "sessionUpdateCount": 1,
                  "toolCallCount": 2,
                  "toolErrorCount": 0,
                  "bridgeWarningCount": 0,
                  "bridgeErrorCount": 0,
                  "unsafeOfficeJsCount": 0,
                  "requestTimeoutCount": 0,
                  "connectionDropCount": 0,
                  "lastEventAt": 1774997881000
                }
              }
            ],
            "totals": {
              "eventCount": 3,
              "sessionUpdateCount": 1,
              "toolCallCount": 2,
              "toolErrorCount": 0,
              "bridgeWarningCount": 0,
              "bridgeErrorCount": 0,
              "unsafeOfficeJsCount": 0,
              "requestTimeoutCount": 0,
              "connectionDropCount": 0,
              "lastEventAt": 1774997881000,
              "pendingCount": 0,
              "connectedSessionCount": 1,
              "disconnectedSessionCount": 0
            }
          }
        }
        """

        let payload = try JSONDecoder().decode(
            BridgeServerStatusResponse.self,
            from: Data(json.utf8)
        )

        XCTAssertTrue(payload.ok)
        XCTAssertEqual(payload.status.host, "localhost")
        XCTAssertEqual(payload.status.port, 4017)
        XCTAssertEqual(payload.status.sessionCount, 1)
        XCTAssertEqual(payload.status.sessions.first?.sessionId, "session-1")
        XCTAssertEqual(payload.status.sessions.first?.metadata.title, "Quarterly Report")
        XCTAssertEqual(payload.status.sessions.first?.documentLabel, "Quarterly Report")
        XCTAssertEqual(payload.status.sessions.first?.documentSummary, "Document ID doc-123")
        XCTAssertEqual(payload.status.sessions.first?.metrics.toolCallCount, 2)
        XCTAssertEqual(payload.status.totals.connectedSessionCount, 1)
    }

    func testDecodesHiddenSharedRuntimeMetadataFromLiveShape() throws {
        let json = """
        {
          "ok": true,
          "status": {
            "startedAt": 1774997880076,
            "uptimeMs": 2872,
            "host": "localhost",
            "port": 4017,
            "httpUrl": "https://localhost:4017",
            "wsUrl": "wss://localhost:4017/ws",
            "sessionCount": 1,
            "sessions": [
              {
                "snapshot": {
                  "sessionId": "session-1",
                  "instanceId": "instance-1",
                  "app": "word",
                  "appName": "Microsoft Word",
                  "documentId": "doc-123",
                  "documentMetadata": {
                    "title": "Quarterly Report",
                    "documentId": "doc-123"
                  },
                  "runtimeState": {
                    "paneVisibility": "hidden",
                    "startupBehavior": "load"
                  },
                  "gateway": {
                    "capabilities": ["observe", "tool_call"]
                  }
                },
                "connectedAt": 1774997880000,
                "lastSeenAt": 1774997881000,
                "recentEvents": [],
                "pendingCount": 0,
                "metrics": {
                  "eventCount": 3,
                  "sessionUpdateCount": 1,
                  "toolCallCount": 2,
                  "toolErrorCount": 0,
                  "bridgeWarningCount": 0,
                  "bridgeErrorCount": 0,
                  "unsafeOfficeJsCount": 0,
                  "requestTimeoutCount": 0,
                  "connectionDropCount": 0,
                  "lastEventAt": 1774997881000
                }
              }
            ],
            "totals": {
              "eventCount": 3,
              "sessionUpdateCount": 1,
              "toolCallCount": 2,
              "toolErrorCount": 0,
              "bridgeWarningCount": 0,
              "bridgeErrorCount": 0,
              "unsafeOfficeJsCount": 0,
              "requestTimeoutCount": 0,
              "connectionDropCount": 0,
              "lastEventAt": 1774997881000,
              "pendingCount": 0,
              "connectedSessionCount": 1,
              "disconnectedSessionCount": 0
            }
          }
        }
        """

        let payload = try JSONDecoder().decode(
            BridgeServerStatusResponse.self,
            from: Data(json.utf8)
        )

        XCTAssertEqual(payload.status.sessions.first?.runtimeVisibilityMode, "hidden")
        XCTAssertTrue(payload.status.sessions.first?.isHiddenSharedRuntimeSession == true)
        XCTAssertTrue(payload.status.sessions.first?.startsAutomatically == true)
    }
}
