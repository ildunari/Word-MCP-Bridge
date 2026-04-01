import SwiftUI

@main
struct WordMCPBridgeHelperApp: App {
    @StateObject private var controller = BridgeController()

    var body: some Scene {
        MenuBarExtra(
            "Word MCP Bridge",
            systemImage: menuBarSymbol
        ) {
            HelperMenuView()
                .environmentObject(controller)
                .frame(width: 340)
                .onAppear {
                    controller.startMonitoring()
                    controller.showOnboarding()
                }
        }
        .menuBarExtraStyle(.window)

        Settings {
            SettingsView()
        }
    }

    private var menuBarSymbol: String {
        if controller.isStarting || controller.isStopping || controller.isLoading {
            return "arrow.triangle.2.circlepath.circle"
        }
        return controller.isBridgeRunning ? "wave.3.right.circle.fill" : "wave.3.right.circle"
    }
}

private struct HelperMenuView: View {
    @EnvironmentObject private var controller: BridgeController
    @State private var didCopyConfig = false

    var body: some View {
        VStack(alignment: .leading, spacing: 14) {
            statusBlock
            totalsBlock
            sessionsBlock
            quickActionsBlock
            actionBlock
        }
        .padding(16)
    }

    private var statusBlock: some View {
        VStack(alignment: .leading, spacing: 6) {
            HStack(alignment: .center, spacing: 10) {
                Text("Local Bridge")
                    .font(.headline)
                Spacer()
                Button {
                    Task { await controller.refresh() }
                } label: {
                    Image(systemName: "arrow.clockwise")
                }
                .buttonStyle(.plain)
                .help("Refresh bridge status")
                .disabled(controller.isLoading || controller.isStarting)

                Menu {
                    Button(didCopyConfig ? "Copied MCP Config" : "Copy MCP Config") {
                        controller.copyMcpConfig()
                        didCopyConfig = true
                        Task {
                            try? await Task.sleep(for: .seconds(1.5))
                            didCopyConfig = false
                        }
                    }
                    Button("Open README") {
                        controller.openRepoReadme()
                    }
                    Button("Open Add-in Folder") {
                        controller.openManifestFolder()
                    }
                    Button("Getting Started") {
                        controller.showOnboarding(force: true)
                    }
                    SettingsLink()
                    Divider()
                    Button("Quit") {
                        NSApplication.shared.terminate(nil)
                    }
                } label: {
                    Image(systemName: "gearshape")
                }
                .menuStyle(.borderlessButton)
                .fixedSize()
                .help("More actions")
            }

            Label(bridgeStatusLabel, systemImage: bridgeStatusIcon)
                .foregroundStyle(controller.setupState.bridgeReachable ? .green : .secondary)
                .accessibilityLabel(bridgeStatusLabel)
            if let snapshot = controller.snapshot {
                Text("Sessions: \(snapshot.status.sessionCount)  •  Uptime: \(formatDuration(snapshot.status.uptimeMs))")
                    .font(.caption)
                    .foregroundStyle(.secondary)
            } else if controller.isLoading {
                HStack(spacing: 8) {
                    ProgressView()
                        .controlSize(.small)
                    Text("Checking bridge status...")
                        .font(.caption)
                        .foregroundStyle(.secondary)
                }
            } else {
                Text(controller.setupState.currentStepSummary)
                    .font(.caption)
                    .foregroundStyle(.secondary)
            }
            if let lastError = controller.lastError {
                Text(lastError)
                    .font(.caption)
                    .foregroundStyle(.red)
                    .accessibilityLabel("Error: \(lastError)")
            }
        }
    }

    private var totalsBlock: some View {
        VStack(alignment: .leading, spacing: 8) {
            Text("Totals")
                .font(.subheadline.weight(.semibold))
            if let totals = controller.snapshot?.status.totals {
                metricsRow("Tool calls", value: totals.toolCallCount)
                metricsRow("Errors", value: totals.bridgeErrorCount)
                metricsRow("Timeouts", value: totals.requestTimeoutCount)
                metricsRow("Dropped", value: totals.connectionDropCount)
                metricsRow("Pending", value: totals.pendingCount ?? 0)
            } else {
                Text("Start the bridge or connect Word to see live totals.")
                    .font(.caption)
                    .foregroundStyle(.secondary)
            }
        }
    }

    private var sessionsBlock: some View {
        VStack(alignment: .leading, spacing: 8) {
            Text("Connected Sessions")
                .font(.subheadline.weight(.semibold))
            if let sessions = controller.snapshot?.status.sessions, !sessions.isEmpty {
                ForEach(sessions.prefix(3)) { session in
                    VStack(alignment: .leading, spacing: 2) {
                        Text(session.metadata?.title ?? session.metadata?.documentId ?? session.sessionId)
                            .font(.caption.weight(.medium))
                        Text("\(session.app)  •  \(session.metrics.toolCallCount) tool calls")
                            .font(.caption2)
                            .foregroundStyle(.secondary)
                    }
                }
                if sessions.count > 3 {
                    Text("+ \(sessions.count - 3) more sessions")
                        .font(.caption2)
                        .foregroundStyle(.tertiary)
                }
            } else {
                Text("Open the Word add-in task pane to register a session.")
                    .font(.caption)
                    .foregroundStyle(.secondary)
            }
        }
    }

    private var actionBlock: some View {
        VStack(alignment: .leading, spacing: 8) {
            Divider()
            if controller.isBridgeRunning {
                Button(primaryActionTitle) {
                    controller.stopBridge()
                }
                .buttonStyle(.bordered)
                .tint(.red)
                .disabled(controller.isStarting || controller.isStopping)
                .controlSize(.large)
                .frame(maxWidth: .infinity)
                .help("Stop the local bridge")
            } else {
                Button(primaryActionTitle) {
                    controller.startBridge()
                }
                .buttonStyle(.borderedProminent)
                .disabled(controller.isStarting || controller.isStopping)
                .controlSize(.large)
                .frame(maxWidth: .infinity)
                .help("Start the local bridge")
            }
        }
    }

    private var quickActionsBlock: some View {
        VStack(alignment: .leading, spacing: 8) {
            Text("Quick Actions")
                .font(.subheadline.weight(.semibold))

            HStack(spacing: 8) {
                Button("Open Word") {
                    controller.openWord()
                }
                .buttonStyle(.bordered)

                Button("Hosted Manifest") {
                    controller.openHostedManifest()
                }
                .buttonStyle(.bordered)
            }

            HStack(spacing: 8) {
                Button("Setup Guide") {
                    controller.openSetupGuide()
                }
                .buttonStyle(.bordered)

                Button("Getting Started") {
                    controller.showOnboarding(force: true)
                }
                .buttonStyle(.bordered)
            }
        }
    }

    private func metricsRow(_ label: String, value: Int) -> some View {
        LabeledContent {
            Text("\(value)")
                .font(.caption.monospacedDigit())
                .foregroundStyle(.secondary)
        } label: {
            Text(label)
                .font(.caption)
        }
        .accessibilityElement(children: .combine)
    }

    private func formatDuration(_ milliseconds: Int) -> String {
        let totalSeconds = max(milliseconds / 1_000, 0)
        if totalSeconds < 60 {
            return "\(totalSeconds)s"
        }
        let hours = totalSeconds / 3_600
        let minutes = (totalSeconds % 3_600) / 60
        if hours > 0 {
            return "\(hours)h \(minutes)m"
        }
        return "\(minutes)m"
    }

    private var primaryActionTitle: String {
        if controller.isStarting {
            return "Starting..."
        }
        if controller.isStopping {
            return "Stopping..."
        }
        return controller.isBridgeRunning ? "Stop Bridge" : "Start Bridge"
    }

    private var bridgeStatusLabel: String {
        if controller.setupState.bridgeReachable {
            return "Reachable on https://127.0.0.1:4017"
        }
        if controller.setupState.bridgeStarting {
            return "Starting bridge"
        }
        if controller.setupState.bridgeProcessRunning {
            return "Bridge process running, waiting for endpoint"
        }
        return "Bridge not reachable"
    }

    private var bridgeStatusIcon: String {
        if controller.setupState.bridgeReachable {
            return "checkmark.circle.fill"
        }
        if controller.setupState.bridgeStarting || controller.setupState.bridgeProcessRunning {
            return "arrow.triangle.2.circlepath.circle"
        }
        return "xmark.circle"
    }
}
