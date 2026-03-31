import AppKit
import SwiftUI

@MainActor
final class OnboardingWindowManager: NSObject {
    static let shared = OnboardingWindowManager()

    private var window: NSWindow?

    func show(controller: BridgeController) {
        if let window {
            window.makeKeyAndOrderFront(nil)
            NSApplication.shared.activate(ignoringOtherApps: true)
            return
        }

        let rootView = OnboardingView()
            .environmentObject(controller)

        let hostingController = NSHostingController(rootView: rootView)
        let window = NSWindow(contentViewController: hostingController)
        window.title = "Word MCP Bridge Setup"
        window.styleMask = [.titled, .closable, .miniaturizable, .resizable]
        window.setContentSize(NSSize(width: 760, height: 680))
        window.center()
        window.isReleasedWhenClosed = false
        window.identifier = NSUserInterfaceItemIdentifier("WordMCPBridgeOnboardingWindow")
        window.delegate = self
        window.makeKeyAndOrderFront(nil)
        window.setFrameAutosaveName("WordMCPBridgeOnboardingWindow")

        self.window = window

        NSApplication.shared.activate(ignoringOtherApps: true)
    }

    func close() {
        window?.performClose(nil)
    }
}

extension OnboardingWindowManager: NSWindowDelegate {
    func windowWillClose(_ notification: Notification) {
        window = nil
    }
}

private struct OnboardingView: View {
    @EnvironmentObject private var controller: BridgeController
    @AppStorage(HelperPreferences.hasCompletedOnboardingKey) private var hasCompletedOnboarding = false
    @State private var installFlow: HelperInstallFlow = .hosted

    var body: some View {
        ScrollView {
            VStack(alignment: .leading, spacing: 24) {
                heroSection
                statusSection
                installSection
                agentSection
                dailyUseSection
            }
            .padding(24)
            .frame(maxWidth: .infinity, alignment: .leading)
        }
        .background(Color(nsColor: .windowBackgroundColor))
    }

    private var heroSection: some View {
        VStack(alignment: .leading, spacing: 12) {
            Text("Set up Word MCP Bridge")
                .font(.largeTitle.bold())
            Text("Install the Word add-in once, then let the helper app handle the bridge and your daily workflow.")
                .foregroundStyle(.secondary)

            HStack(spacing: 12) {
                Label(controller.setupState.currentStepLabel, systemImage: controller.setupState.isReady ? "checkmark.circle.fill" : "arrow.triangle.2.circlepath.circle")
                    .foregroundStyle(controller.setupState.isReady ? .green : .primary)
                Spacer()
                Button(hasCompletedOnboarding ? "Setup Complete" : "Mark Setup Complete") {
                    controller.completeOnboarding()
                    hasCompletedOnboarding = true
                    OnboardingWindowManager.shared.close()
                }
                .buttonStyle(.borderedProminent)
                .disabled(!controller.setupState.isReady)
            }

            Text(controller.setupState.currentStepSummary)
                .font(.callout)
                .foregroundStyle(.secondary)
        }
    }

    private var statusSection: some View {
        VStack(alignment: .leading, spacing: 12) {
            Text("Readiness checks")
                .font(.title3.weight(.semibold))
            readinessRow(
                title: "Hosted install assets",
                isReady: controller.setupState.assetAvailability.hasHostedManifest,
                summary: controller.setupState.assetAvailability.hasHostedManifest
                    ? "The hosted manifest is available."
                    : "The helper cannot find the hosted install manifest yet."
            )
            readinessRow(
                title: "Bridge running",
                isReady: controller.setupState.bridgeReachable,
                summary: controller.setupState.bridgeReachable
                    ? "The local bridge is reachable."
                    : "Start the helper bridge before checking Word."
            )
            readinessRow(
                title: "Word taskpane connected",
                isReady: controller.setupState.hasWordSession,
                summary: controller.setupState.hasWordSession
                    ? "A live Word session is connected."
                    : "Open Word and launch the Word MCP Bridge taskpane."
            )
        }
        .padding(18)
        .background(.regularMaterial, in: RoundedRectangle(cornerRadius: 18, style: .continuous))
    }

    private var installSection: some View {
        VStack(alignment: .leading, spacing: 14) {
            Text("Install the Word add-in")
                .font(.title3.weight(.semibold))
            Picker("Install flow", selection: $installFlow) {
                ForEach(HelperInstallFlow.allCases) { flow in
                    Text(flow.title).tag(flow)
                }
            }
            .pickerStyle(.segmented)

            Text(installFlow.summary)
                .foregroundStyle(.secondary)

            if installFlow == .hosted {
                VStack(alignment: .leading, spacing: 10) {
                    Text("Recommended flow")
                        .font(.headline)
                    Text("1. Reveal the hosted manifest.\n2. Open Word.\n3. In Word, use the add-ins install flow and select that manifest.\n4. Open the Word MCP Bridge taskpane.")
                        .foregroundStyle(.secondary)
                    HStack(spacing: 10) {
                        Button("Reveal Hosted Manifest") {
                            controller.openHostedManifest()
                        }
                        .buttonStyle(.borderedProminent)

                        Button("Open Word") {
                            controller.openWord()
                        }
                        .buttonStyle(.bordered)

                        Button("Open Setup Guide") {
                            controller.openSetupGuide()
                        }
                        .buttonStyle(.bordered)
                    }
                }
            } else {
                VStack(alignment: .leading, spacing: 10) {
                    Text("Developer fallback")
                        .font(.headline)
                    Text("Use this when you are actively working from the repo. You still install the add-in in Word, but the local manifest points at the dev server.")
                        .foregroundStyle(.secondary)
                    HStack(spacing: 10) {
                        Button("Reveal Local Manifest") {
                            controller.openLocalManifest()
                        }
                        .buttonStyle(.borderedProminent)

                        Button("Open Word") {
                            controller.openWord()
                        }
                        .buttonStyle(.bordered)
                    }
                }
            }
        }
        .padding(18)
        .background(.regularMaterial, in: RoundedRectangle(cornerRadius: 18, style: .continuous))
    }

    private var agentSection: some View {
        VStack(alignment: .leading, spacing: 12) {
            Text("Connect your agent")
                .font(.title3.weight(.semibold))
            Text("After the bridge and Word taskpane are up, your MCP host can attach to the local stdio server.")
                .foregroundStyle(.secondary)

            HStack(spacing: 10) {
                Button("Copy MCP Config") {
                    controller.copyMcpConfig()
                }
                .buttonStyle(.borderedProminent)

                Button("Open Setup Guide") {
                    controller.openSetupGuide()
                }
                .buttonStyle(.bordered)
            }
        }
        .padding(18)
        .background(.regularMaterial, in: RoundedRectangle(cornerRadius: 18, style: .continuous))
    }

    private var dailyUseSection: some View {
        VStack(alignment: .leading, spacing: 12) {
            Text("Daily workflow")
                .font(.title3.weight(.semibold))
            Text("Once installed, the normal flow is simple:")
                .foregroundStyle(.secondary)

            VStack(alignment: .leading, spacing: 8) {
                Label("Open Word MCP Bridge Helper", systemImage: "1.circle.fill")
                Label("Open Word", systemImage: "2.circle.fill")
                Label("Open the Word MCP Bridge taskpane", systemImage: "3.circle.fill")
                Label("Use Claude, Cursor, Codex, or another MCP-capable host", systemImage: "4.circle.fill")
            }

            HStack(spacing: 10) {
                Button("Start Bridge") {
                    controller.startBridge()
                }
                .buttonStyle(.borderedProminent)
                .disabled(controller.isBridgeRunning || controller.isStarting)

                Button("Refresh Status") {
                    Task { await controller.refresh() }
                }
                .buttonStyle(.bordered)

                Button("Close") {
                    OnboardingWindowManager.shared.close()
                }
                .buttonStyle(.bordered)
            }
        }
        .padding(18)
        .background(.regularMaterial, in: RoundedRectangle(cornerRadius: 18, style: .continuous))
    }

    private func readinessRow(title: String, isReady: Bool, summary: String) -> some View {
        HStack(alignment: .top, spacing: 12) {
            Image(systemName: isReady ? "checkmark.circle.fill" : "circle.dashed")
                .foregroundStyle(isReady ? .green : .secondary)
            VStack(alignment: .leading, spacing: 3) {
                Text(title)
                    .font(.headline)
                Text(summary)
                    .font(.caption)
                    .foregroundStyle(.secondary)
            }
        }
    }
}
