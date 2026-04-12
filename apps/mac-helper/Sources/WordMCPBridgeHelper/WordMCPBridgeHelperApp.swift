import SwiftUI

@main
struct WordMCPBridgeHelperApp: App {
    @StateObject private var controller: BridgeController

    init() {
        let controller = BridgeController()
        controller.startMonitoring()
        _controller = StateObject(wrappedValue: controller)
    }

    var body: some Scene {
        MenuBarExtra(
            "Word MCP Bridge",
            systemImage: menuBarSymbol
        ) {
            HelperMenuView()
                .environmentObject(controller)
                .frame(width: 400, height: 520)
        }
        .menuBarExtraStyle(.window)

        Settings {
            SettingsView()
        }
    }

    private var menuBarSymbol: String {
        if controller.isStarting || controller.isStopping || controller.isTaskpaneServerStarting || controller.isWordAddinStarting {
            return "arrow.triangle.2.circlepath.circle"
        }
        return controller.isBridgeRunning ? "wave.3.right.circle.fill" : "wave.3.right.circle"
    }
}

private enum HelperPanelSection: String, CaseIterable, Identifiable {
    case overview
    case setup
    case actions

    var id: String { rawValue }

    var title: String {
        switch self {
        case .overview:
            return "Overview"
        case .setup:
            return "Setup"
        case .actions:
            return "Actions"
        }
    }
}

private struct HelperMenuView: View {
    @EnvironmentObject private var controller: BridgeController
    @AppStorage(HelperPreferences.hasCompletedOnboardingKey) private var hasCompletedOnboarding = false
    @AppStorage(HelperPreferences.dismissSetupBannerKey) private var dismissSetupBanner = false
    @State private var selectedSection: HelperPanelSection = .overview
    @State private var didCopyConfig = false
    @State private var showsAdvancedActions = false

    private let metricColumns = [
        GridItem(.flexible(), spacing: 8),
        GridItem(.flexible(), spacing: 8),
    ]

    var body: some View {
        VStack(alignment: .leading, spacing: 0) {
            ScrollView {
                VStack(alignment: .leading, spacing: 12) {
                    headerBlock
                    if showsSetupBanner {
                        setupBanner
                    }
                    Picker("Section", selection: $selectedSection) {
                        ForEach(HelperPanelSection.allCases) { section in
                            Text(section.title).tag(section)
                        }
                    }
                    .pickerStyle(.segmented)

                    segmentContent
                }
                .padding(14)
            }
            .frame(maxWidth: .infinity, maxHeight: .infinity, alignment: .top)
            primaryActionBlock
                .padding(.horizontal, 14)
                .padding(.bottom, 14)
                .background(.regularMaterial)
        }
        .frame(maxWidth: .infinity, maxHeight: .infinity, alignment: .top)
        .onChange(of: controller.setupState.isReady) { _, isReady in
            if isReady {
                dismissSetupBanner = false
            }
        }
    }

    private var headerBlock: some View {
        VStack(alignment: .leading, spacing: 8) {
            HStack(alignment: .center, spacing: 10) {
                VStack(alignment: .leading, spacing: 2) {
                    Text("Word MCP Bridge")
                        .font(.headline)
                    Text(controller.setupState.currentStepLabel)
                        .font(.caption)
                        .foregroundStyle(.secondary)
                }
                Spacer()
                Button {
                    Task { await controller.refresh() }
                } label: {
                    Image(systemName: "arrow.clockwise")
                }
                .buttonStyle(.plain)
                .accessibilityLabel("Refresh bridge status")
                .help("Refresh bridge status")
                .disabled(controller.isLoading || controller.isStarting || controller.isWordAddinStarting)

                Menu {
                    Button("Show Setup") {
                        selectedSection = .setup
                    }
                    Button(didCopyConfig ? "Copied MCP Config" : "Copy MCP Config") {
                        copyConfig()
                    }
                    Divider()
                    Button("Open README") {
                        controller.openRepoReadme()
                    }
                    Button("Open Add-in Folder") {
                        controller.openManifestFolder()
                    }
                    SettingsLink()
                } label: {
                    Image(systemName: "ellipsis.circle")
                }
                .menuStyle(.borderlessButton)
                .fixedSize()
                .accessibilityLabel("More actions")
                .help("More actions")
            }

            HStack(spacing: 8) {
                statusBadge(
                    installStatusLabel,
                    systemImage: installStatusIcon,
                    color: controller.setupState.installReady ? .green : .secondary
                )
                statusBadge(
                    taskpaneStatusLabel,
                    systemImage: taskpaneStatusIcon,
                    color: controller.setupState.taskpaneServerReachable ? .green : .secondary
                )
                statusBadge(
                    bridgeStatusLabel,
                    systemImage: bridgeStatusIcon,
                    color: controller.setupState.bridgeReachable ? .green : .secondary
                )
                if controller.setupState.hasWordSession {
                    statusBadge(
                        "Word connected",
                        systemImage: "doc.text.fill",
                        color: .green
                    )
                } else if controller.isWordRunning {
                    statusBadge(
                        "Word open",
                        systemImage: "doc.text",
                        color: .secondary
                    )
                }
            }

            Text(controller.setupState.currentStepSummary)
                .font(.caption)
                .foregroundStyle(.secondary)
                .fixedSize(horizontal: false, vertical: true)
                .frame(maxWidth: .infinity, alignment: .leading)

            if let lastError = controller.lastError {
                errorLine(lastError)
            }
            if let taskpaneErr = controller.taskpaneServerLastError {
                errorLine(taskpaneErr)
            }
            if let wordErr = controller.wordAddinLastError {
                errorLine(wordErr)
            }
            if let installErr = controller.wordInstallLastError {
                errorLine(installErr)
            }
        }
    }

    private var setupBanner: some View {
        HStack(alignment: .top, spacing: 10) {
            Image(systemName: "checklist")
                .foregroundStyle(Color.accentColor)
                .frame(width: 18)
                .padding(.top, 2)
            VStack(alignment: .leading, spacing: 4) {
                Text("Finish setup in this panel")
                    .font(.subheadline.weight(.semibold))
                Text("Use the Setup section to install the add-in into Word, repair it if needed, and confirm the helper is connected once Word opens the panel.")
                    .font(.caption)
                    .foregroundStyle(.secondary)
                    .fixedSize(horizontal: false, vertical: true)
            }
            Spacer(minLength: 0)
            Button("Open") {
                selectedSection = .setup
            }
            .buttonStyle(.bordered)
            .controlSize(.small)

            Button {
                dismissSetupBanner = true
            } label: {
                Image(systemName: "xmark")
            }
            .buttonStyle(.borderless)
            .foregroundStyle(.secondary)
            .accessibilityLabel("Dismiss setup banner")
        }
        .padding(10)
        .background(
            RoundedRectangle(cornerRadius: 10, style: .continuous)
                .fill(Color.accentColor.opacity(0.08))
        )
    }

    private var showsSetupBanner: Bool {
        !hasCompletedOnboarding && !controller.setupState.isReady && !dismissSetupBanner
    }

    @ViewBuilder
    private var segmentContent: some View {
        switch selectedSection {
        case .overview:
            overviewSection
        case .setup:
            setupSection
        case .actions:
            actionsSection
        }
    }

    private var overviewSection: some View {
        VStack(alignment: .leading, spacing: 10) {
            compactSection("Bridge") {
                if let snapshot = controller.snapshot {
                    infoRow("Sessions", value: "\(snapshot.status.sessionCount)")
                    infoRow("Uptime", value: formatDuration(snapshot.status.uptimeMs))
                    infoRow("Endpoint", value: "https://localhost:4017")
                } else if controller.isLoading {
                    HStack(spacing: 8) {
                        ProgressView()
                            .controlSize(.small)
                        Text("Checking bridge status…")
                            .font(.caption)
                            .foregroundStyle(.secondary)
                    }
                } else {
                    Text("No live bridge snapshot yet.")
                        .font(.caption)
                        .foregroundStyle(.secondary)
                }
            }

            compactSection("Readiness") {
                readinessLine("Word install ready", ready: controller.setupState.installReady)
                readinessLine("Local panel server ready", ready: controller.setupState.taskpaneServerReachable)
                readinessLine("Bridge reachable", ready: controller.setupState.bridgeReachable)
                readinessLine("Word taskpane connected", ready: controller.setupState.hasWordSession)
            }

            compactSection("Why isn't this ready?") {
                Text(controller.setupState.whyNotReadyExplanation)
                    .font(.caption)
                    .foregroundStyle(.secondary)
                    .fixedSize(horizontal: false, vertical: true)
            }

            compactSection("Sessions") {
                if let sessions = controller.snapshot?.status.sessions, !sessions.isEmpty {
                    ForEach(sessions.prefix(2)) { session in
                        VStack(alignment: .leading, spacing: 2) {
                            Text(session.documentLabel)
                                .font(.caption.weight(.medium))
                                .fixedSize(horizontal: false, vertical: true)
                                .frame(maxWidth: .infinity, alignment: .leading)
                            Text("\(session.app) • \(session.documentSummary) • \(session.metrics.toolCallCount) tool calls")
                                .font(.caption2)
                                .foregroundStyle(.secondary)
                        }
                        if session.id != sessions.prefix(2).last?.id {
                            Divider()
                        }
                    }
                } else {
                    Text(noSessionSummary)
                        .font(.caption)
                        .foregroundStyle(.secondary)
                        .fixedSize(horizontal: false, vertical: true)
                }
            }

            if let totals = controller.snapshot?.status.totals {
                compactSection("Live totals") {
                    LazyVGrid(columns: metricColumns, alignment: .leading, spacing: 8) {
                        metricTile("Tool calls", value: totals.toolCallCount)
                        metricTile("Errors", value: totals.bridgeErrorCount)
                        metricTile("Timeouts", value: totals.requestTimeoutCount)
                        metricTile("Dropped", value: totals.connectionDropCount)
                        metricTile("Pending", value: totals.pendingCount ?? 0)
                    }
                }
            }
        }
    }

    private var setupSection: some View {
        VStack(alignment: .leading, spacing: 10) {
            compactSection("Setup checklist") {
                readinessLine(
                    "Production manifest bundled",
                    ready: controller.setupState.assetAvailability.hasProductionManifest,
                    summary: controller.setupState.assetAvailability.hasProductionManifest
                        ? "The helper can install the bundled production manifest into Word."
                        : "The local production manifest is missing from the app bundle or repo."
                )
                readinessLine(
                    "Word install ready",
                    ready: controller.setupState.installReady,
                    summary: controller.setupState.wordInstallStatus.statusSummary
                )
                readinessLine(
                    "Local taskpane server running",
                    ready: controller.setupState.taskpaneServerReachable,
                    summary: controller.setupState.taskpaneServerReachable
                        ? "Word can load the taskpane from this Mac."
                        : "Keep the helper open so it can serve the local taskpane before Word opens the panel."
                )
                readinessLine(
                    "Bridge running",
                    ready: controller.setupState.bridgeReachable,
                    summary: controller.setupState.bridgeReachable
                        ? "The bridge is reachable."
                        : "Start the bridge before checking Word."
                )
                readinessLine(
                    "Word taskpane connected",
                    ready: controller.setupState.hasWordSession,
                    summary: controller.setupState.hasWordSession
                        ? "A live Word session is connected."
                        : noSessionSummary
                )
            }

            compactSection("Install flow") {
                Text(controller.setupState.wordInstallStatus.statusSummary)
                    .font(.caption)
                    .foregroundStyle(.secondary)
                    .fixedSize(horizontal: false, vertical: true)

                actionPair(
                    primaryTitle: "Install in Word",
                    primaryAction: controller.installInWord,
                    secondaryTitle: "Open Word MCP Bridge",
                    secondaryAction: controller.openWordMcpBridge,
                    primaryDisabled: controller.isInstallingWordAddin,
                    secondaryDisabled: controller.isOpeningWordTaskpane
                )
                actionPair(
                    primaryTitle: "Repair Word Install",
                    primaryAction: controller.repairWordInstall,
                    secondaryTitle: "Reinstall in Word",
                    secondaryAction: controller.reinstallInWord,
                    primaryDisabled: controller.isInstallingWordAddin,
                    secondaryDisabled: controller.isInstallingWordAddin
                )
                actionPair(
                    primaryTitle: "Open Word",
                    primaryAction: controller.openWord,
                    secondaryTitle: "Copy MCP Config",
                    secondaryAction: copyConfig
                )

                actionPair(
                    primaryTitle: "Open Setup Guide",
                    primaryAction: controller.openSetupGuide,
                    secondaryTitle: "Reveal Production Manifest",
                    secondaryAction: controller.openProductionManifest
                )
            }

            compactSection("Completion") {
                if hasCompletedOnboarding {
                    Text("Setup guidance is marked complete. You can still review this section any time.")
                        .font(.caption)
                        .foregroundStyle(.secondary)
                        .fixedSize(horizontal: false, vertical: true)
                    Button("Reset Setup Checklist") {
                        controller.resetOnboarding()
                    }
                    .buttonStyle(.bordered)
                } else {
                    Text("Mark setup complete once the local panel server, the bridge, and the Word taskpane are all connected.")
                        .font(.caption)
                        .foregroundStyle(.secondary)
                        .fixedSize(horizontal: false, vertical: true)
                    Button("Mark Setup Complete") {
                        controller.completeOnboarding()
                        hasCompletedOnboarding = true
                        selectedSection = .overview
                    }
                    .buttonStyle(.borderedProminent)
                    .disabled(!controller.setupState.isReady)
                }
            }
        }
    }

    private var actionsSection: some View {
        VStack(alignment: .leading, spacing: 10) {
            compactSection("Word") {
                actionPair(
                    primaryTitle: "Open Word MCP Bridge",
                    primaryAction: controller.openWordMcpBridge,
                    secondaryTitle: "Open Word",
                    secondaryAction: controller.openWord,
                    primaryDisabled: controller.isOpeningWordTaskpane,
                    secondaryDisabled: false
                )
                actionPair(
                    primaryTitle: "Install in Word",
                    primaryAction: controller.installInWord,
                    secondaryTitle: "Repair Word Install",
                    secondaryAction: controller.repairWordInstall,
                    primaryDisabled: controller.isInstallingWordAddin,
                    secondaryDisabled: controller.isInstallingWordAddin
                )
            }

            compactSection("Resources") {
                actionPair(
                    primaryTitle: "Copy MCP Config",
                    primaryAction: copyConfig,
                    secondaryTitle: "Setup Guide",
                    secondaryAction: controller.openSetupGuide
                )
                actionPair(
                    primaryTitle: "Production Manifest",
                    primaryAction: controller.openProductionManifest,
                    secondaryTitle: "Add-in Folder",
                    secondaryAction: controller.openManifestFolder
                )
                actionPair(
                    primaryTitle: "README",
                    primaryAction: controller.openRepoReadme,
                    secondaryTitle: "Dev Manifest",
                    secondaryAction: controller.openDevelopmentManifest
                )
            }

            compactSection("Advanced") {
                DisclosureGroup(
                    isExpanded: $showsAdvancedActions,
                    content: {
                        VStack(alignment: .leading, spacing: 8) {
                            statusLine(
                                wordAddinDevStatus,
                                systemImage: wordAddinDevIcon,
                                color: controller.isWordAddinDevRunning ? .green : .secondary,
                                font: .caption
                            )
                            actionPair(
                                primaryTitle: wordAddinLoadTitle,
                                primaryAction: controller.startWordAddinDevSession,
                                secondaryTitle: "Stop Dev Add-in",
                                secondaryAction: controller.stopWordAddinDevSession,
                                primaryDisabled: controller.isWordAddinStarting || controller.isWordAddinDevRunning,
                                secondaryDisabled: !controller.isWordAddinDevRunning
                            )
                            actionPair(
                                primaryTitle: "Dev Manifest",
                                primaryAction: controller.openDevelopmentManifest,
                                secondaryTitle: "Add-in Folder",
                                secondaryAction: controller.openManifestFolder
                            )
                        }
                        .padding(.top, 8)
                    },
                    label: {
                        VStack(alignment: .leading, spacing: 2) {
                            Text("Developer tools")
                                .font(.caption.weight(.semibold))
                            Text("Sideload the local dev add-in only when you are actively working from this repo.")
                                .font(.caption2)
                                .foregroundStyle(.secondary)
                                .fixedSize(horizontal: false, vertical: true)
                        }
                    }
                )
                .disclosureGroupStyle(.automatic)
            }
        }
    }

    private var primaryActionBlock: some View {
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

            HStack(spacing: 8) {
                SettingsLink {
                    Label("Settings", systemImage: "gearshape")
                }
                .buttonStyle(.bordered)
                .controlSize(.small)

                Spacer(minLength: 0)

                Button("Quit") {
                    NSApplication.shared.terminate(nil)
                }
                .buttonStyle(.bordered)
                .controlSize(.small)
            }
        }
    }

    private func compactSection<Content: View>(
        _ title: String,
        @ViewBuilder content: () -> Content
    ) -> some View {
        VStack(alignment: .leading, spacing: 8) {
            Text(title)
                .font(.subheadline.weight(.semibold))
            content()
        }
        .padding(10)
        .frame(maxWidth: .infinity, alignment: .leading)
        .background(
            RoundedRectangle(cornerRadius: 10, style: .continuous)
                .fill(Color(nsColor: .controlBackgroundColor))
        )
    }

    private func infoRow(_ title: String, value: String) -> some View {
        HStack(alignment: .firstTextBaseline, spacing: 8) {
            Text(title)
                .font(.caption)
                .foregroundStyle(.secondary)
            Spacer(minLength: 8)
            Text(value)
                .font(.caption)
                .multilineTextAlignment(.trailing)
                .fixedSize(horizontal: false, vertical: true)
        }
    }

    private func metricTile(_ title: String, value: Int) -> some View {
        VStack(alignment: .leading, spacing: 2) {
            Text(title)
                .font(.caption2)
                .foregroundStyle(.secondary)
            Text("\(value)")
                .font(.caption.monospacedDigit().weight(.semibold))
        }
        .frame(maxWidth: .infinity, alignment: .leading)
        .padding(8)
        .background(
            RoundedRectangle(cornerRadius: 8, style: .continuous)
                .fill(Color(nsColor: .windowBackgroundColor))
        )
    }

    private func readinessLine(_ title: String, ready: Bool, summary: String? = nil) -> some View {
        HStack(alignment: .top, spacing: 8) {
            Image(systemName: ready ? "checkmark.circle.fill" : "circle.dashed")
                .foregroundStyle(ready ? .green : .secondary)
                .frame(width: 16)
                .padding(.top, 1)
            VStack(alignment: .leading, spacing: 2) {
                Text(title)
                    .font(.caption.weight(.medium))
                if let summary {
                    Text(summary)
                        .font(.caption2)
                        .foregroundStyle(.secondary)
                        .fixedSize(horizontal: false, vertical: true)
                }
            }
        }
    }

    private func actionPair(
        primaryTitle: String,
        primaryAction: @escaping () -> Void,
        secondaryTitle: String,
        secondaryAction: @escaping () -> Void,
        primaryDisabled: Bool = false,
        secondaryDisabled: Bool = false
    ) -> some View {
        HStack(spacing: 8) {
            Button(primaryTitle, action: primaryAction)
                .buttonStyle(.borderedProminent)
                .controlSize(.small)
                .disabled(primaryDisabled)
            Button(secondaryTitle, action: secondaryAction)
                .buttonStyle(.bordered)
                .controlSize(.small)
                .disabled(secondaryDisabled)
        }
    }

    private func statusBadge(
        _ text: String,
        systemImage: String,
        color: Color
    ) -> some View {
        HStack(spacing: 6) {
            Image(systemName: systemImage)
            Text(text)
        }
        .font(.caption)
        .foregroundStyle(color)
        .padding(.horizontal, 8)
        .padding(.vertical, 4)
        .background(
            Capsule(style: .continuous)
                .fill(color.opacity(0.12))
        )
    }

    private func errorLine(_ text: String) -> some View {
        Text(text)
            .font(.caption)
            .foregroundStyle(.red)
            .fixedSize(horizontal: false, vertical: true)
            .frame(maxWidth: .infinity, alignment: .leading)
    }

    private func copyConfig() {
        controller.copyMcpConfig()
        didCopyConfig = true
        Task {
            try? await Task.sleep(for: .seconds(1.5))
            didCopyConfig = false
        }
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
            return "Bridge reachable"
        }
        if controller.setupState.bridgeStarting {
            return "Starting bridge"
        }
        if controller.setupState.bridgeProcessRunning {
            return "Bridge process running"
        }
        return "Bridge offline"
    }

    private var installStatusLabel: String {
        controller.setupState.wordInstallStatus.statusLabel
    }

    private var taskpaneStatusLabel: String {
        if controller.setupState.taskpaneServerReachable {
            return "Local panel ready"
        }
        if controller.setupState.taskpaneServerStarting {
            return "Starting local panel"
        }
        if controller.setupState.taskpaneServerProcessRunning {
            return "Local panel starting"
        }
        return "Local panel offline"
    }

    private var taskpaneStatusIcon: String {
        if controller.setupState.taskpaneServerReachable {
            return "rectangle.connected.to.line.below"
        }
        if controller.setupState.taskpaneServerStarting || controller.setupState.taskpaneServerProcessRunning {
            return "arrow.triangle.2.circlepath"
        }
        return "rectangle.slash"
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

    private var installStatusIcon: String {
        switch controller.setupState.wordInstallStatus.state {
        case .installedCurrent:
            return controller.setupState.wordInstallStatus.requiresWordRestart
                ? "arrow.clockwise.circle"
                : "checkmark.seal.fill"
        case .installedOutdated:
            return "wrench.and.screwdriver"
        case .notInstalled:
            return "square.and.arrow.down"
        case .installFailed:
            return "exclamationmark.triangle"
        case .unavailable:
            return "questionmark.circle"
        }
    }

    private var wordAddinDevStatus: String {
        if controller.isWordAddinStarting {
            return "Starting Word dev add-in…"
        }
        if controller.isWordAddinDevRunning {
            return "Word dev add-in session is running (sideload + dev server)."
        }
        if controller.isWordRunning {
            return "Word is open. Open the Word MCP Bridge taskpane if you want a live session."
        }
        return "Word is not open. Open Word when you want to connect the add-in."
    }

    private var wordAddinDevIcon: String {
        if controller.isWordAddinStarting {
            return "arrow.triangle.2.circlepath"
        }
        return controller.isWordAddinDevRunning ? "puzzlepiece.extension.fill" : "puzzlepiece.extension"
    }

    private var wordAddinLoadTitle: String {
        if controller.isWordAddinStarting {
            return "Starting…"
        }
        if controller.isWordAddinDevRunning {
            return "Dev add-in running"
        }
        return "Load Dev Add-in"
    }

    private var noSessionSummary: String {
        if !controller.setupState.installReady {
            return controller.setupState.wordInstallStatus.statusSummary
        }
        if !controller.setupState.taskpaneServerReachable {
            return "Keep the helper open so Word can load the local taskpane before connecting to the bridge."
        }
        if !controller.setupState.bridgeReachable {
            return "Start the bridge first, then open Word and the taskpane."
        }
        if controller.isWordRunning {
            return "Word is open, but the Word MCP Bridge taskpane is not connected yet."
        }
        return "Word is not open yet. Open Word, then open the Word MCP Bridge taskpane."
    }

    private func statusLine(
        _ text: String,
        systemImage: String,
        color: Color,
        font: Font = .body
    ) -> some View {
        HStack(alignment: .top, spacing: 8) {
            Image(systemName: systemImage)
                .foregroundStyle(color)
                .padding(.top, 2)
                .frame(width: 16)
            Text(text)
                .font(font)
                .foregroundStyle(color)
                .multilineTextAlignment(.leading)
                .layoutPriority(1)
                .fixedSize(horizontal: false, vertical: true)
                .frame(minWidth: 0, maxWidth: .infinity, alignment: .leading)
        }
        .accessibilityElement(children: .combine)
        .accessibilityLabel(text)
    }
}
