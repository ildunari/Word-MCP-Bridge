import Foundation
import ServiceManagement
import SwiftUI

enum HelperPreferences {
    static let autoStartBridgeOnLaunchKey = "autoStartBridgeOnLaunch"
    static let notificationsEnabledKey = "notificationsEnabled"
    static let hasCompletedOnboardingKey = "hasCompletedOnboarding"
    static let dismissSetupBannerKey = "dismissSetupBanner"
}

enum LaunchAtLoginManager {
    static func isEnabled() -> Bool {
        SMAppService.mainApp.status == .enabled
    }

    static func setEnabled(_ enabled: Bool) throws {
        if enabled {
            try SMAppService.mainApp.register()
        } else {
            try SMAppService.mainApp.unregister()
        }
    }
}

struct SettingsView: View {
    @AppStorage(HelperPreferences.autoStartBridgeOnLaunchKey) private var autoStartBridgeOnLaunch = true
    @AppStorage(HelperPreferences.notificationsEnabledKey) private var notificationsEnabled = true
    @State private var launchAtLoginEnabled = LaunchAtLoginManager.isEnabled()
    @State private var launchAtLoginError: String?

    var body: some View {
        Form {
            Section("Startup") {
                Toggle("Launch helper at login", isOn: launchAtLoginBinding)
                Text("Starts the menu bar helper when you sign in so the bridge is one click away.")
                    .font(.caption)
                    .foregroundStyle(.secondary)

                Toggle("Auto-start bridge when helper opens", isOn: $autoStartBridgeOnLaunch)
                Text("When enabled, opening the helper also starts the local bridge automatically.")
                    .font(.caption)
                    .foregroundStyle(.secondary)
            }

            Section("Notifications") {
                Toggle("Bridge health notifications", isOn: $notificationsEnabled)
                Text("Show macOS notifications when Word connects, disconnects, or the bridge stops unexpectedly.")
                    .font(.caption)
                    .foregroundStyle(.secondary)
            }

            Section("Notes") {
                Text("Hosted installs: add the add-in once from the hosted manifest, then use Word as usual.")
                Text("Local dev: use the helper panel’s Actions section when you explicitly want to sideload from this repo.")
            }

            if let launchAtLoginError {
                Section("Launch at Login Error") {
                    Text(launchAtLoginError)
                        .foregroundStyle(.red)
                    Button("Open Login Items Settings") {
                        if let url = URL(string: "x-apple.systempreferences:com.apple.LoginItems-Settings.extension") {
                            NSWorkspace.shared.open(url)
                        }
                    }
                }
            }
        }
        .formStyle(.grouped)
        .padding(20)
        .frame(width: 460)
    }

    private var launchAtLoginBinding: Binding<Bool> {
        Binding(
            get: { launchAtLoginEnabled },
            set: { newValue in
                do {
                    try LaunchAtLoginManager.setEnabled(newValue)
                    launchAtLoginEnabled = LaunchAtLoginManager.isEnabled()
                    launchAtLoginError = nil
                } catch {
                    launchAtLoginEnabled = LaunchAtLoginManager.isEnabled()
                    launchAtLoginError = error.localizedDescription
                }
            }
        )
    }
}
