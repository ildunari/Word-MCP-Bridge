import Foundation
import ServiceManagement
import SwiftUI

enum HelperPreferences {
    static let autoStartBridgeOnLaunchKey = "autoStartBridgeOnLaunch"
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

            Section("Notes") {
                Text("Word still needs the Word MCP Bridge add-in installed once inside Word.")
                Text("After that, your normal flow can be: open the helper, open Word, open the taskpane.")
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
