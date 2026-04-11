import AppKit
import ApplicationServices
import Foundation

enum Mode: String {
    case open
    case status
    case showAddins = "show-addins"
}

enum ExitCode: Int32 {
    case success = 0
    case bridgePaneNeedsManualOpen = 2
    case missingAccessibility = 3
    case wordNotRunning = 4
    case actionFailed = 5
    case badArguments = 64
}

struct Config {
    let mode: Mode
}

func normalized(_ value: String) -> String {
    value
        .replacingOccurrences(of: "‑", with: "-")
        .replacingOccurrences(of: "–", with: "-")
        .replacingOccurrences(of: "—", with: "-")
        .folding(options: [.caseInsensitive, .diacriticInsensitive], locale: .current)
        .trimmingCharacters(in: .whitespacesAndNewlines)
}

func parseConfig() -> Config {
    var mode: Mode = .open
    var iterator = CommandLine.arguments.dropFirst().makeIterator()
    while let arg = iterator.next() {
        switch arg {
        case "--mode":
            guard let raw = iterator.next(), let parsed = Mode(rawValue: raw) else {
                fputs("Invalid or missing value for --mode.\n", stderr)
                exit(ExitCode.badArguments.rawValue)
            }
            mode = parsed
        case "--help", "-h":
            print("Use the shell wrapper at scripts/bridge/launch-word-taskpane.sh for help.")
            exit(ExitCode.success.rawValue)
        default:
            fputs("Unknown argument: \(arg)\n", stderr)
            exit(ExitCode.badArguments.rawValue)
        }
    }
    return Config(mode: mode)
}

func copyAttribute(_ element: AXUIElement, name: String) -> AnyObject? {
    var value: CFTypeRef?
    let error = AXUIElementCopyAttributeValue(element, name as CFString, &value)
    guard error == .success else { return nil }
    return value
}

func attributeString(_ element: AXUIElement, name: String) -> String? {
    copyAttribute(element, name: name) as? String
}

func attributePoint(_ element: AXUIElement, name: String) -> CGPoint? {
    guard let value = copyAttribute(element, name: name)
    else { return nil }
    let axValue = value as! AXValue
    guard AXValueGetType(axValue) == .cgPoint else { return nil }
    var point = CGPoint.zero
    return AXValueGetValue(axValue, .cgPoint, &point) ? point : nil
}

func attributeSize(_ element: AXUIElement, name: String) -> CGSize? {
    guard let value = copyAttribute(element, name: name)
    else { return nil }
    let axValue = value as! AXValue
    guard AXValueGetType(axValue) == .cgSize else { return nil }
    var size = CGSize.zero
    return AXValueGetValue(axValue, .cgSize, &size) ? size : nil
}

func childElements(of element: AXUIElement) -> [AXUIElement] {
    let arrayAttributeNames = [
        kAXChildrenAttribute as String,
        kAXVisibleChildrenAttribute as String,
        "AXChildrenInNavigationOrder",
        "AXSections",
    ]
    let singleAttributeNames = [
        "AXToolbarButton",
        kAXTitleUIElementAttribute as String,
    ]

    var seen = Set<String>()
    var results: [AXUIElement] = []
    for name in arrayAttributeNames {
        guard let values = copyAttribute(element, name: name) as? [AXUIElement] else {
            continue
        }
        for child in values {
            let identifier = String(describing: child)
            if seen.insert(identifier).inserted {
                results.append(child)
            }
        }
    }
    for name in singleAttributeNames {
        guard let child = copyAttribute(element, name: name) else {
            continue
        }
        let elementChild = child as! AXUIElement
        let identifier = String(describing: elementChild)
        if seen.insert(identifier).inserted {
            results.append(elementChild)
        }
    }
    return results
}

func matchesAnyTitle(_ element: AXUIElement, titles: [String]) -> Bool {
    let candidates = [
        attributeString(element, name: kAXTitleAttribute as String),
        attributeString(element, name: kAXDescriptionAttribute as String),
        attributeString(element, name: kAXRoleDescriptionAttribute as String),
        attributeString(element, name: kAXIdentifierAttribute as String),
    ]
    .compactMap { $0 }
    .map(normalized)

    let normalizedTargets = titles.map(normalized)
    for candidate in candidates {
        if normalizedTargets.contains(candidate) {
            return true
        }
    }
    return false
}

func attributeRole(_ element: AXUIElement) -> String {
    attributeString(element, name: kAXRoleAttribute as String) ?? ""
}

func firstMatch(
    root: AXUIElement,
    titles: [String],
    maxDepth: Int = 12,
    predicate: ((AXUIElement) -> Bool)? = nil
) -> AXUIElement? {
    var queue: [(AXUIElement, Int)] = [(root, 0)]
    var visited = Set<String>()

    while !queue.isEmpty {
        let (element, depth) = queue.removeFirst()
        let identifier = String(describing: element)
        if !visited.insert(identifier).inserted {
            continue
        }

        if matchesAnyTitle(element, titles: titles), predicate?(element) ?? true {
            return element
        }

        guard depth < maxDepth else { continue }
        for child in childElements(of: element) {
            queue.append((child, depth + 1))
        }
    }

    return nil
}

func collectButtonTitles(root: AXUIElement, maxDepth: Int = 10) -> [String] {
    var queue: [(AXUIElement, Int)] = [(root, 0)]
    var visited = Set<String>()
    var titles: [String] = []

    while !queue.isEmpty {
        let (element, depth) = queue.removeFirst()
        let identifier = String(describing: element)
        if !visited.insert(identifier).inserted {
            continue
        }

        let role = attributeString(element, name: kAXRoleAttribute as String) ?? ""
        if role == kAXButtonRole as String,
           let title = attributeString(element, name: kAXTitleAttribute as String),
           !title.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty {
            titles.append(title)
        }

        guard depth < maxDepth else { continue }
        for child in childElements(of: element) {
            queue.append((child, depth + 1))
        }
    }

    return Array(Set(titles)).sorted()
}

func press(_ element: AXUIElement) -> Bool {
    if AXUIElementPerformAction(element, kAXPressAction as CFString) == .success {
        return true
    }
    if AXUIElementPerformAction(element, kAXShowMenuAction as CFString) == .success {
        return true
    }

    let clickPoint =
        attributePoint(element, name: "AXActivationPoint") ??
        {
            guard
                let position = attributePoint(element, name: kAXPositionAttribute as String),
                let size = attributeSize(element, name: kAXSizeAttribute as String)
            else { return nil }
            return CGPoint(x: position.x + (size.width / 2), y: position.y + (size.height / 2))
        }()

    guard let clickPoint else { return false }
    guard
        let mouseDown = CGEvent(mouseEventSource: nil, mouseType: .leftMouseDown, mouseCursorPosition: clickPoint, mouseButton: .left),
        let mouseUp = CGEvent(mouseEventSource: nil, mouseType: .leftMouseUp, mouseCursorPosition: clickPoint, mouseButton: .left)
    else {
        return false
    }

    mouseDown.post(tap: CGEventTapLocation.cghidEventTap)
    mouseUp.post(tap: CGEventTapLocation.cghidEventTap)
    return true
}

func wordApplication() -> NSRunningApplication? {
    let bundleIdentifier = "com.microsoft.Word"
    return NSRunningApplication.runningApplications(withBundleIdentifier: bundleIdentifier).first
}

func activateWord(_ app: NSRunningApplication) {
    app.activate(options: [.activateIgnoringOtherApps])
    usleep(400_000)
}

func wordWindowElement(pid: pid_t) -> AXUIElement? {
    let appElement = AXUIElementCreateApplication(pid)
    if let mainWindow = copyAttribute(appElement, name: kAXMainWindowAttribute as String) {
        return (mainWindow as! AXUIElement)
    }
    if let windows = copyAttribute(appElement, name: kAXWindowsAttribute as String) as? [AXUIElement] {
        return windows.first
    }
    return nil
}

func isPaneOpen(in window: AXUIElement) -> Bool {
    firstMatch(
        root: window,
        titles: ["Word MCP Bridge"],
        predicate: { element in
            let role = attributeRole(element)
            return role != kAXButtonRole as String
        }
    ) != nil
}

func findRibbonButton(in window: AXUIElement, titles: [String]) -> AXUIElement? {
    firstMatch(
        root: window,
        titles: titles,
        predicate: { element in
            attributeRole(element) == kAXButtonRole as String
        }
    )
}

func waitForMatch(
    pid: pid_t,
    titles: [String],
    timeoutSeconds: TimeInterval = 5.0,
    predicate: ((AXUIElement) -> Bool)? = nil
) -> AXUIElement? {
    let deadline = Date().addingTimeInterval(timeoutSeconds)
    while Date() < deadline {
        if let window = wordWindowElement(pid: pid),
           let match = firstMatch(root: window, titles: titles, predicate: predicate) {
            return match
        }
        usleep(350_000)
    }
    return nil
}

let config = parseConfig()

guard AXIsProcessTrusted() else {
    fputs("macOS Accessibility access is required for terminal automation of Word.\n", stderr)
    exit(ExitCode.missingAccessibility.rawValue)
}

guard let word = wordApplication() else {
    fputs("Microsoft Word is not running.\n", stderr)
    exit(ExitCode.wordNotRunning.rawValue)
}

activateWord(word)

guard let window = wordWindowElement(pid: word.processIdentifier) else {
    fputs("Could not find a Word window to inspect.\n", stderr)
    exit(ExitCode.actionFailed.rawValue)
}

if config.mode == .status {
    let paneOpen = isPaneOpen(in: window)
    let ribbonButton = findRibbonButton(in: window, titles: ["Open Word MCP Bridge", "Word MCP Bridge"]) != nil
    print("paneOpen=\(paneOpen) ribbonButton=\(ribbonButton)")
    if !ribbonButton {
        let titles = collectButtonTitles(root: window)
        if !titles.isEmpty {
            print("visibleButtons=\(titles.joined(separator: " | "))")
        }
    }
    exit(ExitCode.success.rawValue)
}

if isPaneOpen(in: window) {
    print("Word MCP Bridge pane is already open.")
    exit(ExitCode.success.rawValue)
}

let addinsTitles = ["Add-ins", "Add‑ins", "My Add-ins"]

if config.mode == .showAddins {
    if let addinsButton = waitForMatch(pid: word.processIdentifier, titles: addinsTitles), press(addinsButton) {
        print("Opened the Word Add-ins surface.")
        exit(ExitCode.success.rawValue)
    }
    fputs("Could not find the Add-ins control in the current Word window.\n", stderr)
    exit(ExitCode.actionFailed.rawValue)
}

let ribbonTargets = ["Open Word MCP Bridge", "Word MCP Bridge"]
if let directButton = waitForMatch(
    pid: word.processIdentifier,
    titles: ribbonTargets,
    predicate: { element in
        attributeRole(element) == kAXButtonRole as String
    }
),
   press(directButton) {
    usleep(700_000)
    if let refreshedWindow = wordWindowElement(pid: word.processIdentifier), isPaneOpen(in: refreshedWindow) {
        print("Opened the Word MCP Bridge taskpane from the ribbon.")
    } else {
        print("Triggered the Word MCP Bridge ribbon command.")
    }
    exit(ExitCode.success.rawValue)
}

if let addinsButton = waitForMatch(pid: word.processIdentifier, titles: addinsTitles), press(addinsButton) {
    print("Opened the Word Add-ins surface. A first-time manual click on the Word MCP Bridge tile may still be required.")
    exit(ExitCode.bridgePaneNeedsManualOpen.rawValue)
}

fputs("Could not find either the Word MCP Bridge ribbon command or the Add-ins control.\n", stderr)
exit(ExitCode.actionFailed.rawValue)
