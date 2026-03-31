// swift-tools-version: 6.0
import PackageDescription

let package = Package(
    name: "WordMCPBridgeHelper",
    platforms: [
        .macOS(.v14),
    ],
    products: [
        .executable(name: "WordMCPBridgeHelper", targets: ["WordMCPBridgeHelper"]),
    ],
    targets: [
        .executableTarget(
            name: "WordMCPBridgeHelper",
            path: "Sources/WordMCPBridgeHelper"
        ),
    ]
)
