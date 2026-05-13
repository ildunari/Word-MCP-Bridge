# Word-MCP-Bridge — Claude Project Instructions

Project-scoped guidance. Inherits from `~/.claude/CLAUDE.md` and `~/LocalDev/CLAUDE.md`.

## Active profile

**iOS / macOS / Swift dev**. See [.claude/.active-profile](.active-profile).

## What this repo is

Swift Package providing an MCP bridge to Word. Top-level `Package.swift`.

Cross-agent operational notes live in [AGENTS.md](../AGENTS.md).

## Build / test commands

```bash
swift build
swift test
```

## Working defaults

- Worktrees go in `~/LocalDev/.worktrees/Word-MCP-Bridge/<branch>/` per LocalDev root rules — never inside this repo.
- Prefer the `xcodebuildmcp` / `ios-simulator` MCPs for build/run/UI work over ad-hoc `xcodebuild` shells when available.
- Generated Xcode projects (`*.xcodeproj`) regenerate from `project.yml` via XcodeGen where applicable — don't hand-edit pbxproj.
- Never commit API keys, signing certs, or App Store Connect issuer IDs.
