<!--MODERNIZED:v2-->
# Vistaprevrsrcvr

> Migrated from SourceForge via SF2GH Migrator — modernized to .NET Framework 4.8 with multi-platform release bundles.

[![Build](https://img.shields.io/github/actions/workflow/status/socrtwo/vistaprevrsrcvr-SF/build.yml?style=for-the-badge&color=22c55e&label=build)](https://github.com/socrtwo/vistaprevrsrcvr-SF/actions/workflows/build.yml)
[![Live page](https://img.shields.io/badge/live-page-ff2e93?style=for-the-badge)](https://socrtwo.github.io/vistaprevrsrcvr-SF/)
[![Releases](https://img.shields.io/github/v/release/socrtwo/vistaprevrsrcvr-SF?style=for-the-badge&color=7c3aed)](https://github.com/socrtwo/vistaprevrsrcvr-SF/releases)
[![License](https://img.shields.io/github/license/socrtwo/vistaprevrsrcvr-SF?style=for-the-badge&color=22d3ee)](https://github.com/socrtwo/vistaprevrsrcvr-SF/blob/main/LICENSE)
[![Last commit](https://img.shields.io/github/last-commit/socrtwo/vistaprevrsrcvr-SF?style=for-the-badge&color=34d399)](https://github.com/socrtwo/vistaprevrsrcvr-SF/commits)

🌐 **Live:** https://socrtwo.github.io/vistaprevrsrcvr-SF/
📦 **Downloads:** [Releases](https://github.com/socrtwo/vistaprevrsrcvr-SF/releases)
📂 **Source:** [socrtwo/vistaprevrsrcvr-SF](https://github.com/socrtwo/vistaprevrsrcvr-SF)

---

Recovers previous file versions from Windows Shadow Copies on Vista, 7, 8, 10, and 11 — including Home editions that lack the built-in Previous Versions feature.

## Downloads

Pick a bundle from the [latest release](https://github.com/socrtwo/vistaprevrsrcvr-SF/releases/latest):

| Platform | Bundle | Runtime |
| --- | --- | --- |
| Windows | `…-windows-x64.zip` | .NET Framework 4.8 |
| macOS | `…-macos.zip` | Mono (`brew install mono`) |
| Linux | `…-linux.zip` | Mono (`mono-complete`) |
| ChromeOS | `…-chromeos.zip` | Crostini + Mono |
| Web | `…-web.zip` / [live PWA](https://socrtwo.github.io/vistaprevrsrcvr-SF/) | Any modern browser |
| Android | `…-android.zip` / live PWA | Chrome → Add to Home Screen |
| iOS | `…-ios.zip` / live PWA | Safari → Add to Home Screen |
| Source | `…-source.zip` | — |

Each non-Windows desktop bundle ships with a launcher (`run-linux.sh`, `run-macos.command`, `run-chromeos.sh`) and `PLATFORM-NOTES.md`. See [packaging/PLATFORM-NOTES.md](packaging/PLATFORM-NOTES.md) for setup details.

> **Note:** VSS automation is Windows-only. On macOS / Linux / ChromeOS the UI runs over Mono and is useful for browsing shadow-copy snapshots from mounted Windows volumes.

## Screenshots

Visit the [SourceForge project page](https://sourceforge.net/projects/vistaprevrsrcvr/) to view screenshots.

> **Tip:** If you have screenshots to contribute, open a PR adding them to a `screenshots/` folder!

**Language:** VB.NET
**License:** MIT

## Features

- Accesses Windows Shadow Copy Service (VSS)
- Works on Home editions (which lack the built-in Previous Versions UI)
- Browse and restore previous versions of any file or folder
- Supports Windows Vista, 7, 8, 10, and 11
- Simple file browser interface
- Cross-platform launchers for macOS, Linux and ChromeOS via Mono
- Installable Progressive Web App for Web, Android and iOS

## System Requirements

- **Windows:** Windows 7 or later, .NET Framework 4.8
- **macOS / Linux / ChromeOS:** Mono 6.x or later
- **Mobile / Web:** any modern browser (Chrome, Safari, Edge, Firefox)
- **Build from source:** Visual Studio 2019+ or MSBuild 16+

## Installation & Usage

### Building from Source

1. Open the `.sln` file in Visual Studio
2. Restore NuGet packages if prompted
3. Build the solution (**Build → Build Solution** or `Ctrl+Shift+B`)
4. Find the compiled `.exe` in `bin/Release/`

### Using a Pre-built Release

Download the latest release from the [Releases](../../releases) page and run the `.exe` directly — no install needed.

## Origin

This project was originally hosted on SourceForge and has been migrated to GitHub for easier access and collaboration.

- **SourceForge:** [vistaprevrsrcvr](https://sourceforge.net/projects/vistaprevrsrcvr/)
- **Migrated with:** [SF2GH Migrator](https://github.com/socrtwo/sf-to-github)

## Contributing

Contributions are welcome! Feel free to:

1. Fork this repository
2. Create a feature branch (`git checkout -b my-feature`)
3. Commit your changes (`git commit -m "Add my feature"`)
4. Push to the branch (`git push origin my-feature`)
5. Open a Pull Request

## License

MIT License — see [LICENSE](LICENSE) for details.

---

## 📜 SourceForge heritage

This project originated on **SourceForge** before being migrated to GitHub. The legacy SourceForge entry, if still available, can be searched at:

🔗 https://sourceforge.net/projects/vistaprevrsrcvr/

The repository here at `socrtwo/vistaprevrsrcvr-SF` is the canonical, actively-maintained home. All future updates, issue tracking, and releases happen on GitHub.

## 🛠️ Contributing

Issues and pull requests are welcome at [https://github.com/socrtwo/vistaprevrsrcvr-SF/issues](https://github.com/socrtwo/vistaprevrsrcvr-SF/issues).

## 📝 License

See the [LICENSE](https://github.com/socrtwo/vistaprevrsrcvr-SF/blob/main/LICENSE) file in this repository. If no license file is present, the project is shared as-is for reference and personal use; please contact the maintainer for other use cases.

---

*Maintained by [@socrtwo](https://github.com/socrtwo)*