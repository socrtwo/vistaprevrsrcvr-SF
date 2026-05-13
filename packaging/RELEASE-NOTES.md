# VistaPrevRsrcvr — Multi-platform release

Modernized release of the Previous Version File Recoverer. The build now
targets **.NET Framework 4.8** and ships with launchers for every standard
platform.

## Downloads

| Platform   | File |
| ---------- | ---- |
| Windows    | `VistaPrevRsrcvr-<version>-windows-x64.zip` |
| Linux      | `VistaPrevRsrcvr-<version>-linux.zip` |
| macOS      | `VistaPrevRsrcvr-<version>-macos.zip` |
| ChromeOS   | `VistaPrevRsrcvr-<version>-chromeos.zip` |
| Web (PWA)  | `VistaPrevRsrcvr-<version>-web.zip` |
| Android    | `VistaPrevRsrcvr-<version>-android.zip` |
| iOS        | `VistaPrevRsrcvr-<version>-ios.zip` |
| Source     | `VistaPrevRsrcvr-<version>-source.zip` |

## What changed

- Project upgraded from **.NET Framework 2.0 → 4.8**
- AssemblyInfo bumped to **3.0.0**
- New multi-platform packaging workflow
- Web release ships as an installable PWA (also covers Android / iOS / ChromeOS)
- Unix bundles include Mono-based launchers for Linux, macOS and ChromeOS
- Live page deploys automatically from the `web/` folder

See `PLATFORM-NOTES.md` inside each bundle for setup instructions.
