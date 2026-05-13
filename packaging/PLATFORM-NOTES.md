# Platform notes

VistaPrevRsrcvr is a VB.NET / WinForms application. The full Volume Shadow
Copy Service (VSS) workflow it automates only exists on Windows, but the UI
runs on any platform with a .NET-compatible runtime so you can browse and
recover from shadow-copy snapshots on mounted Windows volumes.

## Windows
Run the included `.exe` directly. Requires .NET Framework 4.8 (preinstalled on
Windows 10 / 11; download from Microsoft for older Windows).

## Linux
1. Install Mono: `sudo apt-get install mono-complete`
2. Run `./run-linux.sh`

## macOS
1. Install Mono (`brew install mono` or download from mono-project.com)
2. Double-click `run-macos.command` (or run it from Terminal)

## ChromeOS
1. Enable the Linux (Crostini) container in Settings > Advanced > Developers
2. Copy this folder into the **Linux files** area
3. From Terminal: `sudo apt-get install mono-complete && ./run-chromeos.sh`

## Android / iOS / Web
Use the Web release (PWA) — install it to the home screen from your browser.
The PWA bundles documentation, recovery references and links into an
offline-capable app. The native VSS automation is not available on mobile, but
you can use the PWA for guidance and to navigate shadow-copy contents on
mounted/synced drives.
