# CLAUDE.md

Recovers previous file versions from Windows Shadow Copies (Vista, 7, 8,
10, 11) — including Home editions that lack the built-in "Previous
Versions" feature. **VSS automation is Windows-only.** Non-Windows
platforms run the same UI under Mono and are useful for browsing
shadow-copy snapshots from mounted Windows volumes. A cross-platform PWA
under `web/` provides a browser-based UI for the read-only browse case.

## Repo map

- `Excel Recovery/` — the C# / VB.NET WinForms source (despite the name —
  it's the Previous Version Explorer app).
- `Previous Version Explorer.sln` — Visual Studio solution; .NET
  Framework 4.8 on Windows, Mono on macOS / Linux / ChromeOS.
- `web/` — PWA implementation. Powers the live page and the cross-platform
  bundles (Web / Android / iOS / Linux / macOS / ChromeOS).
- `packaging/` — release packaging helpers and `PLATFORM-NOTES.md` (per
  bundle setup details — `run-linux.sh`, `run-macos.command`,
  `run-chromeos.sh`).
- `releases/` — pre-packaged release archives committed to the repo.
- `.github/workflows/` — `build.yml` (CI), `pages.yml` (deploy `web/` to
  Pages on push to `main`), `release.yml` (build per-platform bundles on
  `v*` tag).

## Branch policy

Work on the assigned feature branch:

1. Commit and push the feature branch.
2. **Open a PR from the feature branch to `main`** using the GitHub MCP
   tools (`mcp__github__create_pull_request`). Do not merge directly —
   the maintainer reviews and merges.
3. CI runs on the PR; Pages and Release pipelines fire from `main` only.

## Releasing

- Push a `v*` tag to `main` to produce: `…-windows-x64.zip`
  (.NET Framework 4.8), `…-macos.zip` (Mono), `…-linux.zip` (Mono),
  `…-chromeos.zip` (Crostini + Mono), `…-web.zip` plus a `…-android.zip`
  and `…-ios.zip` (both PWA wrappers), and `…-source.zip`.

## Verifying changes

- Windows / .NET 4.8: open `Previous Version Explorer.sln` in Visual
  Studio. CI on `build.yml` validates this build.
- Mono: `mono ExcelRecovery.exe` (or equivalent built binary) on
  macOS / Linux — confirm the UI renders and shadow-copy browsing works
  on mounted Windows volumes.
- PWA: serve `web/` locally and exercise the browse UI.

## Gotchas

- VSS automation (the actually-recover-a-file path) is **Windows-only**.
  Don't add a "recover" code path to the Mono / PWA builds; they are
  browse-only by design.
- Mono support means avoiding .NET APIs that Mono doesn't ship — stick
  to .NET Framework 4.8 surface area, don't reach for .NET 6+ APIs.
- The directory is called `Excel Recovery/` for historical reasons —
  this repo is *not* Excel Recovery (see `corruptexcelrec-SF` for that).
  Don't rename it without coordinating: the `.sln` and CI both reference
  the path.
- `.NET Framework 4.8` is the target. Don't upgrade to .NET 6/7/8
  without explicit instruction — the Mono cross-platform story depends
  on the older framework.
