# Previous Version File Recoverer

Recovers previous file versions from Windows Shadow Copies on Vista, 7, and 8 — including Home editions that lack the built-in Previous Versions feature.

## Screenshots

Visit the [SourceForge project page](https://sourceforge.net/projects/vistaprevrsrcvr/) to view screenshots.

> **Tip:** If you have screenshots to contribute, open a PR adding them to a `screenshots/` folder!

**Language:** VB.NET  
**License:** MIT

## Features

- Accesses Windows Shadow Copy Service (VSS)
- Works on Home editions (which lack the built-in Previous Versions UI)
- Browse and restore previous versions of any file or folder
- Supports Windows Vista, 7, and 8
- Simple file browser interface

## System Requirements

- Windows 7 or later
- Visual Studio 2010+ (Community edition works)
- .NET Framework 4.0 or later

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
