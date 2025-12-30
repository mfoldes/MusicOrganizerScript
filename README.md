# Music Library Organizer

So I had a lot of time on my hands and a fairly enormeous project over the holiday break. I have a lot of music and organizing it has been agonizingly difficult.

I present to you a PowerShell script that organizes messy music collections into a clean `Artist > Album > Track` folder structure using metadata tags. Because, better living through PowerShell.

## Features

- **Tag-based organization** — Reads ID3v2 (MP3) and Vorbis Comments (FLAC) via TagLibSharp
- **Multi-disc album support** — Automatically detects and creates `Disc ##` subfolders
- **Safe operation** — Copies files (non-destructive), handles duplicates gracefully
- **Preview mode** — `-WhatIf` flag shows what would happen without copying anything
- **Detailed logging** — Timestamped log files with full operation details
- **Graceful fallbacks** — Files with missing tags go to `Unknown Artist/Unknown Album`
- **Special character handling** — Properly handles paths containing wildcards and special characters

## Requirements

- PowerShell 5.1 or later
- Windows (uses Windows-specific path handling)
- Internet connection on first run (to auto-install TagLibSharp if not present)

## Installation

```powershell
git clone https://github.com/yourusername/MusicOrganizer.git
cd MusicOrganizer
```

Or simply download `Organize-MusicLibrary.ps1` directly and run it. It should pull Nuget and the TagLibSharp library into your temp directory.

If you run into issues, just install TagLibSharp with nuget command line utility and place the dll in your user temp folder's TagLibSharp directory and the script can pick it up from there.

## Usage

### Basic Usage (Full Send Mode)

```powershell
.\Organize-MusicLibrary.ps1 -SourcePath "D:\Unsorted Music" -DestinationPath "D:\Music Library"
```

### Preview Mode (I want to see what I am about to break Mode)

Use `-WhatIf` to see what the script would do without actually copying any files:

```powershell
.\Organize-MusicLibrary.ps1 -SourcePath "D:\Unsorted Music" -DestinationPath "D:\Music Library" -WhatIf
```

> **Note:** The script will still download and install TagLibSharp during `-WhatIf` mode, as the library is required to read metadata for the preview (at least that's how it's supposed to work. If it doesn't just run it the normal way first and it'll download and set everything up).

### Custom Log Location

```powershell
.\Organize-MusicLibrary.ps1 -SourcePath "D:\Unsorted" -DestinationPath "D:\Library" -LogPath "D:\Logs"
```

## Output Structure

### Single-Disc Albums

```
Music Library/
└── Pink Floyd/
    └── Dark Side of the Moon/
        ├── 01 - Speak to Me.flac
        ├── 02 - Breathe.flac
        └── ...
```

### Multi-Disc Albums

```
Music Library/
└── Pink Floyd/
    └── The Wall/
        ├── Disc 01/
        │   ├── 01 - In The Flesh.flac
        │   └── ...
        └── Disc 02/
            ├── 01 - Hey You.flac
            └── ...
```

## How It Works

The script uses a two-pass approach:

1. **Pass 1 (Analysis)** — Scans all files, reads metadata, and identifies which albums span multiple discs
2. **Pass 2 (Organization)** — Copies files to their destinations with consistent folder structures

This ensures that multi-disc albums always get disc subfolders for *all* discs, not just Disc 2+.

## Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-SourcePath` | Yes | Root directory to scan for music files |
| `-DestinationPath` | Yes | Root directory for organized output |
| `-LogPath` | No | Directory for log files (defaults to current directory) |
| `-WhatIf` | No | Preview mode — no files are copied (TagLibSharp still installs if needed) |

## Supported Formats

- MP3 (`.mp3`)
- FLAC (`.flac`)

Additional formats can be added by modifying the `$SupportedExtensions` array in the script. LibTagSharp is an incredibly versitle library for music metadata tagging, so I know it can read more, I just didn't test them. I 100% know FLAC and MP3 files work.

## Tag Priority

For artist names, the script checks tags in this order:

1. `AlbumArtist` (preferred — consistent across compilations)
2. `Artist` / `Performer` (fallback)
3. `Unknown Artist` (if no artist tag exists)

## Log File Format

Log files are created with timestamps (e.g., `MusicOrganizer_20250115_143245.log`) and include:

- Configuration summary
- List of detected multi-disc albums
- Per-file operation details (source, destination, status)
- Final summary statistics

## Edge Cases Handled

| Scenario | Behavior |
|----------|----------|
| Invalid filename characters (`? * : " < > \|`) | Replaced with `_` |
| Duplicate destinations | Files get `(1)`, `(2)` suffixes |
| Missing tags | Falls back to `Unknown Artist/Unknown Album` or original filename |
| Multi-disc detection | Uses both disc number tags and disc count metadata |
| Paths with special characters | Uses `-LiteralPath` to handle wildcards and brackets correctly |

## Dependencies

- **[TagLibSharp](https://github.com/mono/taglib-sharp)** — .NET library for reading audio metadata. They're great. It's a great project. Go donate beer money to them because this would have been a lot harder if they didn't put this together.

The script will automatically download and install TagLibSharp via NuGet on first run if it's not already present. The installation uses `System.Net.WebClient` to pull the Nuget command line utility (Invoke-WebRequest is not good here), explicitly configures the NuGet source so it should get picked up on the first run, then installs TagLibSharp via the Nuget command line to your temp directory and places the dll there so the script can load and reference it.

## Troubleshooting

### TagLibSharp fails to install

If automatic installation fails:

1. Ensure you have internet connectivity
2. Try running PowerShell as Administrator
3. Manually install via NuGet:
   ```powershell
   nuget install TagLibSharp -OutputDirectory $env:TEMP\TagLibSharp
   ```

### Files with special characters aren't found

The script uses `-LiteralPath` for file operations, which should handle most special characters. If you encounter issues, ensure your PowerShell version is 5.1 or later.

## License

MIT License — feel free to use, modify, and distribute. I don't really care. This is mainly for my personal use, but if you can take it and make it better, feel free. If it helps you out, even better.

