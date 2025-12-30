#Requires -Version 5.1

<#
.SYNOPSIS
    Organizes music files into Artist > Album > Track folder structure based on metadata tags.

.DESCRIPTION
    Recursively scans a source directory for MP3 and FLAC files, reads their metadata tags
    using TagLibSharp, and copies them to a destination directory organized as:
    
    Single-disc:  Artist\Album\## - Title.ext
    Multi-disc:   Artist\Album\Disc ##\## - Title.ext

.PARAMETER SourcePath
    The root directory to scan for music files.

.PARAMETER DestinationPath
    The root directory where organized files will be copied.

.PARAMETER LogPath
    Optional. Directory where log files will be written. Defaults to current directory.

.EXAMPLE
    .\Organize-MusicLibrary.ps1 -SourcePath "D:\Unsorted Music" -DestinationPath "D:\Music Library"

.EXAMPLE
    .\Organize-MusicLibrary.ps1 -SourcePath "D:\Unsorted" -DestinationPath "D:\Organized" -LogPath "D:\Logs" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$SourcePath,

    [Parameter(Mandatory)]
    [string]$DestinationPath,

    [Parameter()]
    [string]$LogPath = (Get-Location).Path

)

# ============================================================================
# CONFIGURATION
# ============================================================================

$SupportedExtensions = @('.mp3', '.flac')
$UnknownArtist = 'Unknown Artist'
$UnknownAlbum = 'Unknown Album'

# ============================================================================
# LOGGING SETUP
# ============================================================================

$script:LogFile = $null

function Initialize-Logging {
    param([string]$LogDirectory)

    if (-not (Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $script:LogFile = Join-Path $LogDirectory "MusicOrganizer_$timestamp.log"

    # Write header
    $header = @"
================================================================================
MUSIC LIBRARY ORGANIZER - LOG FILE
================================================================================
Started:     $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Source:      $SourcePath
Destination: $DestinationPath
Mode:        $(if ($WhatIf) { "PREVIEW (WhatIf)" } else { "LIVE" })
================================================================================

"@
    $header | Out-File -FilePath $script:LogFile -Encoding UTF8
    Write-Verbose "Log file created: $($script:LogFile)"
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a message to both the console and the log file.
    
    .PARAMETER Message
        The message to write.
    
    .PARAMETER Level
        Log level: INFO, SUCCESS, WARNING, ERROR, DEBUG
    
    .PARAMETER NoConsole
        If specified, writes only to the log file.
    #>
    param(
        [Parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [string]$Message,

        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO',

        [switch]$NoConsole
    )

    # Silently return if message is empty or null
    if ([string]::IsNullOrEmpty($Message)) {
        return
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] [$Level] $Message"

    # Write to log file
    if ($script:LogFile) {
        $logLine | Out-File -FilePath $script:LogFile -Append -Encoding UTF8
    }

    # Write to console with color coding
    if (-not $NoConsole) {
        $color = switch ($Level) {
            'SUCCESS' { 'Green' }
            'WARNING' { 'Yellow' }
            'ERROR'   { 'Red' }
            'DEBUG'   { 'DarkGray' }
            default   { 'White' }
        }
        Write-Host $Message -ForegroundColor $color
    }
}

function Write-LogOperation {
    <#
    .SYNOPSIS
        Logs a file operation with full details (primarily to log file).
    #>
    param(
        [string]$SourceFile,
        [string]$DestinationFile,
        [ValidateSet('SUCCESS', 'FAILED', 'SKIPPED', 'WHATIF')]
        [string]$Status,
        [string]$Details = ''
    )

    $logEntry = @"
  SOURCE:      $SourceFile
  DESTINATION: $DestinationFile
  STATUS:      $Status$(if ($Details) { "`n  DETAILS:     $Details" })

"@
    if ($script:LogFile) {
        $logEntry | Out-File -FilePath $script:LogFile -Append -Encoding UTF8
    }
}

# ============================================================================
# TAGLIB SETUP
# ============================================================================

function Initialize-TagLib {
    # Force installation even during -WhatIf mode (we need TagLib to read metadata for preview)
    $savedWhatIfPreference = $WhatIfPreference
    $script:WhatIfPreference = $false
    
    try {
        # Check if already loaded
        if ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'TagLibSharp' }) {
            Write-Log "TagLibSharp already loaded" -Level DEBUG -NoConsole
            return $true
        }

        # Option 1: Try loading from NuGet packages cache
        $nugetPaths = @(
            "$env:USERPROFILE\.nuget\packages\taglibsharp\*\lib\netstandard2.0\TagLibSharp.dll"
            "$env:USERPROFILE\.nuget\packages\taglib-sharp-netstandard\*\lib\netstandard2.0\TagLib.Sharp.dll"
        )
        
        foreach ($pattern in $nugetPaths) {
            $dllPath = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
                       Sort-Object {
                           $versionStr = $_.Directory.Parent.Name -replace '[^\d.]'
                           if ($versionStr -match '^\d+(\.\d+)*$') { [version]$versionStr } else { [version]'0.0' }
                       } -Descending |
                       Select-Object -First 1
            
            if ($dllPath) {
                try {
                    Add-Type -Path $dllPath.FullName
                    Write-Log "Loaded TagLibSharp from: $($dllPath.FullName)" -Level DEBUG -NoConsole
                    return $true
                }
                catch {
                    Write-Log "Failed to load from $($dllPath.FullName): $_" -Level DEBUG -NoConsole
                }
            }
        }

        # Option 2: Install via NuGet
        Write-Log "TagLibSharp not found. Attempting to install via NuGet..." -Level WARNING
        
        try {
            if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -WhatIf:$false | Out-Null
            }

            $installPath = Join-Path $env:TEMP "TagLibSharp"
            if (-not (Test-Path $installPath)) {
                New-Item -ItemType Directory -Path $installPath -Force -WhatIf:$false | Out-Null
            }

            $nugetUrl = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
            $nugetExe = Join-Path $installPath "nuget.exe"
            
            if (-not (Test-Path $nugetExe)) {
                Write-Log "Downloading NuGet CLI..." -Level DEBUG -NoConsole
                # Use WebClient for reliable binary download (Invoke-WebRequest can mangle binaries)
                (New-Object System.Net.WebClient).DownloadFile($nugetUrl, $nugetExe)
            }

            if (-not (Test-Path $nugetExe)) {
                Write-Log "Failed to download nuget.exe" -Level ERROR
                return $false
            }

            Write-Log "Installing TagLibSharp via NuGet..." -Level DEBUG -NoConsole

            #nuget doesn't come with sources. Force add it...
            & $nugetExe sources Add -Name "nuget.org" -Source "https://api.nuget.org/v3/index.json" -NonInteractive | Out-Null
            & $nugetExe install TagLibSharp -OutputDirectory $installPath -NonInteractive | Out-Null
            
            $dllPath = Get-ChildItem -Path "$installPath\TagLibSharp*\lib\netstandard2.0\TagLibSharp.dll" -Recurse -ErrorAction SilentlyContinue |
                       Select-Object -First 1

            if ($dllPath) {
                Add-Type -Path $dllPath.FullName
                Write-Log "TagLibSharp installed and loaded successfully." -Level SUCCESS
                return $true
            }
            else {
                Write-Log "TagLibSharp DLL not found after install in: $installPath" -Level ERROR
            }
        }
        catch {
            Write-Log "Failed to install TagLibSharp: $_" -Level ERROR
        }

        return $false
    }
    finally {
        # Restore original WhatIf preference
        $script:WhatIfPreference = $savedWhatIfPreference
    }
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Get-SanitizedFileName {
    <#
    .SYNOPSIS
        Removes or replaces characters invalid in Windows file/folder names.
    #>
    param([string]$Name)
    
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $null
    }

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $sanitized = ($Name.ToCharArray() | ForEach-Object {
        if ($_ -in $invalid) { '_' } else { $_ }
    }) -join ''
    
    $sanitized = $sanitized.Trim(' .')
    
    if ([string]::IsNullOrWhiteSpace($sanitized)) {
        return $null
    }

    return $sanitized
}

function Get-AlbumKey {
    <#
    .SYNOPSIS
        Creates a unique key for identifying an album (Artist + Album combination).
        Used for grouping files and detecting multi-disc albums.
    #>
    param(
        [string]$Artist,
        [string]$Album
    )

    $artistKey = if ([string]::IsNullOrWhiteSpace($Artist)) { $UnknownArtist } else { $Artist.ToLower().Trim() }
    $albumKey = if ([string]::IsNullOrWhiteSpace($Album)) { $UnknownAlbum } else { $Album.ToLower().Trim() }
    
    return "$artistKey|||$albumKey"
}

function Get-MusicFileMetadata {
    <#
    .SYNOPSIS
        Extracts relevant metadata from an audio file using TagLibSharp.
    #>
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )

    try {
        $tagFile = [TagLib.File]::Create($File.FullName)
        
        $artist = $tagFile.Tag.FirstAlbumArtist
        if ([string]::IsNullOrWhiteSpace($artist)) {
            $artist = $tagFile.Tag.FirstPerformer
        }

        $metadata = [PSCustomObject]@{
            SourceFile  = $File
            Artist      = if ([string]::IsNullOrWhiteSpace($artist)) { $null } else { $artist }
            Album       = if ([string]::IsNullOrWhiteSpace($tagFile.Tag.Album)) { $null } else { $tagFile.Tag.Album }
            Title       = if ([string]::IsNullOrWhiteSpace($tagFile.Tag.Title)) { $null } else { $tagFile.Tag.Title }
            TrackNumber = [uint32]$tagFile.Tag.Track      # 0 if not set
            DiscNumber  = [uint32]$tagFile.Tag.Disc       # 0 if not set
            DiscCount   = [uint32]$tagFile.Tag.DiscCount  # 0 if not set
            AlbumKey    = Get-AlbumKey -Artist $artist -Album $tagFile.Tag.Album
        }

        $tagFile.Dispose()
        return $metadata
    }
    catch {
        Write-Log "Failed to read tags from '$($File.FullName)': $_" -Level WARNING
        return $null
    }
}

function Get-DestinationPath {
    <#
    .SYNOPSIS
        Constructs the destination path for a music file based on its metadata.
    
    .DESCRIPTION
        Builds path as:
        - Single-disc:  DestRoot\Artist\Album\## - Title.ext
        - Multi-disc:   DestRoot\Artist\Album\Disc ##\## - Title.ext
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DestinationRoot,
        
        [Parameter(Mandatory)]
        [PSCustomObject]$Metadata,
        
        [Parameter(Mandatory)]
        [bool]$IsMultiDisc
    )

    # Sanitize artist and album
    $artistFolder = Get-SanitizedFileName -Name $Metadata.Artist
    if (-not $artistFolder) { $artistFolder = $UnknownArtist }

    $albumFolder = Get-SanitizedFileName -Name $Metadata.Album
    if (-not $albumFolder) { $albumFolder = $UnknownAlbum }

    # Build base path
    $basePath = Join-Path $DestinationRoot $artistFolder | Join-Path -ChildPath $albumFolder

    # Add disc folder if multi-disc album
    if ($IsMultiDisc) {
        $discNum = if ($Metadata.DiscNumber -gt 0) { $Metadata.DiscNumber } else { 1 }
        $discFolder = "Disc {0:D2}" -f $discNum
        $basePath = Join-Path $basePath $discFolder
    }

    # Build filename
    $title = Get-SanitizedFileName -Name $Metadata.Title
    $extension = $Metadata.SourceFile.Extension

    if ($title) {
        if ($Metadata.TrackNumber -gt 0) {
            $trackNum = $Metadata.TrackNumber.ToString("D2")
            $fileName = "$trackNum - $title$extension"
        }
        else {
            $fileName = "$title$extension"
        }
    }
    else {
        $fileName = $Metadata.SourceFile.Name
    }

    return Join-Path $basePath $fileName
}

function Copy-MusicFile {
    <#
    .SYNOPSIS
        Copies a music file to its destination, handling duplicates.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Source,
        
        [Parameter(Mandatory)]
        [string]$Destination,
        
        [switch]$WhatIf
    )

    # Handle duplicates
    $finalDestination = $Destination
    $counter = 1
    
    while (Test-Path -LiteralPath $finalDestination) {
        $dir = [System.IO.Path]::GetDirectoryName($Destination)
        $name = [System.IO.Path]::GetFileNameWithoutExtension($Destination)
        $ext = [System.IO.Path]::GetExtension($Destination)
        $finalDestination = Join-Path $dir "$name ($counter)$ext"
        $counter++
    }

    if ($WhatIf) {
        Write-LogOperation -SourceFile $Source -DestinationFile $finalDestination -Status 'WHATIF'
        return @{ Success = $true; Destination = $finalDestination }
    }

    try {
        $destDir = [System.IO.Path]::GetDirectoryName($finalDestination)
        if (-not (Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }

        Copy-Item -LiteralPath $Source -Destination $finalDestination -Force
        Write-LogOperation -SourceFile $Source -DestinationFile $finalDestination -Status 'SUCCESS'
        return @{ Success = $true; Destination = $finalDestination }
    }
    catch {
        Write-LogOperation -SourceFile $Source -DestinationFile $finalDestination -Status 'FAILED' -Details $_.Exception.Message
        return @{ Success = $false; Destination = $finalDestination; Error = $_.Exception.Message }
    }
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

# Initialize logging
Initialize-Logging -LogDirectory $LogPath

# Initialize TagLibSharp
if (-not (Initialize-TagLib)) {
    Write-Log "Cannot proceed without TagLibSharp. Please install it manually or check your internet connection." -Level ERROR
    exit 1
}

# Resolve paths
$SourcePath = (Resolve-Path $SourcePath).Path
$DestinationPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationPath)

Write-Log ""
Write-Log "========================================"
Write-Log "MUSIC LIBRARY ORGANIZER"
Write-Log "========================================"
Write-Log "Source:      $SourcePath"
Write-Log "Destination: $DestinationPath"
Write-Log "Log File:    $($script:LogFile)"
if ($WhatIf) {
    Write-Log "Mode:        PREVIEW (no files will be copied)" -Level WARNING
}
Write-Log ""

# ============================================================================
# PASS 1: ANALYSIS - Discover files and identify multi-disc albums
# ============================================================================

Write-Log "PASS 1: Scanning and analyzing music files..." -Level INFO

$musicFiles = Get-ChildItem -Path $SourcePath -Recurse -File | 
              Where-Object { $_.Extension.ToLower() -in $SupportedExtensions }

$totalFiles = $musicFiles.Count
Write-Log "Found $totalFiles music files." -Level SUCCESS

if ($totalFiles -eq 0) {
    Write-Log "No supported music files found. Exiting." -Level WARNING
    exit 0
}

# Read all metadata and identify multi-disc albums
$fileMetadata = @()
$albumDiscInfo = @{}  # Key = AlbumKey, Value = HashSet of disc numbers

$analyzed = 0
foreach ($file in $musicFiles) {
    $analyzed++
    Write-Progress -Activity "Pass 1: Analyzing Files" `
                   -Status "Reading: $($file.Name)" `
                   -PercentComplete (($analyzed / $totalFiles) * 100)

    $metadata = Get-MusicFileMetadata -File $file
    
    if ($metadata) {
        $fileMetadata += $metadata
        
        # Track disc numbers per album
        if (-not $albumDiscInfo.ContainsKey($metadata.AlbumKey)) {
            $albumDiscInfo[$metadata.AlbumKey] = [System.Collections.Generic.HashSet[uint32]]::new()
        }
        
        # Add disc number (use 1 if not specified)
        $discNum = if ($metadata.DiscNumber -gt 0) { $metadata.DiscNumber } else { 1 }
        [void]$albumDiscInfo[$metadata.AlbumKey].Add($discNum)
        
        # Also consider DiscCount tag
        if ($metadata.DiscCount -gt 1) {
            # Mark as multi-disc even if we only see disc 1
            [void]$albumDiscInfo[$metadata.AlbumKey].Add(0)  # Sentinel for "known multi-disc"
        }
    }
}

Write-Progress -Activity "Pass 1: Analyzing Files" -Completed

# Determine which albums are multi-disc
# An album is multi-disc if: (a) we found multiple disc numbers, or (b) DiscCount > 1 was set
$multiDiscAlbums = [System.Collections.Generic.HashSet[string]]::new()

foreach ($albumKey in $albumDiscInfo.Keys) {
    $discNumbers = $albumDiscInfo[$albumKey]
    
    # Remove sentinel and check
    $actualDiscs = $discNumbers | Where-Object { $_ -gt 0 }
    $hasMultipleDiscs = ($actualDiscs | Measure-Object).Count -gt 1
    $hasDiscCountFlag = $discNumbers.Contains(0)  # Sentinel indicates DiscCount > 1
    
    if ($hasMultipleDiscs -or $hasDiscCountFlag) {
        [void]$multiDiscAlbums.Add($albumKey)
    }
}

$multiDiscCount = $multiDiscAlbums.Count
if ($multiDiscCount -gt 0) {
    Write-Log "Detected $multiDiscCount multi-disc album(s)." -Level INFO
    Write-Log "Multi-disc albums will use 'Disc ##' subfolders." -Level DEBUG -NoConsole
}

# Log multi-disc albums to file
if ($multiDiscCount -gt 0 -and $script:LogFile) {
    "`nMULTI-DISC ALBUMS DETECTED:" | Out-File -FilePath $script:LogFile -Append -Encoding UTF8
    foreach ($key in $multiDiscAlbums) {
        $parts = $key -split '\|\|\|'
        "  - $($parts[0]) / $($parts[1])" | Out-File -FilePath $script:LogFile -Append -Encoding UTF8
    }
    "" | Out-File -FilePath $script:LogFile -Append -Encoding UTF8
}

Write-Log ""

# ============================================================================
# PASS 2: ORGANIZATION - Copy files to destination
# ============================================================================

Write-Log "PASS 2: Organizing files..." -Level INFO
Write-Log "" -NoConsole

if ($script:LogFile) {
    "FILE OPERATIONS:" | Out-File -FilePath $script:LogFile -Append -Encoding UTF8
    "---------------" | Out-File -FilePath $script:LogFile -Append -Encoding UTF8
}

$processed = 0
$succeeded = 0
$failed = 0
$skippedCount = ($totalFiles - $fileMetadata.Count)  # Files we couldn't read metadata from

foreach ($metadata in $fileMetadata) {
    $processed++
    $percentComplete = [math]::Round(($processed / $fileMetadata.Count) * 100, 1)
    
    Write-Progress -Activity "Pass 2: Organizing Files" `
                   -Status "Copying: $($metadata.SourceFile.Name)" `
                   -PercentComplete $percentComplete `
                   -CurrentOperation "$processed of $($fileMetadata.Count) files"

    # Check if this album is multi-disc
    $isMultiDisc = $multiDiscAlbums.Contains($metadata.AlbumKey)

    # Calculate destination
    $destination = Get-DestinationPath -DestinationRoot $DestinationPath `
                                       -Metadata $metadata `
                                       -IsMultiDisc $isMultiDisc

    # Console output
    $artistDisplay = if ($metadata.Artist) { $metadata.Artist } else { $UnknownArtist }
    $albumDisplay = if ($metadata.Album) { $metadata.Album } else { $UnknownAlbum }
    $discDisplay = if ($isMultiDisc -and $metadata.DiscNumber -gt 0) { " [Disc $($metadata.DiscNumber)]" } else { "" }
    
    Write-Host "[$processed/$($fileMetadata.Count)] " -NoNewline -ForegroundColor DarkGray
    Write-Host "$artistDisplay" -NoNewline -ForegroundColor Cyan
    Write-Host " / " -NoNewline
    Write-Host "$albumDisplay$discDisplay" -NoNewline -ForegroundColor Magenta
    Write-Host " / " -NoNewline
    Write-Host "$($metadata.SourceFile.Name)" -ForegroundColor White

    if ($WhatIf) {
        Write-Host "  -> $destination" -ForegroundColor DarkCyan
    }

    # Copy file
    $result = Copy-MusicFile -Source $metadata.SourceFile.FullName -Destination $destination -WhatIf:$WhatIf
    
    if ($result.Success) {
        $succeeded++
    }
    else {
        $failed++
        Write-Host "  ERROR: $($result.Error)" -ForegroundColor Red
    }
}

Write-Progress -Activity "Pass 2: Organizing Files" -Completed

# ============================================================================
# SUMMARY
# ============================================================================

$summaryText = @"

========================================
SUMMARY
========================================
Total files scanned:    $totalFiles
Metadata read:          $($fileMetadata.Count)
Multi-disc albums:      $multiDiscCount
Successfully copied:    $succeeded
Failed:                 $failed
Skipped (no metadata):  $skippedCount
========================================
"@

Write-Log $summaryText

# Write summary to log file
if ($script:LogFile) {
    $summaryText | Out-File -FilePath $script:LogFile -Append -Encoding UTF8
    
    $footer = @"

================================================================================
Completed: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
================================================================================
"@
    $footer | Out-File -FilePath $script:LogFile -Append -Encoding UTF8
    
    Write-Log ""
    Write-Log "Full log written to: $($script:LogFile)" -Level INFO
}
