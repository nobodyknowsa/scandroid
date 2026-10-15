[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$Source,

    [Parameter(Mandatory = $true)]
    [string]$Destination,

    [string]$ExifToolPath = 'C:\script\exiftool\exiftool.exe',

    [ValidateRange(1, 2000)]
    [int]$BatchSize = 200,

    [switch]$WhatIf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$PhotoExtensions = @(
    '.jpg', '.jpeg', '.jpe', '.jfif', '.png', '.heic', '.heif', '.webp', '.avif', '.jxl',
    '.tif', '.tiff', '.bmp', '.gif', '.dng', '.raw', '.cr2', '.cr3', '.nef', '.nrw',
    '.arw', '.sr2', '.srf', '.orf', '.rw2', '.raf', '.pef', '.ptx', '.3fr', '.ari',
    '.bay', '.cap', '.eip', '.erf', '.fff', '.gpr', '.iiq', '.k25', '.kdc', '.mdc',
    '.mef', '.mos', '.mrw', '.rwl', '.x3f', '.dcr', '.jp2', '.j2k', '.jpf', '.jpx',
    '.psd', '.psb', '.exr'
)

$VideoExtensions = @(
    '.mp4', '.mov', '.m4v', '.avi', '.mkv', '.wmv', '.mts', '.m2ts', '.3gp',
    '.mpg', '.mpeg', '.webm', '.vob', '.ts', '.m2v', '.mpg', '.divx', '.flv',
    '.f4v', '.asf', '.ogv', '.mxf', '.rm', '.rmvb', '.vro'
)

$Culture = [Globalization.CultureInfo]::InvariantCulture
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if (-not (Test-Path -LiteralPath $ExifToolPath -PathType Leaf)) {
    throw "ExifTool не найден: $ExifToolPath"
}
$ExifToolPath = (Resolve-Path -LiteralPath $ExifToolPath).Path

function Normalize-DirectoryPath {
    param([string]$Path)
    $Full = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $Path).Path)
    $Root = [IO.Path]::GetPathRoot($Full)
    if ($Full.Length -gt $Root.Length) {
        return $Full.TrimEnd([char[]]@('\', '/'))
    }
    return $Full
}

$SourcePaths = @(
    foreach ($Path in $Source) {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
            throw "Исходный каталог не найден: $Path"
        }
        Normalize-DirectoryPath -Path $Path
    }
) | Select-Object -Unique

# Destination may not be inside a source or contain a source. This prevents recursion.
$Destination = [IO.Path]::GetFullPath($Destination)
$DestinationRoot = [IO.Path]::GetPathRoot($Destination)
if ($Destination.Length -gt $DestinationRoot.Length) {
    $Destination = $Destination.TrimEnd([char[]]@('\', '/'))
}

foreach ($Path in $SourcePaths) {
    $SourceRoot = [IO.Path]::GetPathRoot($Path)
    $SourcePrefix = if ($Path.Length -gt $SourceRoot.Length) { $Path.TrimEnd([char[]]@('\', '/')) + '\' } else { $Path }
    $DestinationPrefix = $Destination.TrimEnd([char[]]@('\', '/')) + '\'
    $DestinationIsSource = $Destination.Equals($Path, [StringComparison]::OrdinalIgnoreCase)
    $DestinationInsideSource = $Destination.StartsWith($SourcePrefix, [StringComparison]::OrdinalIgnoreCase)
    $SourceInsideDestination = $Path.StartsWith($DestinationPrefix, [StringComparison]::OrdinalIgnoreCase)

    if ($DestinationIsSource -or $DestinationInsideSource -or $SourceInsideDestination) {
        throw "Источник и архив не должны пересекаться: $Path"
    }
}

New-Item -ItemType Directory -Path $Destination -Force | Out-Null
$PhotoRoot = Join-Path $Destination 'Photos'
$VideoRoot = Join-Path $Destination 'Videos'
New-Item -ItemType Directory -Path $PhotoRoot -Force | Out-Null
New-Item -ItemType Directory -Path $VideoRoot -Force | Out-Null

$JournalPath = Join-Path $Destination '_media-collector-journal.csv'
$TempDirectory = Join-Path $env:TEMP ('MediaCollector_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $TempDirectory -Force | Out-Null

$Stats = [ordered]@{
    FoundPhotos = 0
    FoundVideos = 0
    CopiedPhotos = 0
    CopiedVideos = 0
    DuplicatePhotos = 0
    Resumed = 0
    Failed = 0
    PhotoFallbackDate = 0
    VideoFallbackDate = 0
    Planned = 0
}

function Get-FileKey {
    param([IO.FileInfo]$File)
    return ($File.FullName.ToLowerInvariant() + '|' + $File.Length + '|' + $File.LastWriteTimeUtc.Ticks)
}

function Get-MetadataValue {
    param([object]$Item, [string[]]$Names)
    foreach ($Name in $Names) {
        foreach ($Property in $Item.PSObject.Properties) {
            if ($Property.Name -eq $Name -or $Property.Name.EndsWith(':' + $Name, [StringComparison]::OrdinalIgnoreCase)) {
                $Value = [string]$Property.Value
                if (-not [string]::IsNullOrWhiteSpace($Value)) {
                    return [PSCustomObject]@{ Value = $Value.Trim(); Tag = $Property.Name }
                }
            }
        }
    }
    return $null
}

function Get-ParsedDate {
    param([string]$RawDate)
    if ([string]::IsNullOrWhiteSpace($RawDate)) { return $null }

    $Candidate = $RawDate.Trim()
    if ($Candidate -match '^(\d{4}):(\d{2}):(\d{2})(.*)$') {
        $Candidate = $Matches[1] + '-' + $Matches[2] + '-' + $Matches[3] + $Matches[4]
    }

    $Parsed = [datetime]::MinValue
    $Styles = [Globalization.DateTimeStyles]::AllowWhiteSpaces
    if ([datetime]::TryParse($Candidate, $Culture, $Styles, [ref]$Parsed) -and $Parsed.Year -ge 1800 -and $Parsed.Year -le 2200) {
        return $Parsed
    }

    if ($Candidate -match '^(\d{4}-\d{2}-\d{2})') {
        if ([datetime]::TryParseExact($Matches[1], 'yyyy-MM-dd', $Culture, [Globalization.DateTimeStyles]::None, [ref]$Parsed)) {
            if ($Parsed.Year -ge 1800 -and $Parsed.Year -le 2200) { return $Parsed }
        }
    }
    return $null
}

function Get-PhotoDate {
    param([object]$Metadata, [IO.FileInfo]$File)
    $Value = Get-MetadataValue -Item $Metadata -Names @(
        'DateTimeOriginal', 'SubSecDateTimeOriginal', 'CreateDate', 'DateCreated',
        'MediaCreateDate', 'TrackCreateDate', 'ContentCreateDate', 'CreationDate'
    )
    if ($null -ne $Value) {
        $Date = Get-ParsedDate -RawDate $Value.Value
        if ($null -ne $Date) { return [PSCustomObject]@{ Date = $Date; Source = $Value.Tag } }
    }
    return [PSCustomObject]@{ Date = $File.LastWriteTime; Source = 'FileLastWriteTime' }
}

function Get-VideoDate {
    param([object]$Metadata, [IO.FileInfo]$File)
    # Prefer embedded video/container creation timestamps. No video hashing is performed.
    $Value = Get-MetadataValue -Item $Metadata -Names @(
        'MediaCreateDate', 'CreateDate', 'TrackCreateDate', 'ContentCreateDate',
        'DateCreated', 'CreationDate', 'DateTimeOriginal'
    )
    if ($null -ne $Value) {
        $Date = Get-ParsedDate -RawDate $Value.Value
        if ($null -ne $Date) { return [PSCustomObject]@{ Date = $Date; Source = $Value.Tag } }
    }
    return [PSCustomObject]@{ Date = $File.LastWriteTime; Source = 'FileLastWriteTime' }
}

function Write-Journal {
    param(
        [string]$Status,
        [string]$Kind,
        [string]$SourceFile,
        [string]$TargetFile,
        [string]$Hash,
        [string]$DateSource,
        [string]$ErrorMessage
    )
    $Entry = [PSCustomObject]@{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Status = $Status
        Kind = $Kind
        Source = $SourceFile
        Destination = $TargetFile
        SHA256 = $Hash
        DateSource = $DateSource
        Error = $ErrorMessage
    }
    if (Test-Path -LiteralPath $JournalPath -PathType Leaf) {
        $Entry | Export-Csv -LiteralPath $JournalPath -NoTypeInformation -Encoding UTF8 -Append
    } else {
        $Entry | Export-Csv -LiteralPath $JournalPath -NoTypeInformation -Encoding UTF8
    }
}

function Get-UniqueTarget {
    param([string]$Directory, [string]$FileName)
    $Base = [IO.Path]::GetFileNameWithoutExtension($FileName)
    $Ext = [IO.Path]::GetExtension($FileName)
    $Candidate = Join-Path $Directory $FileName
    $Index = 2
    while (Test-Path -LiteralPath $Candidate) {
        $Candidate = Join-Path $Directory ($Base + '_' + $Index + $Ext)
        $Index++
    }
    return $Candidate
}

try {
    # Recover completed source files from the journal. Recheck that the recorded destination exists.
    $CompletedKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    if (Test-Path -LiteralPath $JournalPath -PathType Leaf) {
        Write-Host 'Читаю журнал предыдущих запусков...' -ForegroundColor Cyan
        foreach ($Entry in (Import-Csv -LiteralPath $JournalPath)) {
            if ($Entry.Status -notin @('CopiedPhoto', 'DuplicatePhoto', 'CopiedVideo')) { continue }
            if (-not (Test-Path -LiteralPath $Entry.Source -PathType Leaf)) { continue }
            if ($Entry.Status -in @('CopiedPhoto', 'CopiedVideo') -and -not (Test-Path -LiteralPath $Entry.Destination -PathType Leaf)) { continue }
            $ExistingSource = Get-Item -LiteralPath $Entry.Source
            [void]$CompletedKeys.Add((Get-FileKey -File $ExistingSource))
        }
    }

    # Only photo files are hashed for deduplication. Video files are intentionally not hashed.
    $SeenPhotoHashes = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $ExistingPhotoFiles = @(
        Get-ChildItem -LiteralPath $PhotoRoot -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $PhotoExtensions -contains $_.Extension.ToLowerInvariant() }
    )
    Write-Host "Индексирую существующие фотографии: $($ExistingPhotoFiles.Count)" -ForegroundColor Cyan
    foreach ($Existing in $ExistingPhotoFiles) {
        try {
            $Hash = (Get-FileHash -LiteralPath $Existing.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
            [void]$SeenPhotoHashes.Add($Hash)
        } catch {
            $Stats.Failed++
            Write-Warning "Не удалось проверить фото в архиве: $($Existing.FullName)"
        }
    }

    Write-Host 'Сканирую исходные каталоги...' -ForegroundColor Cyan
    $Files = New-Object 'System.Collections.Generic.List[object]'
    foreach ($SourcePath in $SourcePaths) {
        Write-Host "Источник: $SourcePath"
        Get-ChildItem -LiteralPath $SourcePath -File -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $Ext = $_.Extension.ToLowerInvariant()
            if ($PhotoExtensions -contains $Ext) {
                $Files.Add([PSCustomObject]@{ File = $_; Kind = 'Photo' })
            } elseif ($VideoExtensions -contains $Ext) {
                $Files.Add([PSCustomObject]@{ File = $_; Kind = 'Video' })
            }
        }
    }

    $Stats.FoundPhotos = @($Files | Where-Object Kind -eq 'Photo').Count
    $Stats.FoundVideos = @($Files | Where-Object Kind -eq 'Video').Count
    Write-Host "Найдено фото: $($Stats.FoundPhotos); видео: $($Stats.FoundVideos)"

    for ($Offset = 0; $Offset -lt $Files.Count; $Offset += $BatchSize) {
        $End = [Math]::Min($Offset + $BatchSize - 1, $Files.Count - 1)
        $Batch = @($Files.GetRange($Offset, $End - $Offset + 1))
        $Pending = @(
            foreach ($Item in $Batch) {
                $Key = Get-FileKey -File $Item.File
                if ($CompletedKeys.Contains($Key)) { $Stats.Resumed++; continue }
                $Item
            }
        )
        if ($Pending.Count -eq 0) { continue }

        $ArgFile = Join-Path $TempDirectory 'exiftool-args.txt'
        $Arguments = @(
            '-json', '-s', '-SourceFile', '-DateTimeOriginal', '-SubSecDateTimeOriginal',
            '-CreateDate', '-DateCreated', '-MediaCreateDate', '-TrackCreateDate',
            '-ContentCreateDate', '-CreationDate', '--'
        )
        foreach ($Item in $Pending) { $Arguments += $Item.File.FullName }
        [IO.File]::WriteAllLines($ArgFile, [string[]]$Arguments, $Utf8NoBom)

        Write-Progress -Activity 'Сбор медиаархива' -Status "$($End + 1) из $($Files.Count)" -PercentComplete ([int](($End + 1) * 100 / [Math]::Max(1, $Files.Count)))
        $PreviousErrorActionPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            $ExifOutput = @(& $ExifToolPath '-charset' 'filename=UTF8' '-@' $ArgFile 2>$null)
            $ExifExitCode = $LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $PreviousErrorActionPreference
        }
        $ExifExitCode = $LASTEXITCODE
        $MetadataByPath = @{}
        if ($ExifOutput.Count -gt 0) {
            try {
                $Json = $ExifOutput -join [Environment]::NewLine | ConvertFrom-Json
                foreach ($Item in @($Json)) {
                    $SourceMeta = Get-MetadataValue -Item $Item -Names @('SourceFile')
                    if ($null -ne $SourceMeta) { $MetadataByPath[[string]$SourceMeta.Value] = $Item }
                }
            } catch { Write-Warning 'Не удалось разобрать JSON ExifTool для пакета.' }
        }
        if ($ExifExitCode -ne 0) { Write-Warning 'ExifTool сообщил об ошибках в пакете; для файлов без метаданных будет использована дата изменения.' }

        foreach ($Item in $Pending) {
            $File = $Item.File
            $Kind = $Item.Kind
            try {
                $File = Get-Item -LiteralPath $File.FullName -ErrorAction Stop
                $Key = Get-FileKey -File $File
                if ($CompletedKeys.Contains($Key)) { continue }

                $Metadata = [PSCustomObject]@{}
                if ($MetadataByPath.ContainsKey($File.FullName)) { $Metadata = $MetadataByPath[$File.FullName] }

                if ($Kind -eq 'Photo') {
                    $MediaDate = Get-PhotoDate -Metadata $Metadata -File $File
                    if ($MediaDate.Source -eq 'FileLastWriteTime') { $Stats.PhotoFallbackDate++ }
                    $Hash = (Get-FileHash -LiteralPath $File.FullName -Algorithm SHA256 -ErrorAction Stop).Hash

                    if ($SeenPhotoHashes.Contains($Hash)) {
                        Write-Host "Дубликат фото пропущен: $($File.FullName)" -ForegroundColor DarkGray
                        if (-not $WhatIf) {
                            Write-Journal -Status 'DuplicatePhoto' -Kind 'Photo' -SourceFile $File.FullName -TargetFile '' -Hash $Hash -DateSource $MediaDate.Source -ErrorMessage ''
                            [void]$CompletedKeys.Add($Key)
                        }
                        $Stats.DuplicatePhotos++
                        continue
                    }
                    $TargetRoot = $PhotoRoot
                } else {
                    $MediaDate = Get-VideoDate -Metadata $Metadata -File $File
                    if ($MediaDate.Source -eq 'FileLastWriteTime') { $Stats.VideoFallbackDate++ }
                    $Hash = ''
                    $TargetRoot = $VideoRoot
                }

                $Year = $MediaDate.Date.ToString('yyyy', $Culture)
                $Month = $MediaDate.Date.ToString('MM', $Culture)
                $TargetDirectory = Join-Path (Join-Path $TargetRoot $Year) $Month
                $TargetPath = Get-UniqueTarget -Directory $TargetDirectory -FileName $File.Name

                if ($WhatIf) {
                    Write-Host "[WhatIf][$Kind] $($File.FullName) -> $TargetPath"
                    $Stats.Planned++
                    if ($Kind -eq 'Photo') { [void]$SeenPhotoHashes.Add($Hash) }
                    continue
                }

                New-Item -ItemType Directory -Path $TargetDirectory -Force | Out-Null
                Copy-Item -LiteralPath $File.FullName -Destination $TargetPath -ErrorAction Stop

                if ($Kind -eq 'Photo') {
                    $TargetHash = (Get-FileHash -LiteralPath $TargetPath -Algorithm SHA256 -ErrorAction Stop).Hash
                    if ($TargetHash -ne $Hash) {
                        Remove-Item -LiteralPath $TargetPath -Force -ErrorAction SilentlyContinue
                        throw 'SHA-256 исходника и копии не совпадает.'
                    }
                    Write-Journal -Status 'CopiedPhoto' -Kind 'Photo' -SourceFile $File.FullName -TargetFile $TargetPath -Hash $Hash -DateSource $MediaDate.Source -ErrorMessage ''
                    [void]$SeenPhotoHashes.Add($Hash)
                    $Stats.CopiedPhotos++
                } else {
                    # Video is copied without SHA-256 calculation or duplicate detection.
                    Write-Journal -Status 'CopiedVideo' -Kind 'Video' -SourceFile $File.FullName -TargetFile $TargetPath -Hash '' -DateSource $MediaDate.Source -ErrorMessage ''
                    $Stats.CopiedVideos++
                }

                [void]$CompletedKeys.Add($Key)
                Write-Host "Скопировано [$Kind]: $TargetPath"
            } catch {
                $Stats.Failed++
                Write-Warning ("Ошибка [$Kind] " + $File.FullName + ': ' + $_.Exception.Message)
                if (-not $WhatIf) {
                    try {
                        Write-Journal -Status 'Error' -Kind $Kind -SourceFile $File.FullName -TargetFile '' -Hash '' -DateSource '' -ErrorMessage $_.Exception.Message
                    } catch { Write-Warning 'Не удалось записать ошибку в журнал.' }
                }
            }
        }
    }

    Write-Progress -Activity 'Сбор медиаархива' -Completed
    Write-Host ''$ExifOutput = @(& $ExifToolPath '-charset' 'filename=UTF8' '-@' $ArgFile 2>$null)
    Write-Host '========== ИТОГО ==========' -ForegroundColor Cyan
    Write-Host "Найдено фото: $($Stats.FoundPhotos)"
    Write-Host "Найдено видео: $($Stats.FoundVideos)"
    Write-Host "Скопировано фото: $($Stats.CopiedPhotos)"
    Write-Host "Скопировано видео: $($Stats.CopiedVideos)"
    Write-Host "Пропущено дубликатов фото: $($Stats.DuplicatePhotos)"
    Write-Host "Пропущено как обработанные ранее: $($Stats.Resumed)"
    Write-Host "Ошибок: $($Stats.Failed)"
    Write-Host "Фото без даты в метаданных: $($Stats.PhotoFallbackDate)"
    Write-Host "Видео без даты в метаданных: $($Stats.VideoFallbackDate)"
    Write-Host "Предварительно запланировано: $($Stats.Planned)"
    Write-Host "Журнал: $JournalPath"
}
finally {
    Remove-Item -LiteralPath $TempDirectory -Recurse -Force -ErrorAction SilentlyContinue
}
