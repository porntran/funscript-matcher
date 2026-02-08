<#
.SYNOPSIS
    FUNSCRIPT MATCHER - The Ultimate VR Script Matching Tool (v2.1)
    
.DESCRIPTION
    This tool automates the process of matching local VR videos with a large library of Funscripts.
    It uses fuzzy logic, regex pattern matching, and metadata analysis to pair files even if
    filenames are not identical.

    [CORE FEATURES]
    1. FUZZY MATCHING & SCORING:
       - Matches based on Studio, Date, and Keywords.
       - Weighted logic: Exact matches > Partial matches.
       - Short-word protection: Prevents false positives for short words.
    
    2. SMART STUDIO DETECTION:
       - Automatically detects studios from filenames and parent folders.
       - "Smart Learn": Interactively asks to add new studio patterns to the database if inferred from a match.

    3. CACHING SYSTEM:
       - Caches Video and Script libraries to JSON for instant startup.
       - Configurable auto-refresh logic.

    4. INTERACTIVE REFINING:
       - If a match isn't found, users can type keywords directly into the console to re-search instantly.

.PARAMETER VideoPath
    Path to the folder containing VR videos. Default: Current directory.

.PARAMETER FunscriptsPath
    Path to the folder containing your Funscript collection (e.g., Mega Pack).
    Default: Current directory.

.PARAMETER MatchVideo
    "Target Mode". Provide a part of a filename (e.g., "CzechVR 741"). 
    The script will find the best matching video file and process ONLY that file, ignoring history.

.PARAMETER RefreshFunscripts
    Forces a complete re-scan of the script library (updates cache).

.PARAMETER RefreshVideos
    Forces a complete re-scan of the video directory (updates cache).

.PARAMETER ReprocessHistory
    Ignores the history file and re-processes all videos (even those marked as Done).

.PARAMETER CheckFunscript
    "Search Mode". Does not scan videos. Searches the script library for a string and lists the top 10 matches.

.LINK
    https://github.com/yourusername/funscript-matcher
#>

param (
    [string]$VideoPath = $PSScriptRoot,
    [string]$FunscriptsPath = $PSScriptRoot,
    [string]$MatchVideo = "",
    [switch]$RefreshFunscripts,
    [switch]$RefreshVideos,
    [switch]$ReprocessHistory,
    [string]$CheckFunscript = ""
)

$OutputEncoding = [System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Continue"

# ==============================================================================================
#   1. FILE SYSTEM & PATHS
# ==============================================================================================
$ConfigFile    = "$PSScriptRoot\_funscript_matcher_config.json"
$StudioFile    = "$PSScriptRoot\_funscript_matcher_studios.json"
$LibScriptFile = "$PSScriptRoot\_funscript_matcher_funscripts.json"
$LibVideoFile  = "$PSScriptRoot\_funscript_matcher_videos.json"
$HistoryFile   = "$PSScriptRoot\_funscript_matcher_history.txt"

# ==============================================================================================
#   2. DEFAULT CONFIGURATION
# ==============================================================================================
$DefaultConfig = @{
    # -- STARTUP BEHAVIOR --
    # AutoRefresh_*: $true = Update automatically; $false = Use cache (Silent load)
    
    # Normal Mode (Bulk Processing)
    AutoRefresh_Funscripts_Normal = $false
    AutoRefresh_Videos_Normal     = $true  # Keep true to detect newly downloaded videos
    
    # Check/Target Mode (Single Task)
    AutoRefresh_Funscripts_CheckMode = $false
    AutoRefresh_Videos_CheckMode     = $false

    # -- INTERACTION --
    # What happens when you press ENTER on a video prompt?
    # "DONE" = Mark as processed (save to history).
    # "SKIP" = Ignore (will show again next time).
    DefaultEnterAction  = "DONE"
    
    # If regex finds 0 keywords, should we ask the user to type them?
    ManualSearchOnEmpty = $true
    
    # -- MATCHING RULES --
    ReprocessHistory    = $false
    VideoExtensions     = @("*.mp4", "*.mkv", "*.avi", "*.wmv")
    ExcludedPaths       = @("Recycle.Bin", "System Volume Information", "temp", "_UNPACK_", "Sample", "Trailers", "Promos")
    
    # Words to ignore during keyword extraction (Noise)
    StopWords           = @("the", "and", "with", "for", "to", "in", "on", "at", "of", "a", "an", "is", "by", "my", "your", "full", "what", "do", "unknown", "u", "c", "d", "e", "f", "v", "x")
    
    # Numbers to ignore (Resolutions, years, bitrates)
    IgnoredNumbers      = @("180", "190", "200", "220", "1080", "1920", "2048", "2160", "2700", "2880", "3840", "4096", "5400", "5760", "6144", "7680", "8192", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026")
    
    # Regex to scrub from filenames before analysis
    CleanUpPatterns     = @(
        "EPORNER\.COM.*", "sexlikereal\.com.*", "\beporner\.com\b", "\bvrporn\b", "\bbadoink\b",
        "\boculus(rift)?\w*\b", "\bhtc\w*\b", "\bvive\w*\b", "\bindex\b", "\bgearvr\b", "\bcardboard\b", "\bpico\b", "\bpsvr\b", "\bquest\d*\b", "\bsteamvr\b", "\bdaydream\b",
        "\bfisheye\d*\b", "\bbarrel\b", "\bequirectangular\b", "\bcube\w*\b", "\b180x180\b", "\bFOV\d*\b", "\b3dh\b", "\b3dv\b",
        "\b\d+kvr\d+\b", "\b\d+k\w+\b", "\b\d+p\w+\b",
        "\boriginals\b", "\boriginal\b", "\balpha\b", "\bbeta\b", "\bdemo\b", "\bpreview\b", "\btrailer\b", "\bsample\b", "\bproduction\b", "\bproductions\b", "\bvr\b", "\bvideos\b", "\bmovies\b", "\bscenes?\b", "\bparts?\b",
        "\b\d{3,4}x\d{3,4}\b", "\b\d{3,4}p\b", "\b\d{1,2}k\b", "uhd", "fhd", "hd", "hq", "ghq", "sbs", "tb", "ou", "lr", "rl",
        "\bx264\b", "\bx265\b", "\bh264\b", "\bh265\b", "\bhevc\b", "\bav1\b", "\bvp9\b", "\baac\b", "\bmp4\b", "\bmkv\b", "\bavi\b", "\bwmv\b",
        "\b\d+fps\b", "\b\d+bit\b", "\b\d+mbps\b",
        "\bremastered\b", "\bremaster\b", "\bupscaled\b", "\bai_upscale\b", "\bfiles\b", "\bbackup\b"
    )
}

# --- DEFAULT STUDIOS (Regex Database) ---
$DefaultStudios = @{
    "CzechVR"        = @("(?i)\bczech\s*vr\b", "(?i)\bczech\s*casting", "(?i)\bczech\s*network");
    "CzechVRFetish"  = @("(?i)\bczech\s*vr\s*fetish");
    "BabesVR"        = @("(?i)\bbabes\s*vr");
    "SexBabesVR"     = @("(?i)sex\s*babes");
    "WankzVR"        = @("(?i)\bwankz");
    "BaDoinkVR"      = @("(?i)\bbadoink");
    "NaughtyAmerica" = @("(?i)naughty\s*america", "(?i)nam\s*vr");
    "SexLikeReal"    = @("(?i)sex\s*like\s*real", "(?i)^slr\b", "(?i)[_-]slr\b");
    "VRedging"       = @("(?i)vredging");
    "VRCosplayX"     = @("(?i)vr\s*cosplay", "(?i)cosplay\s*vr");
    "VRConk"         = @("(?i)vr\s*conk");
    "RealityLovers"  = @("(?i)reality\s*lovers?");
    "VirtualTaboo"   = @("(?i)virtual\s*taboo");
    "VRBangers"      = @("(?i)vr\s*bangers");
    "VRSpy"          = @("(?i)vr\s*spy");
    "18VR"           = @("(?i)18\s*vr");
    "RealJamVR"      = @("(?i)real\s*jam");
    "AdultTime"      = @("(?i)adult\s*time");
    "EvilAngel"      = @("(?i)evil\s*angel");
    "DarkRoomVR"     = @("(?i)dark\s*room");
    "WetFood"        = @("(?i)wet\s*food");
    "SwallowBay"     = @("(?i)swallow\s*bay");
    "DDFNetwork"     = @("(?i)ddf", "(?i)ddf\s*network");
    "KinkVR"         = @("(?i)kink\s*vr");
    "LethalHardcoreVR" = @("(?i)lethal\s*hardcore");
    "VirtualRealPorn"  = @("(?i)virtual\s*real");
    "SinsVR"         = @("(?i)sins\s*vr", "(?i)xsins");
    "PornStarVR"     = @("(?i)pornstar\s*vr");
    "VRLatina"       = @("(?i)vr\s*latina");
    "HoloGirlsVR"    = @("(?i)hologirls");
    "POVR"           = @("(?i)povr");
    "MilfVR"         = @("(?i)milf\s*vr");
    "StripzVR"       = @("(?i)stripz");
    "TMWVR"          = @("(?i)tmw\s*vr");
    "JAV"            = @("(?i)jav\s*vr", "(?i)japan");
    "Holodexxx"      = @("(?i)holo\s*dexxx");
}

# ==============================================================================================
#   3. HELPER FUNCTIONS
# ==============================================================================================

# Loads the config file. Includes AUTO-REPAIR to add missing keys for new versions.
function Load-Configuration {
    if (-not (Test-Path $ConfigFile)) {
        Write-Host "Creating default config: $ConfigFile" -ForegroundColor Yellow
        $DefaultConfig | ConvertTo-Json -Depth 3 | Out-File $ConfigFile -Encoding utf8
        return $DefaultConfig
    }
    try {
        $LoadedConfig = Get-Content $ConfigFile -Raw -Encoding utf8 | ConvertFrom-Json
        $ConfigChanged = $false
        foreach ($Key in $DefaultConfig.Keys) {
            if ($null -eq $LoadedConfig.$Key) {
                $LoadedConfig | Add-Member -MemberType NoteProperty -Name $Key -Value $DefaultConfig[$Key]
                $ConfigChanged = $true
            }
        }
        if ($ConfigChanged) {
            Write-Host " [INFO] Updated config file with new settings." -ForegroundColor DarkGray
            $LoadedConfig | ConvertTo-Json -Depth 4 | Out-File $ConfigFile -Encoding utf8
        }
        return $LoadedConfig
    } catch {
        Write-Warning "Failed to load config file. Using defaults."
        return $DefaultConfig
    }
}

function Load-StudioDatabase {
    if (-not (Test-Path $StudioFile)) {
        $DefaultStudios | ConvertTo-Json -Depth 3 | Out-File $StudioFile -Encoding utf8
        return $DefaultStudios
    }
    try {
        $json = Get-Content $StudioFile -Raw -Encoding utf8 | ConvertFrom-Json
        $hash = @{}
        foreach ($prop in $json.PSObject.Properties) { $hash[$prop.Name] = $prop.Value }
        return $hash
    } catch { return $DefaultStudios }
}

function Save-StudioDatabase {
    param($Data)
    $Data | ConvertTo-Json -Depth 3 | Out-File $StudioFile -Encoding utf8
    Write-Host "   [DB] Studio database updated!" -ForegroundColor Green
}

# Highlights parts of a string (keywords) in Green
function Write-ColorizedString {
    param ([string]$Text, [array]$Highlights, [string]$BaseColor = "Gray", [string]$HighlightColor = "Green")
    if ($Highlights.Count -eq 0) { Write-Host $Text -ForegroundColor $BaseColor -NoNewline; return }
    $Pattern = "(?i)(" + (($Highlights | ForEach-Object { [regex]::Escape($_) }) -join "|") + ")"
    $Parts = [regex]::Split($Text, $Pattern)
    foreach ($part in $Parts) {
        if ([string]::IsNullOrEmpty($part)) { continue }
        if ($part -match $Pattern) { Write-Host $part -ForegroundColor $HighlightColor -NoNewline } 
        else { Write-Host $part -ForegroundColor $BaseColor -NoNewline }
    }
}

# Extracts date (YY.MM.DD or YYYY-MM-DD)
function Get-DateFromText {
    param ([string]$Text)
    if ($Text -match "\b(20\d{2})[.-](\d{2})[.-](\d{2})\b") { return @{ Y=$Matches[1]; M=$Matches[2]; D=$Matches[3] } }
    if ($Text -match "\b(\d{2})[.-](\d{2})[.-](\d{2})\b") { return @{ Y="20"+$Matches[1]; M=$Matches[2]; D=$Matches[3] } }
    if ($Text -match "\b(20\d{2})(\d{2})(\d{2})\b") { return @{ Y=$Matches[1]; M=$Matches[2]; D=$Matches[3] } }
    return $null
}

# Determines if we need to refresh the library cache
function Should-RefreshLibrary {
    param ( [string]$Name, [bool]$AutoRefresh, [bool]$ForceParam, [string]$DbPath )
    if ($ForceParam) { return $true }
    if (-not (Test-Path $DbPath)) { Write-Host "[$Name] Database missing. Refresh required." -ForegroundColor DarkGray; return $true }
    if ($AutoRefresh) { return $true }
    return $false
}

# Analyzes filename to extract metadata
function Get-SmartKeywords {
    param ([string]$Name, [string]$FullPath, [hashtable]$StudioDB, [object]$Config)
    
    $searchTerms = @()
    $detectedStudio = $null
    
    $ParentFolder = ""
    if (-not [string]::IsNullOrEmpty($FullPath)) {
        $ParentFolder = Split-Path (Split-Path $FullPath -Parent) -Leaf
        if ($FullPath -match "^[a-zA-Z]:\\") {
            if ($ParentFolder -match "^[a-zA-Z]:$") { $ParentFolder = "" }
        }
    }
    
    $CombinedText = "$ParentFolder $Name"

    # Studio Detection
    foreach ($studioName in $StudioDB.Keys) {
        $patterns = $StudioDB[$studioName]
        foreach ($pattern in $patterns) {
            if ($CombinedText -match $pattern) {
                $detectedStudio = $studioName
                $searchTerms += $detectedStudio
                if ($CombinedText -match "$pattern[-_ ]+(\d+)") {
                    $searchTerms += $Matches[$Matches.Count - 1]
                } elseif ($CombinedText -match "(\d+)[-_ ]+$pattern") {
                    $searchTerms += $Matches[1]
                }
                break
            }
        }
        if ($detectedStudio) { break }
    }

    # Date Detection
    $DateInfo = Get-DateFromText -Text $CombinedText
    $DateString = $null
    if ($DateInfo) {
        $YY = $DateInfo.Y.Substring(2,2)
        $DateString = "$YY$($DateInfo.M)$($DateInfo.D)"
    }

    # Cleaning
    $CombinedText = $CombinedText -replace "[._-]", " " 
    $CombinedText = $CombinedText -replace "^[a-zA-Z]:\\?", " "

    foreach ($pattern in $Config.CleanUpPatterns) {
        $CombinedText = $CombinedText -replace "(?i)$pattern", " "
    }
    
    # Tokenization
    $rawWords = $CombinedText.Split(" ") | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
    
    foreach ($word in $rawWords) {
        $w = $word.Trim()
        if ($detectedStudio -and $w -match "(?i)$detectedStudio") { continue }
        if ($Config.StopWords -contains $w.ToLower()) { continue }
        if ($w -match "^\d+$") {
            if ($Config.IgnoredNumbers -contains $w) { continue } 
            if ($w.Length -eq 1) { continue } 
        }
        if ($w.Length -le 2 -and $w -notmatch "^\d+$") { continue } 
        $searchTerms += $w
    }

    return @{ Keywords = ($searchTerms | Select-Object -Unique); Studio = $detectedStudio; DateParams = $DateInfo; DateCompact = $DateString }
}

# ==============================================================================================
#   4. MAIN EXECUTION FLOW
# ==============================================================================================

Write-Host "--- FUNSCRIPT MATCHER v2.1 ---" -ForegroundColor Cyan

# A. LOAD CONFIGURATION
$Config = Load-Configuration
$StudioMap = Load-StudioDatabase
Write-Host "Loaded configuration & $($StudioMap.Count) studios." -ForegroundColor Gray

# B. MODE DETECTION
$IsCheckMode = -not [string]::IsNullOrEmpty($CheckFunscript)
$IsTargetMode = -not [string]::IsNullOrEmpty($MatchVideo)
$IsSingleTaskMode = $IsCheckMode -or $IsTargetMode

# C. REFRESH LOGIC (Based on Mode)
if ($IsSingleTaskMode) {
    # Check/Target Mode uses separate config to avoid unnecessary waiting
    $DoRefreshFunscripts = Should-RefreshLibrary -Name "Funscripts" -AutoRefresh $Config.AutoRefresh_Funscripts_CheckMode -ForceParam $RefreshFunscripts -DbPath $LibScriptFile
    $DoRefreshVideos     = Should-RefreshLibrary -Name "Videos"     -AutoRefresh $Config.AutoRefresh_Videos_CheckMode     -ForceParam $RefreshVideos     -DbPath $LibVideoFile
} else {
    # Normal Bulk Mode
    $DoRefreshFunscripts = Should-RefreshLibrary -Name "Funscripts" -AutoRefresh $Config.AutoRefresh_Funscripts_Normal -ForceParam $RefreshFunscripts -DbPath $LibScriptFile
    $DoRefreshVideos     = Should-RefreshLibrary -Name "Videos"     -AutoRefresh $Config.AutoRefresh_Videos_Normal     -ForceParam $RefreshVideos     -DbPath $LibVideoFile
}

# D. LOAD/REFRESH FUNSCRIPTS
$ScriptLibrary = @()
if ($DoRefreshFunscripts) {
    Write-Host "Indexing FUNSCRIPTS in: $FunscriptsPath" -ForegroundColor Yellow
    if (-not (Test-Path $FunscriptsPath)) { Write-Error "Funscripts Path not found!"; exit }
    $Files = Get-ChildItem -Path $FunscriptsPath -Recurse -Include "*.funscript" -File
    $ScriptLibrary = $Files | Select-Object Name, FullName, @{Name="CleanName"; Expression={$_.Name -replace "[^a-zA-Z0-9]", ""}}, @{Name="Parent"; Expression={Split-Path (Split-Path $_.FullName -Parent) -Leaf}}
    $ScriptLibrary | ConvertTo-Json -Depth 2 -Compress | Out-File -FilePath $LibScriptFile -Encoding utf8
    Write-Host "Funscript DB saved! ($($ScriptLibrary.Count) scripts)" -ForegroundColor Green
} else {
    $ScriptLibrary = Get-Content $LibScriptFile -Encoding utf8 | ConvertFrom-Json
    Write-Host "Funscript DB loaded ($($ScriptLibrary.Count) scripts)." -ForegroundColor Green
}

# E. EXECUTE CHECK MODE (If Active)
if ($IsCheckMode) {
    Write-Host "`n[SEARCH MODE] Query: $CheckFunscript" -ForegroundColor Magenta
    $SmartData = Get-SmartKeywords -Name $CheckFunscript -FullPath "" -StudioDB $StudioMap -Config $Config
    $Keywords = $SmartData.Keywords
    $StudioHint = $SmartData.Studio
    $VideoDate = $SmartData.DateCompact
    
    if ($StudioHint) { Write-Host "Detected Studio: $StudioHint" -ForegroundColor Magenta }
    Write-Host "Detected Keywords: $($Keywords -join ', ')" -ForegroundColor DarkGray
    
    $Candidates = @()
    foreach ($Script in $ScriptLibrary) {
        $MatchScore = 0
        $MatchesFound = @()
        $sName = $Script.Name
        $sNameClean = $Script.CleanName
        
        if ($VideoDate) {
            $sNums = $sName -replace "[^0-9]", ""
            if ($sNums -like "*$VideoDate*") { 
                $MatchScore += 10 
                $MatchesFound += "DATE:$VideoDate"
            }
        }

        foreach ($k in $Keywords) {
            if ($sName -match "(?i)[^a-z0-9]" + [regex]::Escape($k) + "[^a-z0-9]" -or $sName -match "(?i)^" + [regex]::Escape($k)) { 
                $MatchScore += 2 
                $MatchesFound += $k
            } elseif ($sName -match "(?i)" + [regex]::Escape($k)) { 
                $MatchScore += 1 
                $MatchesFound += "$k(part)"
            } elseif ($sNameClean -match "(?i)" + [regex]::Escape(($k -replace " ", ""))) {
                if ($k.Length -lt 5) {
                    $MatchScore += 1
                    $MatchesFound += "$k(clean-weak)"
                } else {
                    $MatchScore += 3
                    $MatchesFound += "$k(clean-strong)"
                }
            }
        }
        
        if ($StudioHint -and $sName -match "(?i)$StudioHint") { 
            $MatchScore += 5 
            $MatchesFound += "STUDIO:$StudioHint"
        }

        if ($MatchScore -ge 2) {
            $Candidates += [PSCustomObject]@{ 
                Name = $Script.Name; 
                Score = $MatchScore; 
                Parent = $Script.Parent; 
                MatchedTerms = ($MatchesFound -join ", ") 
            }
        }
    }
    
    $Candidates = $Candidates | Sort-Object Score -Descending | Select-Object -First 10
    
    Write-Host "`nTop 10 Matches:" -ForegroundColor Yellow
    foreach ($c in $Candidates) {
        Write-Host "   [$($c.Score)] " -NoNewline -ForegroundColor Yellow
        Write-ColorizedString -Text $c.Name -Highlights $Keywords -BaseColor "White" -HighlightColor "Green"
        Write-Host " [Folder: $($c.Parent)]" -NoNewline -ForegroundColor DarkGray
        Write-Host " ($($c.MatchedTerms))" -ForegroundColor DarkGray
    }
    exit
}

# F. LOAD/REFRESH VIDEO LIBRARY
$VideoLibrary = @()
if ($DoRefreshVideos) {
    Write-Host "Scanning VIDEOS in: $VideoPath" -ForegroundColor Yellow
    $AllVideos = Get-ChildItem -Path $VideoPath -Recurse -Include $Config.VideoExtensions -File -ErrorAction SilentlyContinue
    
    $VideoLibrary = $AllVideos | Where-Object {
        $vPath = $_.FullName
        $shouldExclude = $false
        foreach ($ex in $Config.ExcludedPaths) {
            $cleanEx = $ex.Trim('\').Trim('/').Trim()
            $escapedEx = [regex]::Escape($cleanEx)
            if ($vPath -match "(?i)[\\/]$escapedEx[\\/]" -or $vPath -match "(?i)[\\/]$escapedEx$") { 
                $shouldExclude = $true; break 
            }
        }
        -not $shouldExclude
    } | Select-Object Name, FullName, BaseName, DirectoryName
    
    $VideoLibrary | ConvertTo-Json -Depth 2 -Compress | Out-File -FilePath $LibVideoFile -Encoding utf8
    Write-Host "Video DB saved! ($($VideoLibrary.Count) videos)" -ForegroundColor Green
} else {
    if (-not (Test-Path $LibVideoFile)) { Write-Error "Video DB missing. Run with -RefreshVideos"; exit }
    $VideoLibrary = Get-Content $LibVideoFile -Encoding utf8 | ConvertFrom-Json
    Write-Host "Video DB loaded ($($VideoLibrary.Count) videos)." -ForegroundColor Green
}

# G. HISTORY MANAGEMENT
$ProcessedSet = New-Object System.Collections.Generic.HashSet[string]
if (Test-Path $HistoryFile) { Get-Content $HistoryFile | ForEach-Object { if ($_.Trim()) { [void]$ProcessedSet.Add($_.Trim()) } } }
else { New-Item -Path $HistoryFile -ItemType File -Force | Out-Null }

$ShouldReprocessHistory = $ReprocessHistory -or $Config.ReprocessHistory
if ($ShouldReprocessHistory) { Write-Host "WARNING: Reprocessing History is ENABLED." -ForegroundColor Magenta }

# --- TARGET MODE (v1.9 / v2.0) ---
if ($IsTargetMode) {
    Write-Host "`n[TARGET MODE] Looking for video matching: '$MatchVideo'" -ForegroundColor Magenta
    
    $QueryParts = $MatchVideo -split "\s+" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    
    $RankedVideos = $VideoLibrary | ForEach-Object {
        $score = 0
        $name = $_.Name
        foreach ($part in $QueryParts) {
            if ($name -match "(?i)[^a-z0-9]" + [regex]::Escape($part) + "[^a-z0-9]" -or $name -match "(?i)" + [regex]::Escape($part)) {
                $score++
            }
        }
        [PSCustomObject]@{ Video = $_; Score = $score }
    } | Sort-Object Score -Descending | Select-Object -First 1

    if ($null -eq $RankedVideos -or $RankedVideos.Score -eq 0) {
        Write-Host "   -> No video found matching that query." -ForegroundColor Red
        exit
    }

    $Target = $RankedVideos.Video
    Write-Host "   Candidate found: " -NoNewline
    Write-Host $Target.Name -ForegroundColor Yellow
    
    $confirm = Read-Host "   Is this the video you want to match? (Y/n)"
    if ($confirm -match "^n") { Write-Host "   -> Aborted." -ForegroundColor Gray; exit }
    
    $VideoLibrary = @($Target)
    $ShouldReprocessHistory = $true 
    Write-Host "   -> Locked on target. Starting match..." -ForegroundColor Green
}

# ==============================================================================================
#   5. MAIN PROCESSING LOOP
# ==============================================================================================
$TotalCount = $VideoLibrary.Count
$CurrentIndex = 0

foreach ($VideoObj in $VideoLibrary) {
    $CurrentIndex++
    $VideoPath = $VideoObj.FullName
    $VideoName = $VideoObj.BaseName
    
    # 1. Cache Validation (Check if file still exists)
    if (-not (Test-Path -LiteralPath $VideoPath)) {
        Write-Host "[$CurrentIndex / $TotalCount] Skipped: File not found (Stale Cache)" -ForegroundColor DarkGray
        continue
    }

    try { $DestPath = [System.IO.Path]::ChangeExtension($VideoPath, ".funscript") } catch { continue }
    
    # 2. PRIORITY SKIP: SCRIPT ALREADY EXISTS
    if (Test-Path -LiteralPath $DestPath) { 
        Write-Host "[$CurrentIndex / $TotalCount] Skipped: $VideoName (Script file already exists)" -ForegroundColor DarkGray
        if (-not $ProcessedSet.Contains($VideoPath)) { Add-Content -Path $HistoryFile -Value $VideoPath }
        continue 
    }
    
    # 3. SKIP REASON: IN HISTORY
    $IsReProcessing = $false
    if ($ProcessedSet.Contains($VideoPath)) {
        if (-not $ShouldReprocessHistory) {
            Write-Host "[$CurrentIndex / $TotalCount] Skipped: $VideoName (Already in History - Marked as DONE)" -ForegroundColor DarkGray
            continue 
        } else {
            $IsReProcessing = $true
        }
    }

    # >> ANALYZE <<
    $SmartData = Get-SmartKeywords -Name $VideoName -FullPath $VideoPath -StudioDB $StudioMap -Config $Config
    
    # 4. MANUAL OVERRIDE CHECK (If empty)
    if ($SmartData.Keywords.Count -eq 0 -and $null -eq $SmartData.DateCompact) {
        if ($Config.ManualSearchOnEmpty) {
            Write-Host "`n[$CurrentIndex / $TotalCount] $VideoName" -ForegroundColor Cyan
            if ($IsReProcessing) { Write-Host "   (Re-processing: Was in History)" -ForegroundColor Cyan }
            Write-Host "   [!] No keywords found." -ForegroundColor Yellow
            $UserQuery = Read-Host "   Enter Scene Name / Keywords manually (or Enter to skip)"
            
            if (-not [string]::IsNullOrWhiteSpace($UserQuery)) {
                $SmartData = Get-SmartKeywords -Name $UserQuery -FullPath "" -StudioDB $StudioMap -Config $Config
            }
        }
    }

    $Keywords = $SmartData.Keywords
    $StudioHint = $SmartData.Studio
    $VideoDate = $SmartData.DateCompact 

    # 5. FINAL SKIP: NO KEYWORDS
    if ($Keywords.Count -eq 0 -and $null -eq $VideoDate) {
        Write-Host "[$CurrentIndex / $TotalCount] Skipped: $VideoName (No keywords detected)" -ForegroundColor DarkGray
        Add-Content -Path $HistoryFile -Value $VideoPath
        continue
    }

    # ==========================
    #  INTERACTIVE REFINING LOOP
    # ==========================
    $Searching = $true
    while ($Searching) {
        Write-Host "`n[$CurrentIndex / $TotalCount] Analyzing: $VideoName" -ForegroundColor Cyan
        if ($IsReProcessing) { Write-Host "   (Re-processing: Was in History)" -ForegroundColor Cyan }
        
        Write-Host "   Path: " -NoNewline -ForegroundColor DarkGray
        $PathHighlights = $Keywords
        if ($VideoDate) { $PathHighlights += $VideoDate }
        Write-ColorizedString -Text $VideoPath -Highlights $PathHighlights -BaseColor "DarkGray" -HighlightColor "Green"
        Write-Host ""
        if ($StudioHint) { Write-Host "   [STUDIO DETECTED]: $StudioHint" -ForegroundColor Magenta }
        if ($Keywords.Count -gt 0) { Write-Host "   [KEYWORDS]: $($Keywords -join ', ')" -ForegroundColor DarkGray }

        # >> MATCHING ALGORITHM <<
        $Candidates = @()
        foreach ($Script in $ScriptLibrary) {
            $MatchScore = 0
            $sName = $Script.Name
            $sNameClean = $Script.CleanName
            
            # Date Match (+10)
            if ($VideoDate) {
                $sNums = $sName -replace "[^0-9]", ""
                if ($sNums -like "*$VideoDate*") { $MatchScore += 10 }
            }

            # Keyword Match
            foreach ($k in $Keywords) {
                if ($sName -match "(?i)[^a-z0-9]" + [regex]::Escape($k) + "[^a-z0-9]" -or $sName -match "(?i)^" + [regex]::Escape($k)) { 
                    $MatchScore += 2 # Exact Word
                } elseif ($sName -match "(?i)" + [regex]::Escape($k)) { 
                    $MatchScore += 1 # Partial
                } elseif ($sNameClean -match "(?i)" + [regex]::Escape(($k -replace " ", ""))) {
                    # Weighted Clean Match
                    if ($k.Length -lt 5) {
                        $MatchScore += 1 # Weak
                    } else {
                        $MatchScore += 3 # Strong
                    }
                }
            }
            # Studio Match (+5)
            if ($StudioHint -and $sName -match "(?i)$StudioHint") { $MatchScore += 5 }

            # Inferred Studio Logic
            if ($MatchScore -ge 2) {
                $InferredStudio = $null
                foreach ($sKey in $StudioMap.Keys) {
                    if ($sName -match "(?i)^$sKey" -or $sName -match "(?i)[-_ ]$sKey[-_ ]") { 
                        $InferredStudio = $sKey; break 
                    }
                }
                if (-not $InferredStudio -and $sName -match "^([a-zA-Z0-9]+)\s*[-_]") {
                    $InferredStudio = $Matches[1] 
                }

                $Candidates += [PSCustomObject]@{ 
                    ScriptFullName = $Script.FullName; 
                    Score = $MatchScore; 
                    Name = $Script.Name; 
                    Parent = $Script.Parent;
                    HiddenStudio = $InferredStudio 
                }
            }
        }
        
        $Candidates = $Candidates | Sort-Object Score -Descending

        if ($Candidates.Count -eq 0) {
            Write-Host "   -> No match found." -ForegroundColor Red
        }

        # >> DISPLAY <<
        Write-Host "   Select match:" -ForegroundColor Yellow
        $i = 1
        $DisplayLimit = [math]::Min($Candidates.Count, 12)
        for ($x = 0; $x -lt $DisplayLimit; $x++) {
            $c = $Candidates[$x]
            Write-Host "   [$i] (Score: $($c.Score)) " -NoNewline -ForegroundColor White
            $CandHighlights = $Keywords
            if ($VideoDate) { $CandHighlights += $VideoDate }
            Write-ColorizedString -Text $c.Name -Highlights $CandHighlights -BaseColor "Gray" -HighlightColor "Green"
            Write-Host " [Folder: $($c.Parent)]" -ForegroundColor DarkGray
            $i++
        }

        # >> PROMPT <<
        $DefaultAction = if ($Config.DefaultEnterAction -eq "SKIP") { "SKIP" } else { "DONE" }
        Write-Host "   [Enter] = $DefaultAction / [S] = SKIP / [Text] = ADD KEYWORDS" -ForegroundColor DarkGray
        
        $choice = Read-Host "   Option"
        
        # CHOICE 1: Default Action
        if ([string]::IsNullOrWhiteSpace($choice)) {
            if ($Config.DefaultEnterAction -eq "SKIP") {
                Write-Host "   -> Skipped (Default)." -ForegroundColor Gray
            } else {
                if (-not $ProcessedSet.Contains($VideoPath)) { Add-Content -Path $HistoryFile -Value $VideoPath }
                Write-Host "   -> Done (Saved to history)." -ForegroundColor DarkGray
            }
            $Searching = $false
        }
        # CHOICE 2: Skip
        elseif ($choice -eq "s") {
            Write-Host "   -> Skipped." -ForegroundColor Gray
            $Searching = $false
        }
        # CHOICE 3: Selection
        elseif ($choice -match "^\d+$" -and [int]$choice -le $DisplayLimit) {
            $SelectedCandidate = $Candidates[[int]$choice - 1]
            try {
                [System.IO.File]::Copy($SelectedCandidate.ScriptFullName, $DestPath, $true)
                Write-Host "   -> OK." -ForegroundColor Green
                if (-not $ProcessedSet.Contains($VideoPath)) { Add-Content -Path $HistoryFile -Value $VideoPath }

                # Smart Learn Logic
                if (-not $StudioHint -and $SelectedCandidate.HiddenStudio) {
                    $CandidateStudio = $SelectedCandidate.HiddenStudio
                    if ($VideoName -match "(?i)$CandidateStudio") {
                        Write-Host "`n   [SMART LEARN]" -ForegroundColor Magenta
                        Write-Host "   Matched '$CandidateStudio', but didn't recognize it." -ForegroundColor Gray
                        $Learn = Read-Host "   Add '$CandidateStudio' to DB? (y/n)"
                        if ($Learn -eq "y") {
                            $NewPattern = "(?i)$CandidateStudio"
                            if (-not $StudioMap[$CandidateStudio]) { $StudioMap[$CandidateStudio] = @() }
                            $StudioMap[$CandidateStudio] += $NewPattern
                            Save-StudioDatabase -Data $StudioMap
                        }
                    }
                }
            } catch {
                Write-Host "   -> ERROR: Failed to copy file." -ForegroundColor Red
            }
            $Searching = $false
        }
        # CHOICE 4: Refine Keywords
        else {
            $ExtraKeywords = $choice -split "\s+" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $Keywords += $ExtraKeywords
            Write-Host "   [REFINE] Adding keywords: $($ExtraKeywords -join ', ')" -ForegroundColor Yellow
        }
    }
}
Write-Host "`n--- ALL DONE ---" -ForegroundColor Cyan
if (-not $CheckFunscript) { Pause }