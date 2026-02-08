# Funscript Matcher

**The ultimate tool for organizing your VR script collection.**

`Funscript Matcher` is a PowerShell utility that scans your local VR video collection and fuzzy-matches them against a massive library of funscripts (e.g., the Mega Pack). It handles messy filenames, detects studios automatically, and allows for interactive search refinement. The script copies the source Funscript to the video destination, renaming it to match the video file automatically.

---

## ✨ Key Features

* **🧠 Fuzzy Logic Scoring:** Matches files even if filenames aren't identical (e.g., `Video_01.mp4` matches `Studio - Video 01.funscript`).
* **📂 Library Caching:** Scans your massive script collection once and saves it to JSON for instant startup speeds.
* **🎯 Target Mode:** Want to match *one* specific video? Use `-MatchVideo` to find and process just that file instantly.
* **🔍 Interactive Refine:** Script didn't find the match? Type new keywords directly into the console to search again without restarting.
* **🤖 Smart Learn:** The tool learns new studio patterns from your manual matches.
* **⚡ Check Mode:** Verify if you have a script for a scene *before* you download the video.

---

## 🚀 Installation & Setup

### 1. Download
Download `_funscript_matcher.ps1` and place it **inside the folder where your VR videos are located** (e.g., `U:\VR Videos`).

### 2. First Run (Indexing)
Open PowerShell in the folder and run:
```powershell
.\_funscript_matcher.ps1

```

*Tip: Stop the script anytime by pressing `CTRL + C`.*

**First time setup:**
If you did not place the script in the same folder as your scripts, you need to tell the tool where your funscripts are located:

```powershell
.\_funscript_matcher.ps1 -FunscriptsPath "X:\Path\To\MegaPack"

```

*After the first run, configuration files are created automatically.*

---

## 🎮 Usage Guide

### 1. The Standard Scan (Bulk Mode)

To match all videos in the current folder, run:

```powershell
.\_funscript_matcher.ps1

```

The script will iterate through every video.

* **[Enter]**: Accept the default action (Mark as Done or Skip, configurable).
* **[Number]**: Select a match (copies the script next to the video).
* **[S]**: Skip this video.
* **[Text]**: Type keywords (e.g. `StudioVR 741`) to refine the search manually.

### 2. Target Mode (Sniper Mode)

If you just downloaded a video and want to match it immediately without scanning everything else:

```powershell
.\_funscript_matcher.ps1 -MatchVideo "studiovr 741"

```

* The script searches your video library for the best filename match.
* It asks for confirmation.
* It processes **only** that file (ignoring history/skip lists).

### 3. Check Mode (Search Only)

To see if a script exists in your library (without scanning videos):

```powershell
.\_funscript_matcher.ps1 -CheckFunscript "StudioVR Adele"

```

* Displays the top 10 matches with their scores and matched keywords.

---

## ⚙️ Configuration

On the first run, `_funscript_matcher_config.json` is created. You can edit it to customize behavior:

| Setting | Default | Description |
| --- | --- | --- |
| `AutoRefresh_Funscripts_Normal` | `false` | If true, re-scans script folder on every startup. |
| `AutoRefresh_Videos_Normal` | `true` | If true, re-scans video folder on every startup (recommended). |
| `DefaultEnterAction` | `"DONE"` | What happens when you press Enter? `"DONE"` (save to history) or `"SKIP"`. |
| `ManualSearchOnEmpty` | `true` | If no keywords are found, ask user to type them instead of skipping. |

### Advanced Filtering

You can modify the JSON files to improve accuracy. ***Sharing updated configuration files is highly encouraged!***

* **StopWords:** A list of common English function words (e.g., `the`, `and`, `with`, `for`). They carry no unique meaning. Removing them prevents the script from matching unrelated files just because they share a generic word like "The".
* **IgnoredNumbers:** A specific blacklist of numbers representing technical specifications or dates (e.g., `1920`, `4096`, `180`, `2024`). It forces the script to ignore resolutions, years, and Field of View (FOV) so it can focus on actual Scene IDs (like `741` or `01`).
* **CleanUpPatterns:** A collection of Regular Expressions (Regex) targeting technical jargon and branding. It surgically removes website URLs, hardware names (`Oculus`, `Vive`), projection types (`fisheye`, `180x180`), and codecs (`hevc`, `h264`) to leave only the clean scene name and studio.

*Note: JSON configuration files always take priority over the script's default settings.*

---

## 🧠 How the Matching Works

The script assigns points to potential matches to find the best candidate:

* **+10 Points:** Date match (YYYY-MM-DD or YY.MM.DD). This is the strongest link.
* **+5 Points:** Studio match (e.g. "StudioVR").
* **+3 Points:** "Clean" match (Keywords found inside a string without spaces, e.g. "Flower" inside "FlowerQueen"). *Only applies to words 5+ chars long.*
* **+2 Points:** Exact keyword match.
* **+1 Point:** Partial keyword match.

---

## ❓ Troubleshooting

**"File cannot be loaded because running scripts is disabled"**
Run PowerShell and type the following command to allow the script to run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

```

**"The script isn't finding my file!"**

1. Check if the filename contains "junk" words defined in `StopWords` in the config.
2. Use the **Interactive Refine**: Type the keywords manually at the prompt when the script asks.

---

## 📄 Command Reference

**SYNOPSIS**

```powershell
.\_funscript_matcher.ps1 [[-VideoPath] <String>] [[-FunscriptsPath] <String>] 
[-MatchVideo <String>] [-CheckFunscript <String>] 
[-RefreshFunscripts] [-RefreshVideos] [-ReprocessHistory]

```

**PARAMETERS**

* `-VideoPath`: Directory containing video files. *Default: Current Directory.*
* `-FunscriptsPath`: Directory containing .funscript files. *Default: Current Directory.*
* `-MatchVideo`: Searches the cached video list for a filename matching the string. Processes **ONLY** that video, ignoring the history file.
* `-CheckFunscript`: Searches the script library for the string and displays top 10 matches. Does not process any files.
* `-RefreshFunscripts`: Forces a re-scan of the FunscriptsPath and updates the JSON cache.
* `-RefreshVideos`: Forces a re-scan of the VideoPath and updates the JSON cache.
* `-ReprocessHistory`: Ignores `_funscript_matcher_history.txt` and scans all videos again. *(Note: Files that already have a .funscript next to them are always skipped).*

**EXAMPLES**

1. **Standard Run:**
`.\_funscript_matcher.ps1`
2. **Setup Library:**
`.\_funscript_matcher.ps1 -FunscriptsPath "X:\MegaPack" -RefreshFunscripts`
3. **Match specific video:**
`.\_funscript_matcher.ps1 -MatchVideo "sweet gwendoline"`
4. **Search library:**
`.\_funscript_matcher.ps1 -CheckFunscript "czechvr 123"`
