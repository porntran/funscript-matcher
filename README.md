`Funscript Matcher` is a PowerShell utility that scans your local VR video collection and fuzzy-matches them against a massive library of funscripts (e.g., the Mega Pack). It handles messy filenames, detects studios automatically, and allows for interactive search refinement. This tool automates the process of matching local videos with a library of Funscripts. It uses fuzzy logic, regex pattern matching and user input to pair files even if filenames are not identical. Script copy the source Funscript to video destination, while renaming funscript to match the name of video file.

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
.\_funscript_matcher.ps1

Stop the script anytime by pressing CTRL + C

If you did not listen to point no. 1, you need to tell the script where your funscripts are. 
.\_funscript_matcher.ps1 -FunscriptsPath "X:\Path\To\MegaPack"

After first run, config files are created with default setup, see the details below.

🎮 Usage Guide
**1. The Standard Scan (Bulk Mode)**
To match all videos in the folder run:

.\_funscript_matcher.ps1

The script will iterate through every video.
[Enter]: Accept the default action (Mark as Done or Skip, configurable).
[Number]: Select a match (copies the script next to the video).
[S]: Skip this video.
[Text]: Type keywords (e.g. StudioVR 741) to refine the search.

**2. Target Mode**
If you just downloaded a video and want to match it immediately without scanning everything else:

_.\_funscript_matcher.ps1 -MatchVideo "studiovr 741"_

The script searches your video library for the best filename match.
It asks for confirmation.
It processes only that file (ignoring history/skip lists).

**3. Check Mode (Search Only)**
To see if a script exists in your library (without scanning videos):

_.\_funscript_matcher.ps1 -CheckFunscript "StudioVR Adele"_

Displays the top 10 matches with their scores and matched keywords.

⚙️ Configuration
On the first run, _funscript_matcher_config.json is created.
You can edit it:
AutoRefresh_Funscripts_Normal  false  If true, re-scans script folder on every startup.
AutoRefresh_Videos_Normal  true  If true, re-scans video folder on every startup (recommended).
DefaultEnterAction  "DONE"  What happens when you press Enter? "DONE" (save to history) or "SKIP".
ManualSearchOnEmpty  true  If no keywords are found, ask user to type them instead of skipping.
IgnoredNumbersList  ...  Numbers to ignore (resolutions like 1080, 2160, years like 2024).

StopWords  A list of common English function words (e.g., the, and, with, for).
They carry no unique meaning. Removing them prevents the script from matching unrelated files just because they both share a generic word like "The".

IgnoredNumbers  A specific blacklist of numbers representing technical specifications or dates (e.g., 1920, 4096, 180, 2024).
It forces the script to ignore resolutions, years, and Field of View (FOV) so it can focus on actual Scene IDs (like 741 or 01).

CleanUpPatterns  A collection of Regular Expressions (Regex) targeting technical jargon and branding.
It surgically removes website URLs, hardware names (Oculus, Vive), projection types (fisheye, 180x180), and codecs (hevc, h264) to leave only the clean scene name and studio.

You can modify other json files freely, ie. add studios manually, modify the regexp, add StopWords etc., json files are always prior over the default config if script.
**Sharing updated studios file and IgnoredNumbers,StopWords and CleanUpPatterns is uttermost welcome to help other!!**

🧠 How the Matching Works
The script assigns points to potential matches:
+10 Points: Date match (YYYY-MM-DD or YY.MM.DD). This is the strongest link.
+5 Points: Studio match (e.g. "StudioVR").
+3 Points: "Clean" match (Keywords found inside a string without spaces, e.g. "Flower" inside "FlowerQueen"). Only applies to words 5+ chars long.
+2 Points: Exact keyword match.
+1 Point: Partial keyword match.

❓ Troubleshooting
"File cannot be loaded because running scripts is disabled"
Run PowerShell and type:
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

"The script isn't finding my file!"
1. Check if the filename contains "junk" words defined in StopWords in the config.
2. Use the Interactive Refine: Type the keywords manually at the prompt.



NAME
    Funscript Matcher v2.1

SYNOPSIS
    .\_funscript_matcher.ps1 [[-VideoPath] <String>] [[-FunscriptsPath] <String>] 
    [-MatchVideo <String>] [-CheckFunscript <String>] 
    [-RefreshFunscripts] [-RefreshVideos] [-ReprocessHistory]

DESCRIPTION
    Automates the matching of VR video files to Funscripts (.funscript).
    Uses fuzzy logic, caching, and interactive user input.

PARAMETERS
    -VideoPath <String>
        Directory containing video files. Default: Current Directory.
    -FunscriptsPath <String>
        Directory containing .funscript files. Default: Current Directory.
    -MatchVideo <String>
        Searches the cached video list for a filename matching the string. Processes ONLY that video, ignoring the history file.
    -CheckFunscript <String>
        Searches the script library for the string and displays top 10 matches. Does not process any files.
    -RefreshFunscripts
        Forces a re-scan of the FunscriptsPath and updates the JSON cache.
    -RefreshVideos
        Forces a re-scan of the VideoPath and updates the JSON cache.
    -ReprocessHistory
        Ignores '_funscript_matcher_history.txt' and scans all videos again.
        (Note: Files that already have a .funscript next to them are always skipped).

EXAMPLES
    1. Standard Run:
       .\_funscript_matcher.ps1
    2. Setup Library:
       .\_funscript_matcher.ps1 -FunscriptsPath "X:\MegaPack" -RefreshFunscripts
    3. Match specific video:
       .\_funscript_matcher.ps1 -MatchVideo "sweet gwendoline"
    4. Search library:
       .\_funscript_matcher.ps1 -CheckFunscript "czechvr 123"
