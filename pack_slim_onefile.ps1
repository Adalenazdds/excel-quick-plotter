param(
	[string]$EntryScript = "main.py",
	[string]$AppName = "EXCEL-Quick-Plotter",
	[string]$VenvName = "build_env",
	[switch]$KeepVenv,
	[switch]$DryRun
)

$ErrorActionPreference = "Stop"

function Write-Step {
	param([string]$Message)
	Write-Host "`n[STEP] $Message" -ForegroundColor Cyan
}

function Write-Info {
	param([string]$Message)
	Write-Host "[INFO] $Message" -ForegroundColor Gray
}

function Write-WarnText {
	param([string]$Message)
	Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

if (-not (Test-Path $EntryScript)) {
	throw "Entry script not found: $EntryScript"
}

if (-not (Test-Path "icon.ico")) {
	throw "icon.ico not found in project root."
}

Write-Step "Collecting Python source files"
$pyFiles = Get-ChildItem -Path $projectRoot -Filter "*.py" -File | Where-Object {
	$_.Name -notmatch "^(pack_.*|setup|conftest)\.py$"
}

if (-not $pyFiles) {
	throw "No Python files found in project root."
}

Write-Step "Scanning imports to infer runtime dependencies"
$allImports = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$fullImports = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$stdlibAllow = @(
	"__future__","abc","argparse","array","ast","asyncio","base64","bisect","builtins","collections",
	"contextlib","copy","csv","ctypes","dataclasses","datetime","decimal","enum","functools","gc","glob",
	"hashlib","heapq","hmac","html","http","importlib","inspect","io","itertools","json","logging","math",
	"numbers","operator","os","pathlib","pickle","platform","pprint","queue","random","re","secrets","shlex",
	"shutil","signal","site","socket","sqlite3","statistics","string","subprocess","sys","tempfile","textwrap",
	"threading","time","traceback","types","typing","unittest","urllib","uuid","warnings","weakref","xml","zipfile"
)
$stdlibSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$stdlibAllow | ForEach-Object { [void]$stdlibSet.Add($_) }

$importRegex = '^(?:\s*from\s+([A-Za-z_][\w\.]*))|(?:\s*import\s+(.+))'

foreach ($file in $pyFiles) {
	$lines = Get-Content -Path $file.FullName
	foreach ($line in $lines) {
		if ($line -match '^\s*#') { continue }
		if ($line -notmatch $importRegex) { continue }

		if ($Matches[1]) {
			$full = $Matches[1].Trim()
			if ($full) { [void]$fullImports.Add($full) }
			$root = $full.Split('.')[0].Trim()
			if ($root) { [void]$allImports.Add($root) }
			continue
		}

		if ($Matches[2]) {
			$parts = $Matches[2] -split ','
			foreach ($part in $parts) {
				$token = ($part -split '\s+as\s+')[0].Trim()
				if (-not $token) { continue }
				[void]$fullImports.Add($token)
				$root = $token.Split('.')[0].Trim()
				if ($root) { [void]$allImports.Add($root) }
			}
		}
	}
}

$localModuleNames = Get-ChildItem -Path $projectRoot -Filter "*.py" -File |
	ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) }
$localSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$localModuleNames | ForEach-Object { [void]$localSet.Add($_) }

$pipMap = @{
	"pyqt5" = "PyQt5"
	"matplotlib" = "matplotlib"
	"pandas" = "pandas"
	"numpy" = "numpy"
	"seaborn" = "seaborn"
	"scipy" = "scipy"
	"xlwings" = "xlwings"
	"mplcursors" = "mplcursors"
	"keyboard" = "keyboard"
	"pynput" = "pynput"
	"pythoncom" = "pywin32"
	"win32com" = "pywin32"
	"win32api" = "pywin32"
	"win32gui" = "pywin32"
	"win32con" = "pywin32"
	"pywintypes" = "pywin32"
}

$installPkgs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($imp in $allImports) {
	if ($stdlibSet.Contains($imp)) { continue }
	if ($localSet.Contains($imp)) { continue }
	if ($pipMap.ContainsKey($imp.ToLowerInvariant())) {
		[void]$installPkgs.Add($pipMap[$imp.ToLowerInvariant()])
	}
}

[void]$installPkgs.Add("pyinstaller")

$installList = $installPkgs | Sort-Object
Write-Info ("Detected runtime pip packages: " + ($installList -join ", "))

Write-Step "Preparing clean venv: $VenvName"
$venvPath = Join-Path $projectRoot $VenvName
if (Test-Path $venvPath) {
	Write-Info "Removing existing venv: $venvPath"
	Remove-Item -Path $venvPath -Recurse -Force
}

python -m venv $VenvName
$pyExe = Join-Path $venvPath "Scripts\python.exe"
if (-not (Test-Path $pyExe)) {
	throw "Failed to create venv python executable: $pyExe"
}

& $pyExe -m pip install --upgrade pip setuptools wheel
if ($installList.Count -gt 0) {
	& $pyExe -m pip install @installList
}

Write-Step "Building dynamic exclude-module list"
# Remove obvious heavy modules that are not used by this project.
$excludeCandidates = @(
	"PyQt5.QtWebEngineCore",
	"PyQt5.QtWebEngineWidgets",
	"PyQt5.QtWebChannel",
	"PyQt5.QtNetwork",
	"PyQt5.QtQml",
	"PyQt5.QtQuick",
	"PyQt5.QtQuickWidgets",
	"PyQt5.QtSql",
	"PyQt5.QtTest",
	"PyQt5.QtMultimedia",
	"PyQt5.QtMultimediaWidgets",
	"PyQt5.QtWebSockets",
	"PyQt5.QtPositioning",
	"PyQt5.QtLocation",
	"PyQt5.QtBluetooth",
	"PyQt5.QtNfc",
	"PyQt5.QtSensors",
	"PyQt5.QtSerialPort",
	"PyQt5.QtTextToSpeech",
	"PyQt5.QtDesigner",
	"PyQt5.QtHelp",
	"PyQt5.QtPrintSupport",
	"PyQt5.QtOpenGL",
	"PyQt5.QtSvg",
	"PyQt5.QtXml",
	"PyQt5.QtXmlPatterns",
	"tkinter",
	"PySide2",
	"PySide6",
	"PyQt6",
	"IPython",
	"jupyter",
	"notebook",
	"pytest",
	"unittest",
	"matplotlib.backends.backend_tkagg",
	"matplotlib.backends.backend_tkcairo",
	"matplotlib.backends.backend_wx",
	"matplotlib.backends.backend_wxagg",
	"matplotlib.backends.backend_wxcairo",
	"matplotlib.backends.backend_gtk3",
	"matplotlib.backends.backend_gtk3agg",
	"matplotlib.backends.backend_gtk3cairo",
	"matplotlib.backends.backend_gtk4",
	"matplotlib.backends.backend_gtk4agg",
	"matplotlib.backends.backend_gtk4cairo",
	"matplotlib.backends.backend_macosx",
	"matplotlib.backends.backend_webagg",
	"matplotlib.backends.backend_webagg_core",
	"matplotlib.backends.backend_nbagg"
)

$excludeList = [System.Collections.Generic.List[string]]::new()
foreach ($cand in $excludeCandidates) {
	if ($fullImports.Contains($cand)) {
		continue
	}

	if ($cand -like "PyQt5.*") {
		$isUsedQtSubmodule = $false
		foreach ($used in $fullImports) {
			if ($used -like "PyQt5.*" -and $cand.StartsWith($used + ".", [System.StringComparison]::OrdinalIgnoreCase)) {
				$isUsedQtSubmodule = $true
				break
			}
		}
		if ($isUsedQtSubmodule) {
			continue
		}
	}

	$root = $cand.Split('.')[0]
	if ($allImports.Contains($root) -and $cand -notlike "matplotlib.backends.*" -and $cand -notlike "PyQt5.*") {
		continue
	}
	$excludeList.Add($cand)
}

Write-Info ("Exclude modules count: " + $excludeList.Count)

Write-Step "Resolving UPX path"
$upxCmd = Get-Command upx -ErrorAction SilentlyContinue
$upxDir = $null
if ($upxCmd) {
	$upxDir = Split-Path -Parent $upxCmd.Source
	Write-Info "UPX detected at: $($upxCmd.Source)"
} else {
	Write-WarnText "UPX not found in PATH. PyInstaller will still run, but compression may be reduced."
}

Write-Step "Running PyInstaller in onefile mode (-F)"
$pyiArgs = @(
	"-m", "PyInstaller",
	"-F",
	"--noconfirm",
	"--clean",
	"--noconsole",
	"--name", $AppName,
	"--icon", "icon.ico",
	"--add-data", "style.qss;.",
	"--hidden-import", "keyboard",
	"--hidden-import", "pynput",
	"--hidden-import", "pynput.keyboard",
	"--hidden-import", "pynput.keyboard._win32",
	"--hidden-import", "pythoncom",
	"--hidden-import", "pywintypes",
	"--hidden-import", "mplcursors"
)

if ($upxDir) {
	$pyiArgs += @("--upx-dir", $upxDir)
}

foreach ($ex in $excludeList) {
	$pyiArgs += @("--exclude-module", $ex)
}

$pyiArgs += $EntryScript

Write-Info ("PyInstaller command: `"$pyExe`" " + ($pyiArgs -join " "))

if (-not $DryRun) {
	& $pyExe @pyiArgs
}

Write-Step "Cleaning temporary packaging files"
$specPath = Join-Path $projectRoot ("$AppName.spec")
$buildPath = Join-Path $projectRoot "build"
if (Test-Path $specPath) {
	Remove-Item -Path $specPath -Force
	Write-Info "Removed spec file: $specPath"
}
if (Test-Path $buildPath) {
	Remove-Item -Path $buildPath -Recurse -Force
	Write-Info "Removed build directory: $buildPath"
}

if (-not $KeepVenv) {
	if (Test-Path $venvPath) {
		Remove-Item -Path $venvPath -Recurse -Force
		Write-Info "Removed temporary venv: $venvPath"
	}
} else {
	Write-Info "Keeping venv as requested: $venvPath"
}

$exePath = Join-Path $projectRoot ("dist\$AppName.exe")
Write-Host "`nBuild finished." -ForegroundColor Green
if (Test-Path $exePath) {
	Write-Host "EXE: $exePath" -ForegroundColor Green
} else {
	Write-WarnText "Expected EXE not found: $exePath"
}

Write-Host "Excluded heavy modules:" -ForegroundColor Cyan
$excludeList | ForEach-Object { Write-Host " - $_" }
