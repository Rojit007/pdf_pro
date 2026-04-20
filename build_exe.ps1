param(
    [ValidateSet("onefile", "onedir")]
    [string]$Mode = "onefile"
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$entryScript = Join-Path $projectRoot "pdf_pro.py"
$requirements = Join-Path $projectRoot "requirements.txt"

if (-not (Test-Path $entryScript)) {
    throw "Entry script not found: $entryScript"
}
if (-not (Test-Path $requirements)) {
    throw "Requirements file not found: $requirements"
}

$pythonCmd = if (Get-Command py -ErrorAction SilentlyContinue) { "py" } else { "python" }

Write-Host "Installing dependencies and PyInstaller..."
& $pythonCmd -m pip install --upgrade pip
& $pythonCmd -m pip install -r $requirements
& $pythonCmd -m pip install pyinstaller

$pyiArgs = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--windowed",
    "--name", "PDFStudio",
    "--collect-submodules", "fitz",
    "--collect-all", "PIL",
    "--collect-all", "fitz",
    "--hidden-import", "tkinter",
    "--hidden-import", "tkinter.ttk",
    "--hidden-import", "tkinter.filedialog",
    "--hidden-import", "tkinter.messagebox",
    "--hidden-import", "tkinter.colorchooser",
    "--hidden-import", "tkinter.simpledialog"
)

if ($Mode -eq "onefile") {
    $pyiArgs += "--onefile"
}

$pyiArgs += $entryScript

Write-Host "Building executable in $Mode mode..."
& $pythonCmd @pyiArgs

if ($Mode -eq "onefile") {
    $outputPath = Join-Path $projectRoot "dist\PDFStudio.exe"
} else {
    $outputPath = Join-Path $projectRoot "dist\PDFStudio\PDFStudio.exe"
}

Write-Host "Build complete: $outputPath"
