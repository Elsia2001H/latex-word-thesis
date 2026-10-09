param(
    [Parameter(Mandatory=$true)][string]$PythonRoot,
    [Parameter(Mandatory=$true)][string]$PandocExe
)
$ErrorActionPreference = 'Stop'
$base = $PSScriptRoot
$sourcePython = (Resolve-Path -LiteralPath $PythonRoot).Path
$sourcePandoc = (Resolve-Path -LiteralPath $PandocExe).Path
$python = Join-Path $sourcePython 'python.exe'
if (-not (Test-Path -LiteralPath $python)) { throw 'PythonRoot must contain python.exe' }
& $python -c "import sys,struct,tkinter; assert sys.version_info[:2]==(3,12) and struct.calcsize('P')==8, 'Use Python 3.12 x64 with Tcl/Tk'"
if ($LASTEXITCODE -ne 0) { throw 'Python environment check failed' }
$runtime = Join-Path $base 'runtime'
$tools = Join-Path $base 'tools'
if (Test-Path -LiteralPath $runtime) { throw 'runtime already exists; use a fresh source checkout to build' }
New-Item -ItemType Directory -Path $runtime,$tools,(Join-Path $base 'output') -Force | Out-Null
Get-ChildItem -LiteralPath $sourcePython -File | Copy-Item -Destination $runtime
foreach ($name in @('DLLs','tcl')) {
    Copy-Item -LiteralPath (Join-Path $sourcePython $name) -Destination $runtime -Recurse
}
& robocopy (Join-Path $sourcePython 'Lib') (Join-Path $runtime 'Lib') /E /XD site-packages __pycache__ /NFL /NDL /NJH /NJS /NP
if ($LASTEXITCODE -ge 8) { throw 'Python library copy failed' }
& $python -m pip install --target (Join-Path $runtime 'Lib\site-packages') -r (Join-Path $base 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed' }
Copy-Item -LiteralPath $sourcePandoc -Destination (Join-Path $tools 'pandoc.exe')
Get-ChildItem -LiteralPath (Split-Path $sourcePandoc) -File |
    Where-Object {$_.Name -match 'COPYING|COPYRIGHT|LICENSE'} | Copy-Item -Destination $tools
# Isolate from global Python installations and user site packages.
Copy-Item -LiteralPath (Join-Path $base 'python312._pth.template') -Destination (Join-Path $runtime 'python312._pth')
& (Join-Path $runtime 'python.exe') (Join-Path $base 'program\launcher.py') --check
if ($LASTEXITCODE -ne 0) { throw 'Portable dependencies copied; install/configure TeX Live and desktop Word before use' }
Write-Output 'Ready. Double-click the launcher BAT file. Review third-party licenses before redistributing binaries.'
