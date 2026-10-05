# pdf_to_xls.ps1 — Convert a Raport kasowy PDF to Excel
#
# Usage:
#   .\pdf_to_xls.ps1                        # converts all PDFs in this folder
#   .\pdf_to_xls.ps1 "rk kwiecien.PDF"     # converts one specific file
#   .\pdf_to_xls.ps1 "rk kwiecien.PDF" "out.xlsx"

param(
    [string]$PdfPath   = "",
    [string]$XlsxPath  = ""
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PyScript  = Join-Path $ScriptDir "pdf_to_xls.py"

# Verify python script exists
if (-not (Test-Path $PyScript)) {
    Write-Error "Nie znaleziono skryptu: $PyScript"
    exit 1
}

# Verify python is available
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Error "Python nie jest zainstalowany lub nie jest w PATH."
    exit 1
}

# Collect list of PDF files to process
$files = @()
if ($PdfPath -ne "") {
    if (-not (Test-Path $PdfPath)) {
        Write-Error "Nie znaleziono pliku: $PdfPath"
        exit 1
    }
    $files = @($PdfPath)
} else {
    # No argument — process every PDF in the script's folder
    $files = Get-ChildItem -Path $ScriptDir -Filter "*.PDF" |
             Select-Object -ExpandProperty FullName
    if ($files.Count -eq 0) {
        Write-Error "Brak plikow PDF w katalogu: $ScriptDir"
        exit 1
    }
}

# Convert each file
foreach ($pdf in $files) {
    $xlsxArgs = @($PyScript, $pdf)
    if ($XlsxPath -ne "" -and $files.Count -eq 1) {
        $xlsxArgs += $XlsxPath
    }
    python @xlsxArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Blad przy konwersji: $pdf"
    }
}
