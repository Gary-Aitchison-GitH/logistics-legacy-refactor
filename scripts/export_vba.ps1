# Exports VBA components from every .xlsm in legacy/ (recursively)
$repo = Split-Path -Parent $PSCommandPath
$legacy = Join-Path $repo "legacy"
$exports = Join-Path $legacy "exports"
New-Item -ItemType Directory -Force -Path $exports | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Export-One([string]$wbPath) {
    $wb = $excel.Workbooks.Open($wbPath)
    try {
        $name = [IO.Path]::GetFileNameWithoutExtension($wbPath)
        $outDir = Join-Path $exports $name
        New-Item -ItemType Directory -Force -Path $outDir | Out-Null

        $vbproj = $wb.VBProject
        foreach ($comp in $vbproj.VBComponents) {
            switch ($comp.Type) {
                1 { $ext = ".bas" }   # Std module
                2 { $ext = ".cls" }   # Class module
                3 { $ext = ".frm" }   # UserForm
                100 { $ext = ".cls" } # Document module
                default { $ext = ".txt" }
            }
            $safe = ($comp.Name -replace '[^\w\-.]', '_') + $ext
            $comp.Export((Join-Path $outDir $safe))
        }
    }
    finally {
        $wb.Close($false)
    }
}

Get-ChildItem $legacy -Recurse -Filter *.xlsm -File | ForEach-Object { Export-One $_.FullName }

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
