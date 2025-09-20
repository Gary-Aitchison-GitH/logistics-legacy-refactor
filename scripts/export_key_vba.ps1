# Export VBA from key administrative Excel files
$repo = Split-Path -Parent $PSCommandPath
$legacy = Join-Path $repo "..\legacy"
$exports = Join-Path $legacy "exports"
New-Item -ItemType Directory -Force -Path $exports | Out-Null

$keyFiles = @(
    "C&E Farming\VBA Functions\Trip Sheets.xlsm",
    "C&E Farming\VBA Functions\Accounts.xlsm", 
    "C&E Farming\VBA Functions\CD3 Admin.xlsm",
    "C&E Farming\VBA Functions\Finance Admin.xlsm",
    "C&E Farming\VBA Functions\Fleet Admin.xlsm",
    "C&E Farming\VBA Functions\Invoice.xlsm",
    "C&E Farming\VBA Functions\Schedules Admin.xlsm",
    "C&E Farming\VBA Functions\Summary Report.xlsm"
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Export-One([string]$wbPath) {
    Write-Host "Processing: $wbPath"
    
    if (-not (Test-Path $wbPath)) {
        Write-Host "File not found: $wbPath"
        return
    }
    
    try {
        $wb = $excel.Workbooks.Open($wbPath)
        try {
            $name = [IO.Path]::GetFileNameWithoutExtension($wbPath)
            $outDir = Join-Path $exports $name
            New-Item -ItemType Directory -Force -Path $outDir | Out-Null

            $vbproj = $wb.VBProject
            Write-Host "  Found $($vbproj.VBComponents.Count) VBA components"
            
            foreach ($comp in $vbproj.VBComponents) {
                switch ($comp.Type) {
                    1 { $ext = ".bas" }   # Std module
                    2 { $ext = ".cls" }   # Class module
                    3 { $ext = ".frm" }   # UserForm
                    100 { $ext = ".cls" } # Document module
                    default { $ext = ".txt" }
                }
                $safe = ($comp.Name -replace '[^\w\-.]', '_') + $ext
                $outPath = Join-Path $outDir $safe
                $comp.Export($outPath)
                Write-Host "    Exported: $safe (Type: $($comp.Type))"
            }
        }
        finally {
            $wb.Close($false)
        }
    }
    catch {
        Write-Host "Error processing $wbPath : $($_.Exception.Message)"
    }
}

foreach ($file in $keyFiles) {
    $fullPath = Join-Path $legacy $file
    Export-One $fullPath
}

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "VBA export complete!"
