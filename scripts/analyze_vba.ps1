# Analyze VBA components in key Excel files
$repo = Split-Path -Parent $PSCommandPath
$legacy = Join-Path $repo "..\legacy"

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

Write-Host "=== VBA Analysis Report ==="
Write-Host ""

foreach ($file in $keyFiles) {
    $fullPath = Join-Path $legacy $file
    Write-Host "Analyzing: $file"
    
    if (Test-Path $fullPath) {
        Write-Host "  File exists: YES"
        Write-Host "  Size: $([math]::Round((Get-Item $fullPath).Length / 1KB, 2)) KB"
        
        # Try to get basic info without opening
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            
            $wb = $excel.Workbooks.Open($fullPath, 0, $true)  # Read-only
            $vbproj = $wb.VBProject
            
            Write-Host "  VBA Components: $($vbproj.VBComponents.Count)"
            foreach ($comp in $vbproj.VBComponents) {
                $typeName = switch ($comp.Type) {
                    1 { "Standard Module" }
                    2 { "Class Module" }
                    3 { "UserForm" }
                    100 { "Document Module" }
                    default { "Unknown ($($comp.Type))" }
                }
                Write-Host "    - $($comp.Name) [$typeName]"
            }
            
            $wb.Close($false)
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        catch {
            Write-Host "  ERROR: $($_.Exception.Message)"
        }
    } else {
        Write-Host "  File exists: NO"
    }
    Write-Host ""
}

Write-Host "=== Analysis Complete ==="
