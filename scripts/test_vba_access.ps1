# Quick VBA Code Sampling Script
# This attempts to get basic VBA information without full export

try {
    Write-Host "=== C&E Farming VBA Analysis ==="
    Write-Host ""
    
    # Test if Excel is available
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "✓ Excel COM object created successfully"
    
    # Test with one file first
    $testFile = 'c:\Users\garyu\projects\logistics-legacy-refactor\logistics-legacy-refactor\legacy\C&E Farming\VBA Functions\Trip Sheets.xlsm'
    
    if (Test-Path $testFile) {
        Write-Host "✓ Test file found: Trip Sheets.xlsm"
        
        try {
            $wb = $excel.Workbooks.Open($testFile, 0, $true, $null, "", "", $true)
            Write-Host "✓ Workbook opened successfully"
            
            # Get basic workbook info
            Write-Host "  Worksheets: $($wb.Worksheets.Count)"
            foreach ($ws in $wb.Worksheets) {
                Write-Host "    - $($ws.Name)"
            }
            
            # Try to access VBA project
            try {
                $vbproj = $wb.VBProject
                Write-Host "✓ VBA Project accessed"
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
            }
            catch {
                Write-Host "✗ VBA Project access denied: $($_.Exception.Message)"
                Write-Host "  This likely means 'Trust access to the VBA project object model' is disabled"
            }
            
            $wb.Close($false)
        }
        catch {
            Write-Host "✗ Error opening workbook: $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "✗ Test file not found: $testFile"
    }
    
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "✓ Excel application closed"
    
} catch {
    Write-Host "✗ Failed to create Excel COM object: $($_.Exception.Message)"
    Write-Host "  Make sure Excel is installed and accessible"
}

Write-Host ""
Write-Host "=== Analysis Complete ==="
