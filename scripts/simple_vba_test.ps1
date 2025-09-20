Write-Host "=== C&E Farming VBA Analysis ==="
Write-Host ""

try {
    # Test if Excel is available
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "Excel COM object created successfully"
    
    # Test with one file first
    $testFile = 'c:\Users\garyu\projects\logistics-legacy-refactor\logistics-legacy-refactor\legacy\C&E Farming\VBA Functions\Trip Sheets.xlsm'
    
    if (Test-Path $testFile) {
        Write-Host "Test file found: Trip Sheets.xlsm"
        
        try {
            $wb = $excel.Workbooks.Open($testFile, 0, $true)
            Write-Host "Workbook opened successfully"
            
            # Get basic workbook info
            $wsCount = $wb.Worksheets.Count
            Write-Host "Worksheets: $wsCount"
            
            # Try to access VBA project
            try {
                $vbproj = $wb.VBProject
                Write-Host "VBA Project accessed"
                $compCount = $vbproj.VBComponents.Count
                Write-Host "VBA Components: $compCount"
                
                foreach ($comp in $vbproj.VBComponents) {
                    $compName = $comp.Name
                    $compType = $comp.Type
                    Write-Host "  - $compName (Type: $compType)"
                }
            }
            catch {
                Write-Host "VBA Project access denied - likely security setting"
            }
            
            $wb.Close($false)
        }
        catch {
            Write-Host "Error opening workbook"
        }
    }
    else {
        Write-Host "Test file not found"
    }
    
    $excel.Quit()
    Write-Host "Excel application closed"
    
} catch {
    Write-Host "Failed to create Excel COM object"
}

Write-Host ""
Write-Host "Analysis Complete"
