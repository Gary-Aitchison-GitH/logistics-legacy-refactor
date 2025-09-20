# Simple VBA export for Trip Sheets
try {
    $filePath = "c:\Users\garyu\projects\logistics-legacy-refactor\logistics-legacy-refactor\legacy\C&E Farming\VBA Functions\Trip Sheets.xlsm"
    $outputDir = "c:\Users\garyu\projects\logistics-legacy-refactor\logistics-legacy-refactor\legacy\exports\TripSheets"
    
    Write-Host "Checking file existence..."
    if (Test-Path $filePath) {
        Write-Host "File found: $filePath"
        Write-Host "File size: $([math]::Round((Get-Item $filePath).Length / 1KB, 2)) KB"
        
        New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
        
        Write-Host "Creating Excel application..."
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        Write-Host "Opening workbook..."
        $wb = $excel.Workbooks.Open($filePath)
        
        Write-Host "Accessing VBA project..."
        $vbproj = $wb.VBProject
        Write-Host "Found $($vbproj.VBComponents.Count) VBA components"
        
        foreach ($comp in $vbproj.VBComponents) {
            $typeName = switch ($comp.Type) {
                1 { "Standard Module (.bas)" }
                2 { "Class Module (.cls)" }
                3 { "UserForm (.frm)" }
                100 { "Document Module (.cls)" }
                default { "Unknown Type $($comp.Type) (.txt)" }
            }
            
            Write-Host "  Component: $($comp.Name) - $typeName"
            
            $ext = switch ($comp.Type) {
                1 { ".bas" }
                2 { ".cls" }
                3 { ".frm" }
                100 { ".cls" }
                default { ".txt" }
            }
            
            $fileName = ($comp.Name -replace '[^\w\-.]', '_') + $ext
            $fullOutputPath = Join-Path $outputDir $fileName
            
            Write-Host "    Exporting to: $fullOutputPath"
            $comp.Export($fullOutputPath)
        }
        
        Write-Host "Closing workbook..."
        $wb.Close($false)
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        Write-Host "SUCCESS: VBA export completed!"
        Write-Host "Output directory: $outputDir"
        
        # List exported files
        Write-Host "Exported files:"
        Get-ChildItem $outputDir | ForEach-Object { Write-Host "  - $($_.Name)" }
        
    } else {
        Write-Host "ERROR: File not found: $filePath"
    }
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    Write-Host "Stack trace: $($_.ScriptStackTrace)"
}
