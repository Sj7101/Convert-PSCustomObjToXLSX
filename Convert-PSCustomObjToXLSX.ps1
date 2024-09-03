function Export-CustomObjectToExcel {
    param (
        [Parameter(Mandatory=$true)]
        [array]$Data,
        [string]$OutputFile = "Output.xlsx"
    )

    # Get the properties of the PSCustomObject (only NoteProperties)
    $properties = ($Data[0] | Get-Member -MemberType NoteProperty).Name

    Write-Host "Properties detected:" -ForegroundColor Green
    $properties | ForEach-Object { Write-Host $_ }

    # Create a new Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Add a new workbook and get the first worksheet
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Sheets.Item(1)

    # Write headers to the first row
    for ($i = 0; $i -lt $properties.Count; $i++) {
        $header = $properties[$i]
        Write-Host "Writing header: $header" -ForegroundColor Cyan
        $worksheet.Cells.Item(1, $i + 1) = $header
    }

    # Write data rows starting from the second row
    for ($row = 0; $row -lt $Data.Count; $row++) {
        Write-Host "Processing row $($row + 1)" -ForegroundColor Yellow
        for ($col = 0; $col -lt $properties.Count; $col++) {
            $propertyName = $properties[$col]
            $value = $Data[$row].$propertyName
            
            # Handle potential null or empty values
            if ($null -eq $value) {
                $value = ""
            }

            Write-Host "Writing value: $value to row $($row + 2), column $($col + 1)" -ForegroundColor Magenta
            $worksheet.Cells.Item($row + 2, $col + 1) = $value
        }
    }

    # Save the workbook
    try {
        $workbook.SaveAs((Resolve-Path $OutputFile).Path)
        Write-Host "Exported data to $OutputFile." -ForegroundColor Green
    } catch {
        Write-Error "Failed to save the Excel file: $_"
    }

    $workbook.Close()
    $excel.Quit()

    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Export-CustomObjectToExcel -Data $data -OutputFile "C:\Path\To\Your\Output.xlsx"