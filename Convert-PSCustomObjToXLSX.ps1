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

    # Resolve the output file path
    $outputFilePath = Resolve-Path $OutputFile
    Write-Host "Resolved output file path: $outputFilePath" -ForegroundColor Blue

    # Ensure the workbook is saved properly
    try {
        $workbook.SaveAs($outputFilePath.Path)
        Write-Host "Exported data to $OutputFile." -ForegroundColor Green
    } catch {
        Write-Error "Failed to save the Excel file: $_"
    } finally {
        # Close the workbook and quit Excel
        $workbook.Close($false)  # Ensure not to save changes on close
        $excel.Quit()

        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Example usage
$data = @(
    [PSCustomObject]@{Name="John"; Age=30; City="New York"; Property4="Value4"; Property5="Value5"; Property6="Value6"; Property7="Value7"; Property8="Value8"; Property9="Value9"; Property10="Value10"; Property11="Value11"},
    [PSCustomObject]@{Name="Jane"; Age=25; City="Los Angeles"; Property4="Value4"; Property5="Value5"; Property6="Value6"; Property7="Value7"; Property8="Value8"; Property9="Value9"; Property10="Value10"; Property11="Value11"},
    [PSCustomObject]@{Name="Doe"; Age=40; City="Chicago"; Property4="Value4"; Property5="Value5"; Property6="Value6"; Property7="Value7"; Property8="Value8"; Property9="Value9"; Property10="Value10"; Property11="Value11"}
)

Export-CustomObjectToExcel -Data $data -OutputFile "C:\Path\To\Your\Output.xlsx"
