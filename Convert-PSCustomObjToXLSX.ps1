function Export-CustomObjectToExcel {
    param (
        [Parameter(Mandatory=$true)]
        [array]$Data,
        [string]$OutputFile = "Output.xlsx"
    )

    $properties = ($Data[0] | Get-Member -MemberType NoteProperty).Name

    Write-Host "Properties detected:" -ForegroundColor Green
    $properties | ForEach-Object { Write-Host $_ }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false  # Set to true to manually inspect Excel if needed
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Sheets.Item(1)

    for ($i = 0; $i -lt $properties.Count; $i++) {
        $header = $properties[$i]
        Write-Host "Writing header: $header" -ForegroundColor Cyan
        $worksheet.Cells.Item(1, $i + 1) = $header
    }

    for ($row = 0; $row -lt $Data.Count; $row++) {
        Write-Host "Processing row $($row + 1)" -ForegroundColor Yellow
        for ($col = 0; $col -lt $properties.Count; $col++) {
            $propertyName = $properties[$col]
            $value = $Data[$row].$propertyName
            
            if ($null -eq $value) {
                $value = ""
            }

            Write-Host "Writing value: $value to row $($row + 2), column $($col + 1)" -ForegroundColor Magenta
            $worksheet.Cells.Item($row + 2, $col + 1) = $value
        }
    }

    # Force recalculation of the worksheet
    $worksheet.Calculate()

    # Save the workbook
    $outputFilePath = Resolve-Path $OutputFile
    try {
        $workbook.SaveAs($outputFilePath.Path, 51)  # Explicitly save as .xlsx format
        Write-Host "Exported data to $OutputFile." -ForegroundColor Green
    } catch {
        Write-Error "Failed to save the Excel file: $_"
    } finally {
        $workbook.Close($false)
        $excel.Quit()

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
