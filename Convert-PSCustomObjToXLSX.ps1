function Export-CustomObjectToExcel {
    param (
        [Parameter(Mandatory=$true)]
        [array]$Data,
        [string]$OutputFile = "Output.xlsx"
    )

    # Create a new Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Add a new workbook and get the first worksheet
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Sheets.Item(1)

    # Get the properties (columns) from the first item in the PSCustomObject array
    $properties = $Data[0].PSObject.Properties.Name

    # Write headers to the first row
    for ($i = 0; $i -lt $properties.Count; $i++) {
        $worksheet.Cells.Item(1, $i + 1) = $properties[$i]
    }

    # Write data rows starting from the second row
    for ($row = 0; $row -lt $Data.Count; $row++) {
        for ($col = 0; $col -lt $properties.Count; $col++) {
            $worksheet.Cells.Item($row + 2, $col + 1) = $Data[$row].$($properties[$col])
        }
    }

    # Save the workbook
    $workbook.SaveAs((Resolve-Path $OutputFile).Path)
    $workbook.Close()
    $excel.Quit()

    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host "Exported data to $OutputFile."
}