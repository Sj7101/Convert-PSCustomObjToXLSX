# Sample PSCustomObject array
$data = @(
    [PSCustomObject]@{Name="John"; Age=30; City="New York"},
    [PSCustomObject]@{Name="Jane"; Age=25; City="Los Angeles"},
    [PSCustomObject]@{Name="Doe"; Age=40; City="Chicago"}
)

# Function to export PSCustomObject to .xlsx
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

    # Write headers based on the PSCustomObject properties
    $properties = $Data[0].PSObject.Properties.Name
    for ($i = 0; $i -lt $properties.Count; $i++) {
        $worksheet.Cells.Item(1, $i + 1) = $properties[$i]
    }

    # Write data rows
    for ($row = 0; $row -lt $Data.Count; $row++) {
        for ($col = 0; $col -lt $properties.Count; $col++) {
            $worksheet.Cells.Item($row + 2, $col + 1) = $Data[$row].PSObject.Properties[$properties[$col]].Value
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

# Example usage
Export-CustomObjectToExcel -Data $data -OutputFile "C:\Path\To\Your\Output.xlsx"