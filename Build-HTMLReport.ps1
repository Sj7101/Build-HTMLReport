function Build-HTMLReport {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]
        $AllObjects,    # One table

        [Parameter(Mandatory = $false)]
        [PSCustomObject[]]
        $ServiceNow,    # Second table (2 PSCustomObjects => 2 rows)

        [string]$Description,
        [string]$FooterText
    )

    #---------------------------------------------------------------------------
    # Helper function: Builds one HTML table from an array of PSCustomObjects.
    # Each PSCustomObject becomes a row; each property is a column.
    # If the array is empty or $null, returns an empty string (no table).
    #---------------------------------------------------------------------------
    function Build-OneTable {
        param(
            [PSCustomObject[]]$Data,
            [string]$Heading
        )

        if (!$Data -or $Data.Count -eq 0) {
            return ""  # No table if array is empty or $null
        }

        # Collect all property names (columns) from these objects
        $allProps = $Data | ForEach-Object {
            $_.PSObject.Properties.Name
        } | Select-Object -Unique

        # Start building the table HTML
        $tableHtml = @"
    <div class="table-container">
        <h2>$Heading</h2>
        <table>
            <tr>
"@

        # Headers
        foreach ($propName in $allProps) {
            $tableHtml += "<th>$propName</th>"
        }
        $tableHtml += "</tr>"

        # One row per PSCustomObject
        foreach ($obj in $Data) {
            $tableHtml += "<tr>"
            foreach ($propName in $allProps) {
                $value = $obj.$propName
                $tableHtml += "<td>$value</td>"
            }
            $tableHtml += "</tr>"
        }

        $tableHtml += "</table></div>"
        return $tableHtml
    }

    #---------------------------------------------------------------------------
    # Begin the full HTML document
    #---------------------------------------------------------------------------
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Custom HTML Report</title>
    <style>
        body { font-family: Arial, sans-serif; }
        .container { display: flex; flex-wrap: wrap; }
        .table-container {
            width: 48%;
            margin: 1%;
            box-sizing: border-box;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }
        th, td {
            border: 1px solid black;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
        pre {
            white-space: pre-wrap;
            font-size: 16px;
        }
    </style>
</head>
<body>
<h1>Custom HTML Report</h1>

<pre>
$Description
</pre>

<div class="container">
"@

    # Build one table for $AllObjects
    $html += Build-OneTable -Data $AllObjects -Heading "All Objects"

    # Build one table for $ServiceNow
    $html += Build-OneTable -Data $ServiceNow -Heading "ServiceNow"

    # Close the container + add footer + end HTML
    $html += @"
</div>

<pre>
$FooterText
</pre>

</body>
</html>
"@

    #---------------------------------------------------------------------------
    # Write to file and return the HTML
    #---------------------------------------------------------------------------
    $OutputPath = "D:\PowerShell\Test\CustomReport.html"
    $html | Out-File -FilePath $OutputPath -Encoding utf8
    Write-Host "HTML report generated at $OutputPath"

    return $html
}
