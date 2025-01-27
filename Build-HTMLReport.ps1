function Build-HTMLReport {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]$AllObjects,   # One table

        [Parameter(Mandatory = $false)]
        [PSCustomObject[]]$ServiceNow,   # Another single table (multiple rows if multiple PSCustomObjects)

        [string]$Description,
        [string]$FooterText
    )

    #---------------------------------------------------------------------------
    # Helper: Builds one HTML table from an array of PSCustomObjects
    # (all items in one table). Each PSCustomObject => one row.
    #---------------------------------------------------------------------------
    function Build-OneTable {
        param(
            [PSCustomObject[]]$Data,
            [string]$Heading
        )

        # If array is empty or not provided, return nothing
        if (!$Data -or $Data.Count -eq 0) { 
            return ""
        }

        # Gather all property names from these objects for columns
        $allProps = $Data | ForEach-Object {
            $_.PSObject.Properties.Name
        } | Select-Object -Unique

        # Start the table HTML
        $tableHtml = @"
    <div class="table-container">
        <h2>$Heading</h2>
        <table>
            <tr>
"@

        # Table headers
        foreach ($propName in $allProps) {
            $tableHtml += "<th>$propName</th>"
        }
        $tableHtml += "</tr>"

        # One row per PSCustomObject in $Data
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

    # 1) One table for $AllObjects
    $html += Build-OneTable -Data $AllObjects -Heading "All Objects"

    # 2) One table for $ServiceNow – call it "Service Now Queue" as requested
    $html += Build-OneTable -Data $ServiceNow -Heading "Service Now Queue"

    #---------------------------------------------------------------------------
    # Close container + add footer + end HTML
    #---------------------------------------------------------------------------
    $html += @"
</div>

<pre>
$FooterText
</pre>

</body>
</html>
"@

    # Output to file & return the HTML
    $OutputPath = "D:\PowerShell\Test\CustomReport.html"
    $html | Out-File -FilePath $OutputPath -Encoding utf8
    Write-Host "HTML report generated at $OutputPath"

    return $html
}
