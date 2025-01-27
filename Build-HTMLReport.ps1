function Build-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject[]]
        $AllObjects,   # multiple object => multiple separate tables

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]
        $ServiceNow,   # single table for all items

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]
        $Patching,     # single table for all items

        [string]$Description,
        [string]$FooterText
    )

    #-----------------------------------------------------------------
    # Helper 1: Build multiple small tables (one per object)
    #           => For $AllObjects array
    #-----------------------------------------------------------------
    function Build-MultiObjectTables {
        param(
            [PSCustomObject[]]$ObjArray,
            [string]$SectionHeading
        )

        # If $ObjArray is null or empty, return nothing
        if (!$ObjArray -or $ObjArray.Count -eq 0) {
            return ""
        }

        $htmlBlock = @"
    <div style="width:100%">
        <h1>$SectionHeading</h1>
    </div>
"@

        # Each item => one table
        foreach ($obj in $ObjArray) {
            # Use the object's 'Name' property (if any) as a sub-heading
            $tableHeading = $obj.Name

            $htmlBlock += @"
    <div class="table-container">
        <h2>$tableHeading</h2>
        <table>
"@

            # Build a row for each property => (PropertyName | Value)
            foreach ($prop in $obj.PSObject.Properties) {
                if ($prop.Name -eq 'Name') {
                    # We already used Name as <h2>, skip repeating
                    continue
                }

                $propName  = $prop.Name
                $propValue = $prop.Value

                $htmlBlock += "<tr><td>$propName</td><td>$propValue</td></tr>"
            }
            $htmlBlock += "</table></div>"
        }

        return $htmlBlock
    }

    #-----------------------------------------------------------------
    # Helper 2: Build exactly ONE table from an array of PSCustomObjects
    #           => For $ServiceNow, $Patching, etc.
    #           => Each item => one row, each property => one column
    #-----------------------------------------------------------------
    function Build-SingleTable {
        param(
            [PSCustomObject[]]$Data,
            [string]$Heading
        )

        # If array is null/empty, return nothing
        if (!$Data -or $Data.Count -eq 0) {
            return ""
        }

        # Gather all property names for columns
        $allProps = $Data |
            ForEach-Object { $_.PSObject.Properties.Name } |
            Select-Object -Unique

        # Start the table
        $block = @"
    <div style="width:100%">
        <h1>$Heading</h1>
    </div>
    <div class="table-container" style="width:100%">
        <table>
            <tr>
"@
        # Headers
        foreach ($p in $allProps) {
            $block += "<th>$p</th>"
        }
        $block += "</tr>"

        # One row per object
        foreach ($obj in $Data) {
            $block += "<tr>"
            foreach ($p in $allProps) {
                $value = $obj.$p
                $block += "<td>$value</td>"
            }
            $block += "</tr>"
        }
        $block += "</table></div>"
        return $block
    }

    #-----------------------------------------------------------------
    # Start building the main HTML
    #-----------------------------------------------------------------
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Custom HTML Report</title>
    <style>
        body { font-family: Arial, sans-serif; }
        .container { display: flex; flex-wrap: wrap; }
        .table-container { width: 48%; margin: 1%; box-sizing: border-box; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid black; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        pre { white-space: pre-wrap; font-size: 16px; }
    </style>
</head>
<body>
<h1>Custom HTML Report</h1>

<pre>
$Description
</pre>

<div class="container">
"@

    # 1) $AllObjects => MULTIPLE tables, 2-wide layout
    $html += Build-MultiObjectTables -ObjArray $AllObjects -SectionHeading "All Objects"

    # 2) $ServiceNow => ONE table with multiple rows (if multiple items)
    if ($ServiceNow) {
        $html += Build-SingleTable -Data $ServiceNow -Heading "ServiceNow"
    }

    # 3) $Patching => ONE table with multiple rows
    if ($Patching) {
        $html += Build-SingleTable -Data $Patching -Heading "Patching"
    }

    #-----------------------------------------------------------------
    # Close container + add footer
    #-----------------------------------------------------------------
    $html += @"
</div>

<pre>
$FooterText
</pre>
</body>
</html>
"@

    # Write file & return HTML
    $OutputPath = "D:\PowerShell\Test\CustomReport.html"
    $html | Out-File -FilePath $OutputPath -Encoding utf8
    Write-Host "HTML report generated at $OutputPath"

    return $html
}
