function Build-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject[]]
        $CustomObjects,   # multiple objects -> each gets its own table, 2-wide

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]
        $ServiceNow,      # single table with multiple rows (if >1 items)

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]
        $TASKS,           # single table

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]
        $PATCHING,        # single table

        [string]$Description,
        [string]$FooterText
    )

    #-----------------------------------------------------------------
    # Helper 1: Build multiple tables (one per object) for $CustomObjects
    #           => 2-wide layout using CSS flex + .table-container {width: 48%}
    #-----------------------------------------------------------------
    function Build-MultiObjectTables {
        param(
            [PSCustomObject[]]$ObjArray,
            [string]$SectionHeading
        )

        if (!$ObjArray -or $ObjArray.Count -eq 0) {
            return ""
        }

        $htmlBlock = @"
    <div style="width:100%">
        <h1>$SectionHeading</h1>
    </div>
"@

        foreach ($obj in $ObjArray) {
            # Use the 'Name' property as <h2> heading if present
            $tableHeading = $obj.Name

            $htmlBlock += @"
    <div class="table-container">
        <h2>$tableHeading</h2>
        <table>
"@
            # Each property -> one row of "PropertyName | Value"
            foreach ($prop in $obj.PSObject.Properties) {
                # Skip the 'Name' property if you don't want it repeated in the table
                if ($prop.Name -eq 'Name') { continue }

                $propName  = $prop.Name
                $propValue = $prop.Value

                # Example: Convert Link or URL properties into clickable text
                if ($propName -in @('Link','URL')) {
                    # Use the same object's Name as the link text
                    $linkText = if ($obj.PSObject.Properties.Name -contains 'Name') {
                        $obj.Name
                    } else {
                        $propValue  # fallback if no Name property
                    }
                    $propValue = "<a href='$propValue' target='_blank'>$linkText</a>"
                }

                $htmlBlock += "<tr><td>$propName</td><td>$propValue</td></tr>"
            }

            $htmlBlock += "</table></div>"
        }

        return $htmlBlock
    }

    #-----------------------------------------------------------------
    # Helper 2: Build ONE table from an array of PSCustomObjects
    #           => Each object is a row, each property is a column
    #           => Table heading from first object's TableName (if present),
    #              otherwise "Untitled Table"
    #-----------------------------------------------------------------
    function Build-SingleTable {
        param(
            [PSCustomObject[]]$ObjArray
        )

        if (!$ObjArray -or $ObjArray.Count -eq 0) {
            return ""
        }

        # Use the TableName property of the first object if it exists
        $firstObj = $ObjArray[0]
        $heading = if ($firstObj.PSObject.Properties.Name -contains 'TableName') {
            $firstObj.TableName
        } else {
            "Untitled Table"
        }

        # Gather all property names from all objects for columns
        $allProps = $ObjArray |
            ForEach-Object { $_.PSObject.Properties.Name } |
            Select-Object -Unique

        # Remove 'TableName' from columns if you don't want to show it
        $allProps = $allProps | Where-Object { $_ -ne 'TableName' }

        $htmlBlock = @"
    <div style="width:100%">
        <h1>$heading</h1>
    </div>
    <div class="table-container" style="width:100%">
        <table>
            <tr>
"@

        # Build <th> for each property
        foreach ($p in $allProps) {
            $htmlBlock += "<th>$p</th>"
        }
        $htmlBlock += "</tr>"

        # One row per PSCustomObject
        foreach ($obj in $ObjArray) {
            $htmlBlock += "<tr>"
            foreach ($p in $allProps) {
                $cellValue = $obj.$p

                # Convert Link or URL -> clickable anchor
                if ($p -in @('Link','URL') -and $cellValue) {
                    # If object has a Name property, use it as link text
                    $linkText = if ($obj.PSObject.Properties.Name -contains 'Name') {
                        $obj.Name
                    } else {
                        $cellValue
                    }
                    $cellValue = "<a href='$cellValue' target='_blank'>$linkText</a>"
                }

                $htmlBlock += "<td>$cellValue</td>"
            }
            $htmlBlock += "</tr>"
        }

        $htmlBlock += "</table></div>"
        return $htmlBlock
    }

    #-------------------------------
    # Start building the full HTML
    #-------------------------------
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Custom HTML Report</title>
    <style>
        body { font-family: Arial, sans-serif; }
        .container { display: flex; flex-wrap: wrap; }
        .table-container { width: 48%; margin: 1%; }
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

    # 1) Show the main $CustomObjects in multiple small tables (2-wide)
    $html += Build-MultiObjectTables -ObjArray $CustomObjects -SectionHeading "Main Items"

    # 2) One table for $ServiceNow
    if ($ServiceNow) {
        $html += Build-SingleTable -ObjArray $ServiceNow
    }

    # 3) One table for $TASKS
    if ($TASKS) {
        $html += Build-SingleTable -ObjArray $TASKS
    }

    # 4) One table for $PATCHING
    if ($PATCHING) {
        $html += Build-SingleTable -ObjArray $PATCHING
    }

    # Close container + Footer
    $html += @"
</div>
<pre>
$FooterText
</pre>
</body>
</html>
"@

    # Output the HTML to a file
    $OutputPath = "D:\PowerShell\Test\CustomReport.html"
    $html | Out-File -FilePath $OutputPath -Encoding utf8
    Write-Host "HTML report generated at $OutputPath"

    # Return the HTML for emailing, etc.
    return $html
}
