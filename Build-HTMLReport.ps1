function Build-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject[]]
        $AllObjects,   # multiple objects => multiple separate tables (no hyperlink logic)

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]
        $ServiceNow,   # single table (hyperlink logic applies here only)

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]
        $Tasks,        # single table (no hyperlink logic)

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]
        $Patching,     # single table (no hyperlink logic)

        [string]$Description,
        [string]$FooterText
    )

    #-----------------------------------------------------------------
    # Helper 1: Build multiple small tables (one per object), 
    #           used for $AllObjects
    #           => no hyperlink logic here
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

        # Big heading for this section
        $htmlBlock = @"
    <div style="width:100%">
        <h1>$SectionHeading</h1>
    </div>
"@

        # Each PSCustomObject => one table
        foreach ($obj in $ObjArray) {
            # Use the object's 'Name' property (if any) as sub-heading
            $tableHeading = $obj.Name

            $htmlBlock += @"
    <div class="table-container">
        <h2>$tableHeading</h2>
        <table>
"@
            # Each property => one row (PropertyName | Value)
            foreach ($prop in $obj.PSObject.Properties) {
                if ($prop.Name -eq 'Name') {
                    continue  # skip repeating the Name property in the table
                }

                $propName  = $prop.Name
                $propValue = $prop.Value

                # NO hyperlink conversion here
                $htmlBlock += "<tr><td>$propName</td><td>$propValue</td></tr>"
            }
            $htmlBlock += "</table></div>"
        }

        return $htmlBlock
    }

    #-----------------------------------------------------------------
    # Helper 2: Build ONE table for single arrays like $Tasks, $Patching
    #           => NO hyperlink logic
    #-----------------------------------------------------------------
    function Build-SingleTableNoLinks {
        param(
            [PSCustomObject[]]$Data,
            [string]$Heading
        )

        if (!$Data -or $Data.Count -eq 0) {
            return ""
        }

        # Collect all property names for columns
        $allProps = $Data | ForEach-Object {
            $_.PSObject.Properties.Name
        } | Select-Object -Unique

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

        # One row per PSCustomObject
        foreach ($obj in $Data) {
            $block += "<tr>"
            foreach ($p in $allProps) {
                $value = $obj.$p
                # NO hyperlink conversion here
                $block += "<td>$value</td>"
            }
            $block += "</tr>"
        }
        $block += "</table></div>"
        return $block
    }

    #-----------------------------------------------------------------
    # Helper 3: Build ONE table specifically for $ServiceNow 
    #           => WITH hyperlink logic (Link/URL => <a> tag)
    #-----------------------------------------------------------------
    function Build-SingleTableServiceNow {
        param(
            [PSCustomObject[]]$Data,
            [string]$Heading
        )

        if (!$Data -or $Data.Count -eq 0) {
            return ""
        }

        # Collect all property names for columns
        $allProps = $Data | ForEach-Object {
            $_.PSObject.Properties.Name
        } | Select-Object -Unique

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

        # One row per PSCustomObject
        foreach ($obj in $Data) {
            $block += "<tr>"
            foreach ($p in $allProps) {
                $value = $obj.$p

                # ONLY here do we convert Link/URL => hyperlink
                if ($p -in @('Link','URL') -and $value -is [string] -and $value -match '^https?://') {
                    $value = "<a href='$value' target='_blank'>$value</a>"
                }
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

    # 1) Multiple small tables for $AllObjects
    $html += Build-MultiObjectTables -ObjArray $AllObjects -SectionHeading "All Objects"

    # 2) $ServiceNow => single table with hyperlink logic
    if ($ServiceNow) {
        $html += Build-SingleTableServiceNow -Data $ServiceNow -Heading "ServiceNow"
    }

    # 3) $Tasks => single table, no hyperlink logic
    if ($Tasks) {
        $html += Build-SingleTableNoLinks -Data $Tasks -Heading "Tasks"
    }

    # 4) $Patching => single table, no hyperlink logic
    if ($Patching) {
        $html += Build-SingleTableNoLinks -Data $Patching -Heading "Patching"
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

    $OutputPath = "D:\PowerShell\Test\CustomReport.html"
    $html | Out-File -FilePath $OutputPath -Encoding utf8

    Write-Host "HTML report generated at $OutputPath"
    return $html
}
