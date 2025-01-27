function Build-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject[]]
        $AllObjects,   # multiple objects => multiple separate tables (no hyperlink logic)

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]
        $ServiceNow,   # single table, with hyperlink logic that uses a custom link text

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
    # Helper 1: Build multiple small tables (one per object) for $AllObjects
    #           => NO hyperlink logic
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
            $tableHeading = $obj.Name
            $htmlBlock += @"
    <div class="table-container">
        <h2>$tableHeading</h2>
        <table>
"@
            foreach ($prop in $obj.PSObject.Properties) {
                if ($prop.Name -eq 'Name') { continue }
                $propName  = $prop.Name
                $propValue = $prop.Value

                # NO hyperlink logic here
                $htmlBlock += "<tr><td>$propName</td><td>$propValue</td></tr>"
            }
            $htmlBlock += "</table></div>"
        }
        return $htmlBlock
    }

    #-----------------------------------------------------------------
    # Helper 2: Build ONE table for $Tasks or $Patching
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

        $allProps = $Data | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique

        $block = @"
    <div style="width:100%">
        <h1>$Heading</h1>
    </div>
    <div class="table-container" style="width:100%">
        <table>
            <tr>
"@
        foreach ($p in $allProps) {
            $block += "<th>$p</th>"
        }
        $block += "</tr>"

        foreach ($obj in $Data) {
            $block += "<tr>"
            foreach ($p in $allProps) {
                $value = $obj.$p
                # NO hyperlink logic
                $block += "<td>$value</td>"
            }
            $block += "</tr>"
        }
        $block += "</table></div>"
        return $block
    }

    #-----------------------------------------------------------------
    # Helper 3: Build ONE table for $ServiceNow with hyperlink logic.
    #           => The link text is replaced with the object's Name property
    #              (or a further-trimmed version if you prefer).
    #-----------------------------------------------------------------
    function Build-SingleTableServiceNow {
        param(
            [PSCustomObject[]]$Data,
            [string]$Heading
        )

        if (!$Data -or $Data.Count -eq 0) {
            return ""
        }

        $allProps = $Data | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique

        $block = @"
    <div style="width:100%">
        <h1>$Heading</h1>
    </div>
    <div class="table-container" style="width:100%">
        <table>
            <tr>
"@
        foreach ($p in $allProps) {
            $block += "<th>$p</th>"
        }
        $block += "</tr>"

        foreach ($obj in $Data) {
            $block += "<tr>"
            foreach ($p in $allProps) {
                $value = $obj.$p

                # If the property is Link or URL, show link text = $obj.Name
                if ($p -in @('Link','URL') -and $value -is [string] -and $value -match '^https?://') {
                    
                    # Option A: Use full $obj.Name as link text
                    # $linkText = $obj.Name

                    # Option B: If you want to replace "Service Now " with "", do:
                    $linkText = $obj.Name -replace '^Service Now\s*', '' 
                    # => "Service Now Incident Queue" becomes "Incident Queue"
                    # => "Service Now Change Queue" => "Change Queue"

                    $value = "<a href='$value' target='_blank'>$linkText</a>"
                }
                $block += "<td>$value</td>"
            }
            $block += "</tr>"
        }
        $block += "</table></div>"
        return $block
    }

    #-----------------------------------------------------------------
    # Build the main HTML
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

    # 1) $AllObjects => multiple tables
    $html += Build-MultiObjectTables -ObjArray $AllObjects -SectionHeading "All Objects"

    # 2) $ServiceNow => single table with custom hyperlink logic
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
