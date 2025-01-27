function Build-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject[]]$AllObjects,  # color-coded, one table per object
        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]$ServiceNow,   # single table, embed link in 'Name' if 'Link' is present
        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]$Tasks,        # single table, embed link in 'Name'
        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]$Patching,     # single table, plain text
        [string]$Description,
        [string]$FooterText
    )

    # We rely on $Script:Config.Thresholds being loaded earlier in the script
    # e.g. $Script:Config = ConvertFrom-Json (Get-Content "C:\Path\Thresholds.json" -Raw)

    #----------------------------------------------------------------
    # Helper: Evaluate numeric conditions like "<50", ">=100 && <200",
    # or "!=189". Return one of: "none", "green", "yellow", or "red".
    #----------------------------------------------------------------
    function Get-CellClass {
        param(
            [string]$Environment,
            [string]$PropName,
            [string]$PropValue
        )

        # default no color
        $cellClass = "none"

        # fetch the array of rules for this environment
        $envRules = $Script:Config.Thresholds[$Environment]
        if (-not $envRules) {
            return $cellClass  # environment not in JSON => no color
        }

        # find rule for this property
        $ruleBlock = $envRules | Where-Object { $_.PropertyName -eq $PropName }
        if (-not $ruleBlock) {
            return $cellClass  # no rule => no color
        }

        # parse numeric for conditions
        $numericVal = ($PropValue -replace '[^0-9\.]', '')
        [double]$num = 0
        [double]::TryParse($numericVal, [ref]$num) | Out-Null

        function Test-Condition {
            param($number, $cond)
            $expr = $cond
                .Replace(">=", "$number -ge ")
                .Replace("<=", "$number -le ")
                .Replace(">",  "$number -gt ")
                .Replace("<",  "$number -lt ")
                .Replace("==", "$number -eq ")
                .Replace("!=", "$number -ne ")
                .Replace("&&", "-and")
                .Replace("||", "-or")
            try {
                [ScriptBlock]::Create($expr).Invoke() -eq $true
            }
            catch {
                $false
            }
        }

        # Check Green => Yellow => Red if present
        if ($ruleBlock.Green) {
            if (Test-Condition $num ($ruleBlock.Green)) {
                $cellClass = "green"
            }
        }
        if ($ruleBlock.Yellow -and $cellClass -eq "none") {
            if (Test-Condition $num ($ruleBlock.Yellow)) {
                $cellClass = "yellow"
            }
        }
        if ($ruleBlock.Red -and $cellClass -eq "none") {
            if ($ruleBlock.Red -eq "olderThan7Days") {
                # For date logic, parse $PropValue as [datetime], compare to (Get-Date).AddDays(-7), etc.
                # e.g.:
                # try {
                #     $dateVal = [datetime]$PropValue
                #     if ($dateVal -lt (Get-Date).AddDays(-7)) { $cellClass = "red" }
                # }
            }
            elseif (Test-Condition $num ($ruleBlock.Red)) {
                $cellClass = "red"
            }
        }

        return $cellClass
    }

    #----------------------------------------------------------------
    # Single Table: embed link logic for $ServiceNow, $Tasks, etc.
    # One table, multiple rows (no environment color logic).
    # If you want color logic here too, you can adapt the same approach.
    #----------------------------------------------------------------
    function Build-SingleTableEmbedLink {
        param(
            [PSCustomObject[]]$Data,
            [string]$Heading
        )

        if (!$Data -or $Data.Count -eq 0) {
            return ""
        }

        # gather all properties
        $allProps = $Data | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique

        # remove "Link" if you want to hide that column, or do an approach that merges link into 'Name'
        $allProps = $allProps | Where-Object { $_ -notin @('Link') }

        $htmlSnippet = @"
<div style="width:100%">
  <h1>$Heading</h1>
</div>
<div class="table-container" style="width:100%">
  <table>
    <tr>
"@
        foreach ($p in $allProps) {
            $htmlSnippet += "<th>$p</th>"
        }
        $htmlSnippet += "</tr>"

        foreach ($row in $Data) {
            $htmlSnippet += "<tr>"

            # store link
            $hasLink = $false
            $linkVal = $row.Link
            if ($linkVal -is [string] -and $linkVal -match '^https?://') {
                $hasLink = $true
            }

            foreach ($p in $allProps) {
                $val = $row.$p

                # if p is 'Name' and we have a link, embed
                if ($p -eq 'Name' -and $hasLink) {
                    $val = "<a href='$linkVal' target='_blank'>$val</a>"
                }

                $htmlSnippet += "<td>$val</td>"
            }

            $htmlSnippet += "</tr>"
        }

        $htmlSnippet += "</table></div>"
        return $htmlSnippet
    }

    #----------------------------------------------------------------
    # Single Table No Links: for $Patching
    #----------------------------------------------------------------
    function Build-SingleTableNoLinks {
        param(
            [PSCustomObject[]]$Data,
            [string]$Heading
        )

        if (!$Data -or $Data.Count -eq 0) {
            return ""
        }

        $allProps = $Data | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique

        $htmlSnippet = @"
<div style="width:100%">
  <h1>$Heading</h1>
</div>
<div class="table-container" style="width:100%">
  <table>
    <tr>
"@
        foreach ($p in $allProps) {
            $htmlSnippet += "<th>$p</th>"
        }
        $htmlSnippet += "</tr>"

        foreach ($row in $Data) {
            $htmlSnippet += "<tr>"
            foreach ($p in $allProps) {
                $val = $row.$p
                $htmlSnippet += "<td>$val</td>"
            }
            $htmlSnippet += "</tr>"
        }

        $htmlSnippet += "</table></div>"
        return $htmlSnippet
    }

    #----------------------------------------------------------------
    # Start building the final HTML
    # (One table per AllObjects item => 2-col layout)
    #----------------------------------------------------------------
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

    /* color classes */
    .none   { background-color: transparent; }
    .red    { background-color: #ffcccc; }
    .yellow { background-color: #ffffcc; }
    .green  { background-color: #ccffcc; }
</style>
</head>
<body>
<h1>Custom HTML Report</h1>
<pre>$Description</pre>
<div class="container">
"@

    #----------------------------------------------------------------
    # Build separate small tables for each object in $AllObjects
    # 2-column format: row => (PropertyName | Value)
    # With environment color logic
    #----------------------------------------------------------------
    foreach ($obj in $AllObjects) {
        # read environment
        $envName = $obj.Environment  # must match e.g. "CA2015", "PA", etc. in $Script:Config.Thresholds
        $tableHeading = $obj.Name    # displayed as <h2>

        $html += @"
<div class="table-container">
  <h2>$tableHeading</h2>
  <table>
"@

        foreach ($prop in $obj.PSObject.Properties) {
            if ($prop.Name -in @('Name','Environment')) { continue }

            $propName  = $prop.Name
            $propValue = $prop.Value

            # color logic
            $cellClass = Get-CellClass -Environment $envName -PropName $propName -PropValue $propValue

            $html += "<tr><td>$propName</td><td class='$cellClass'>$propValue</td></tr>"
        }

        $html += "</table></div>"
    }

    #----------------------------------------------------------------
    # Now handle ServiceNow, Tasks, Patching as single tables
    # (No environment color logic unless you want to add it)
    #----------------------------------------------------------------
    if ($ServiceNow) {
        $html += Build-SingleTableEmbedLink -Data $ServiceNow -Heading "ServiceNow"
    }
    if ($Tasks) {
        $html += Build-SingleTableEmbedLink -Data $Tasks -Heading "Tasks"
    }
    if ($Patching) {
        $html += Build-SingleTableNoLinks -Data $Patching -Heading "Patching"
    }

    $html += @"
</div>
<pre>$FooterText</pre>
</body>
</html>
"@

    # Write out
    $OutputPath = "D:\PowerShell\Test\CustomReport.html"
    $html | Out-File $OutputPath -Encoding utf8
    Write-Host "HTML report generated at $OutputPath"
    return $html
}
