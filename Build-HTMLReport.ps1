function Build-HTMLReport {
    param(
        [PSCustomObject[]]$AllObjects,        # The objects needing color-coded tables
        [string]$Description,
        [string]$FooterText
    )

    #----------------------------------------------------------------
    # Helper: Evaluate numeric conditions like "<50", ">=100 && <200",
    # or "!=189". Return "red"/"yellow"/"green"/"none".
    # (You can expand this for date logic, "olderThan7Days", etc.)
    #----------------------------------------------------------------
    function Get-CellClass {
        param(
            [string]$PropName,
            [string]$PropValue,
            [Object[]]$EnvRules
        )

        $cellClass = "none"  # default no color

        # 1) Find matching rule block
        $ruleBlock = $EnvRules | Where-Object { $_.PropertyName -eq $PropName }
        if (-not $ruleBlock) {
            return $cellClass  # no rule => no color
        }

        # 2) Possibly parse numeric. For date logic, you can do special checks:
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

        # Check in order: Green => Yellow => Red
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
            if (Test-Condition $num ($ruleBlock.Red)) {
                $cellClass = "red"
            }
        }

        return $cellClass
    }

    #----------------------------------------------------------------
    # Start building HTML
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
    # Group $AllObjects by "Environment" property
    #----------------------------------------------------------------
    $groups = $AllObjects | Group-Object -Property Environment

    foreach ($group in $groups) {
        $envName = $group.Name   # e.g. "CA2015"
        
        # Lookup that environment's rules from $Script:Config
        $envRules = $Script:Config.Thresholds[$envName]
        if (-not $envRules) {
            Write-Host "No threshold rules found for environment '$envName'"
        }

        # Build one table for all objects in this environment
        $html += "<div class='table-container'>"
        $html += "<h2>$envName</h2>"
        $html += "<table>"

        # Collect union of property names from all objects
        $allProps = $group.Group | ForEach-Object {
            $_.PSObject.Properties.Name
        } | Select-Object -Unique

        # Table header row
        $html += "<tr>"
        $html += "<th>Name</th>"  # if you want a specific Name column
        foreach ($p in $allProps) {
            if ($p -ne 'Environment' -and $p -ne 'Name') {
                $html += "<th>$p</th>"
            }
        }
        $html += "</tr>"

        # Table rows
        foreach ($obj in $group.Group) {
            $html += "<tr>"

            # Name cell
            $objName = $obj.Name
            $html += "<td>$objName</td>"

            # Other props
            foreach ($p in $allProps) {
                if ($p -eq 'Environment' -or $p -eq 'Name') { continue }
                $val = $obj.$p
                $colorClass = "none"

                # If we have thresholds for this environment, do color check
                if ($envRules) {
                    $colorClass = Get-CellClass -PropName $p -PropValue $val -EnvRules $envRules
                }

                $html += "<td class='$colorClass'>$val</td>"
            }

            $html += "</tr>"
        }

        $html += "</table></div>"
    }

    $html += @"
</div>
<pre>$FooterText</pre>
</body></html>
"@

    # Save or return HTML
    $OutputPath = "D:\PowerShell\Test\CustomReport.html"
    $html | Out-File -FilePath $OutputPath -Encoding utf8
    Write-Host "HTML report generated at $OutputPath"
    return $html
}
