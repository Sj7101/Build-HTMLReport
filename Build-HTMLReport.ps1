function Build-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject[]]$AllObjects,  # Each object => one table with 2 columns
        [string]$Description,
        [string]$FooterText
    )

    # Ensure $Script:Config.Thresholds is loaded earlier in the script

    #----------------------------------------------------------------
    # Helper for color logic. Checks $Script:Config.Thresholds[envName]
    # for a rule matching the property name, then sets .green/.yellow/.red
    # or .none if no match or condition fails.
    #----------------------------------------------------------------
    function Get-CellClass {
        param(
            [string]$Environment,
            [string]$PropName,
            [string]$PropValue
        )

        # Default => no color
        $cellClass = "none"

        # 1) Get environment's rule array
        $envRules = $Script:Config.Thresholds[$Environment]
        if (-not $envRules) {
            return $cellClass  # environment not in JSON => no color
        }

        # 2) Find a rule for this property
        $ruleBlock = $envRules | Where-Object { $_.PropertyName -eq $PropName }
        if (-not $ruleBlock) {
            return $cellClass  # no rule for this property => no color
        }

        # Possibly parse numeric logic (like "<50", "!=189"). For date logic
        # or custom "olderThan7Days", you can do special checks below.
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

        # Evaluate Green => Yellow => Red if defined
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
            # For special keys like "olderThan7Days", handle that logic here
            if ($ruleBlock.Red -eq "olderThan7Days") {
                # e.g. parse date & compare to (Get-Date).AddDays(-7)
                # but for now, ignoring date logic
            }
            elseif (Test-Condition $num ($ruleBlock.Red)) {
                $cellClass = "red"
            }
        }

        return $cellClass
    }

    #----------------------------------------------------------------
    # Start building the HTML
    # One .table-container per PSCustomObject => 2 col table
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

    # For each PSCustomObject => a separate table
    foreach ($obj in $AllObjects) {
        # 1) Retrieve the environment from the object
        $envName = $obj.Environment  # e.g. "CA2015", "PA", "CB", etc.

        # 2) We'll display the object's 'Name' property as the table heading
        $tableHeading = $obj.Name

        # Build the <table>
        $html += @"
<div class="table-container">
  <h2>$tableHeading</h2>
  <table>
"@

        # Row per property => 2 columns: PropName | Value
        foreach ($prop in $obj.PSObject.Properties) {
            # Skip the 'Name' & 'Environment' props from the table if you want
            if ($prop.Name -in @('Name','Environment')) { continue }

            $propName  = $prop.Name
            $propValue = $prop.Value
            # Check color
            $cellClass = Get-CellClass -Environment $envName -PropName $propName -PropValue $propValue

            # Build row
            $html += "<tr><td>$propName</td><td class='$cellClass'>$propValue</td></tr>"
        }

        $html += "</table></div>"
    }

    $html += @"
</div>
<pre>$FooterText</pre>
</body>
</html>
"@

    # Write file
    $OutputPath = "D:\PowerShell\Test\CustomReport.html"
    $html | Out-File $OutputPath -Encoding utf8
    Write-Host "HTML report generated at $OutputPath"
    return $html
}
