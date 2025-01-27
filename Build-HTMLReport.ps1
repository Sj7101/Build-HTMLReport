function Build-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject[]]$AllObjects,      # Color-coded, one table per object

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]$ServiceNow,       # Single table, embed link in 'Name'

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]$Tasks,            # Single table, embed link in 'Name'

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]$Patching,         # Single table, plain text

        [Parameter(Mandatory=$true)]
        [string]$Description,                # Passed as a parameter

        [Parameter(Mandatory=$true)]
        [string]$FooterText                  # Passed as a parameter
    )

    # Ensure $Script:Config.Thresholds and $Script:Config.email are loaded earlier in the script
    # Example:
    # $Script:Config = ConvertFrom-Json (Get-Content "C:\Path\Thresholds.json" -Raw)

    #----------------------------------------------------------------
    # Helper: Evaluate conditions and determine cell class
    #----------------------------------------------------------------
    function Get-CellClass {
        param(
            [string]$Environment,
            [string]$PropName,
            [string]$PropValue
        )

        # Default class
        $cellClass = "none"

        # Retrieve environment-specific rules
        $envRules = $null
        try {
            # Dynamically access the environment's threshold rules
            $envRules = $Script:Config.Thresholds | Select-Object -ExpandProperty $Environment -ErrorAction Stop
        }
        catch {
            Write-Host "DEBUG: No threshold rules found for environment '$Environment'"
            return $cellClass
        }

        # Find the rule for the specific property
        $ruleBlock = $envRules | Where-Object { $_.PropertyName -eq $PropName }
        if (-not $ruleBlock) {
            Write-Host "DEBUG: No rule found for PropertyName='$PropName' in environment='$Environment'"
            return $cellClass
        }

        # Handle special cases like date-based conditions
        if ($ruleBlock.Red -eq "olderThan7Days") {
            try {
                $dateVal = [datetime]$PropValue
                if ($dateVal -lt (Get-Date).AddDays(-7)) {
                    $cellClass = "red"
                }
            }
            catch {
                Write-Host "DEBUG: Failed to parse date for PropertyName='$PropName' with value='$PropValue'"
            }
            return $cellClass
        }

        # Attempt to parse numeric value
        $numericVal = ($PropValue -replace '[^0-9\.]', '')
        [double]$num = 0
        if (-not [double]::TryParse($numericVal, [ref]$num)) {
            Write-Host "DEBUG: Non-numeric value for PropertyName='$PropName': '$PropValue'"
            return $cellClass
        }

        # Function to safely evaluate condition strings
        function Test-Condition {
            param(
                [double]$number,
                [string]$cond
            )

            [string]$expr = $cond
            $expr = ($expr).Replace(">=", "$number -ge ")
            $expr = ($expr).Replace("<=", "$number -le ")
            $expr = ($expr).Replace(">",  "$number -gt ")
            $expr = ($expr).Replace("<",  "$number -lt ")
            $expr = ($expr).Replace("==", "$number -eq ")
            $expr = ($expr).Replace("!=", "$number -ne ")
            $expr = ($expr).Replace("&&", "-and")
            $expr = ($expr).Replace("||", "-or")

            try {
                return ([ScriptBlock]::Create($expr).Invoke() -eq $true)
            }
            catch {
                Write-Host "DEBUG: Failed to evaluate condition '$cond' for number '$number'"
                return $false
            }
        }

        # Check conditions in order: Green -> Yellow -> Red
        if ($ruleBlock.Green) {
            if (Test-Condition -number $num -cond $ruleBlock.Green) {
                $cellClass = "green"
            }
        }

        if ($ruleBlock.Yellow -and $cellClass -eq "none") {
            if (Test-Condition -number $num -cond $ruleBlock.Yellow) {
                $cellClass = "yellow"
            }
        }

        if ($ruleBlock.Red -and $cellClass -eq "none") {
            if (Test-Condition -number $num -cond $ruleBlock.Red) {
                $cellClass = "red"
            }
        }

        return $cellClass
    }

    #----------------------------------------------------------------
    # Helper: Build Single Table with Embedded Links
    #----------------------------------------------------------------
    function Build-SingleTableEmbedLink {
        param(
            [PSCustomObject[]]$Data,
            [string]$Heading
        )

        if (-not $Data -or $Data.Count -eq 0) {
            return ""
        }

        # Gather all property names
        $allProps = $Data | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique
        # Optionally exclude 'Link' if you want
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

            # Check if there's a valid link
            $linkVal = $row.Link
            $hasLink = ($linkVal -is [string] -and $linkVal -match '^https?://')

            foreach ($p in $allProps) {
                $val = $row.$p

                # Embed link in 'Name' if applicable
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
    # Helper: Build Single Table without Links (Plain Text)
    #----------------------------------------------------------------
    function Build-SingleTableNoLinks {
        param(
            [PSCustomObject[]]$Data,
            [string]$Heading
        )

        if (-not $Data -or $Data.Count -eq 0) {
            return ""
        }

        # Gather all property names
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
    # Start Building the HTML Report
    #----------------------------------------------------------------
    # Determine Day, Date, and AM/PM based on current time
    $currentDay = (Get-Date).ToString("dddd")             # Full day name, e.g., "Monday"
    $currentDate = (Get-Date).ToString("dd MMMM yyyy")   # Date in "DD Month YYYY", e.g., "27 October 2025"
    $timePart = (Get-Date).ToString("tt")                 # Returns "AM" or "PM"

    # Build the Email Header
    $emailHeader = "ZL $timePart Shift Turnover - $currentDay $currentDate"

    # Start the HTML content
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

    /* Color classes */
    .none   { background-color: transparent; }
    .red    { background-color: #ffcccc; }
    .yellow { background-color: #ffffcc; }
    .green  { background-color: #ccffcc; }
</style>
</head>
<body>
<h1>$emailHeader</h1>

<pre>$Description</pre>
<div class="container">
"@

    #----------------------------------------------------------------
    # Build Tables for AllObjects
    # One table per PSCustomObject (2 columns: PropertyName | Value)
    #----------------------------------------------------------------
    foreach ($obj in $AllObjects) {
        $envName = $obj.Environment
        $tableHeading = $obj.Name

        $html += @"
<div class="table-container">
  <h2>$tableHeading</h2>
  <table>
"@

        foreach ($prop in $obj.PSObject.Properties) {
            # Include 'Environment' row
            if ($prop.Name -eq 'Name') { continue }

            $propName  = $prop.Name
            $propValue = $prop.Value

            # Determine cell class based on thresholds
            $cellClass = "none"
            if ($prop.Name -ne 'Environment') {
                $cellClass = Get-CellClass -Environment $envName -PropName $propName -PropValue $propValue
            }

            $html += "<tr><td>$propName</td><td class='$cellClass'>$propValue</td></tr>"
        }

        $html += "</table></div>"
    }

    #----------------------------------------------------------------
    # Build Tables for ServiceNow, Tasks, Patching
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

    #----------------------------------------------------------------
    # Output the HTML to a file
    #----------------------------------------------------------------
    $OutputPath = "D:\PowerShell\Test\CustomReport.html"
    $html | Out-File -FilePath $OutputPath -Encoding utf8
    Write-Host "HTML report generated at $OutputPath"

    #----------------------------------------------------------------
    # Send the Email
    #----------------------------------------------------------------
    # Construct the email subject dynamically
    $emailSubject = "ZL $timePart Shift Turnover - $currentDay $currentDate"

    try {
        Send-MailMessage `
            -SmtpServer $Script:Config.email.smtp `
            -To $Script:Config.email.to `
            -From $Script:Config.email.from `
            -Subject $emailSubject `
            -Body $html `
            -BodyAsHtml `
            -ErrorAction Stop

        Write-Host "Email sent successfully to $($Script:Config.email.to)"
    }
    catch {
        Write-Host "ERROR: Failed to send email. $_"
    }

    return $html
}
