function Build-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject[]]
        $AllObjects,   # multiple objects => multiple separate tables (no hyperlink logic)

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]
        $ServiceNow,   # single table, embed link into Name

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]
        $Tasks,        # single table, embed link into Name

        [Parameter(Mandatory=$false)]
        [PSCustomObject[]]
        $Patching,     # single table, plain text

        [string]$Description,
        [string]$FooterText
    )

    #-----------------------------------------------------------------
    # 1) Build-MultiObjectTables (for $AllObjects)
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
        $envName = $obj.Environment
        if (-not $envName) { continue }

        # Grab the rules for this environment
        $rules = $Script:Config.Thresholds.$envName

        $tableHeading = $obj.Name
        $htmlBlock += @"
    <div class="table-container">
        <h2>$tableHeading</h2>
        <table>
"@

        foreach ($prop in $obj.PSObject.Properties) {
            if ($prop.Name -eq 'Name') { continue }

            $propName = $prop.Name
            $propValue = $prop.Value
            $cellClass = "none"  # default

            # Find any matching rule object, e.g. { PropertyName='TIPSyncDate', Red='olderThan7Days' }
            $rule = $rules | Where-Object { $_.PropertyName -eq $propName }
            if ($rule) {
                # For each color condition in [Green,Yellow,Red], check if it matches the "date" rule
                # We'll do an example for the Red case:
                if ($rule.Red -eq "olderThan7Days") {
                    # parse $propValue to a [datetime] if possible
                    try {
                        $dateVal = [datetime] $propValue
                        # If it's older than 7 days => set Red
                        if ($dateVal -lt (Get-Date).AddDays(-7)) {
                            $cellClass = "red"
                        }
                    }
                    catch {
                        # If parse fails, you might keep it 'none'
                        # or default to red if you want
                    }
                }
                # Also handle other numeric logic for $rule.Red / .Green / .Yellow if needed...
            }

            # Finally add the row with the assigned color class
            $htmlBlock += "<tr><td>$propName</td><td class='$cellClass'>$propValue</td></tr>"
        }

        $htmlBlock += "</table></div>"
    }

    return $htmlBlock
}


    #-----------------------------------------------------------------
    # 2) Build-SingleTableNoLinks (for $Patching)
    #    => one table, no hyperlink logic, shows all columns
    #-----------------------------------------------------------------
    function Build-SingleTableNoLinks {
        param(
            [PSCustomObject[]]$Data,
            [string]$Heading
        )

        if (!$Data -or $Data.Count -eq 0) {
            return ""
        }

        # Gather all property names
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
        foreach ($p in $allProps) {
            $block += "<th>$p</th>"
        }
        $block += "</tr>"

        foreach ($obj in $Data) {
            $block += "<tr>"
            foreach ($p in $allProps) {
                $value = $obj.$p
                # No hyperlink logic; show raw text
                $block += "<td>$value</td>"
            }
            $block += "</tr>"
        }
        $block += "</table></div>"
        return $block
    }

    #-----------------------------------------------------------------
    # 3) Build-SingleTableEmbedLink (for $ServiceNow, $Tasks)
    #    => one table, hides TableName & Link columns
    #    => if there's a Link property, embed it in the Name column
    #-----------------------------------------------------------------
    function Build-SingleTableEmbedLink {
        param(
            [PSCustomObject[]]$Data,
            [string]$Heading
        )

        if (!$Data -or $Data.Count -eq 0) {
            return ""
        }

        # Gather all property names
        $allProps = $Data | ForEach-Object {
            $_.PSObject.Properties.Name
        } | Select-Object -Unique

        # We remove 'TableName' and 'Link' from visible columns
        $allProps = $allProps | Where-Object { $_ -notin @('TableName','Link') }

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

            # If there's a Link property, we store it so we can embed it in Name
            $linkValue = $obj.Link
            $hasLink   = $false
            if ($linkValue -is [string] -and $linkValue -match '^https?://') {
                $hasLink = $true
            }

            foreach ($p in $allProps) {
                $value = $obj.$p

                if ($p -eq 'Name' -and $hasLink) {
                    # Optionally remove "Service Now " prefix from the name
                    $displayName = $value -replace '^Service Now\s*', ''

                    # Embed the link in the Name cell
                    $value = "<a href='$linkValue' target='_blank'>$displayName</a>"
                }

                $block += "<td>$value</td>"
            }
            $block += "</tr>"
        }

        $block += "</table></div>"
        return $block
    }

    foreach ($obj in $AllObjects) {
    Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Dead Servers' -Value ''
    Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Search Status' -Value ''
    Add-Member -InputObject $obj -MemberType NoteProperty -Name 'UI Check' -Value ''
}

    #-----------------------------------------------------------------
    # Build the final HTML
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

    # 1) Multiple small tables for $AllObjects (no link logic)
    $html += Build-MultiObjectTables -ObjArray $AllObjects -SectionHeading "All Objects"

    # 2) Single table for $ServiceNow, embed link in Name, hide TableName/Link columns
    if ($ServiceNow) {
        $html += Build-SingleTableEmbedLink -Data $ServiceNow -Heading "ServiceNow"
    }

    # 3) Single table for $Tasks, same embed logic, hide TableName/Link columns
    if ($Tasks) {
        $html += Build-SingleTableEmbedLink -Data $Tasks -Heading "Tasks"
    }

    # 4) Single table for $Patching, plain text
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

<#
JSON

{
  "Thresholds": {
    "CA2015": [
      {
        "PropertyName": "Live Servers",
        "Red": "!=189"
      },
      {
        "PropertyName": "Queueing",
        "Red": ">0"
      },
      {
        "PropertyName": "Active_Search_Partition_Count",
        "Green": "<300000000",
        "Yellow": ">=3000000001 && <379999999",
        "Red": "<380000000"
      },
      {
        "PropertyName": "DBQueue Status 0",
        "Red": ">1000000"
      },
      {
        "PropertyName": "DBQueue Status 284"
      },
      {
        "PropertyName": "DBQueue Status 286"
      },
      {
        "PropertyName": "DBQueue Status 5000"
      }
    ],
    "PA": [
      {
        "PropertyName": "Live Servers",
        "Red": "!=39"
      },
      {
        "PropertyName": "Queueing",
        "Red": ">0"
      },
      {
        "PropertyName": "MTA_Monitor",
        "Red": "=0"
      },
      {
        "PropertyName": "Active_Search_Partition_Count",
        "Green": "<300000000",
        "Yellow": ">=3000000001 && <379999999",
        "Red": "<380000000"
      }
    ],
    "MARRS": [
      {
        "PropertyName": "Live Servers",
        "Red": "!=105"
      },
      {
        "PropertyName": "Queueing",
        "Red": ">0"
      }
    ],
    "PZL": [
      {
        "PropertyName": "Queueing",
        "Red": ">0"
      },
      {
        "PropertyName": "Live Servers",
        "Red": "!=13"
      }
    ],
    "1ZLA": [
      {
        "PropertyName": "Queueing",
        "Red": ">0"
      },
      {
        "PropertyName": "Live Servers",
        "Red": "!=80"
      }
    ],
    "AGE": [
      {
        "PropertyName": "Queueing",
        "Red": ">0"
      },
      {
        "PropertyName": "Live Servers",
        "Red": "!=8"
      }
    ],
    "DI": [
      {
        "PropertyName": "Live Servers",
        "Red": "!=38"
      },
      {
        "PropertyName": "Queueing",
        "Red": ">0"
      }
    ],
    "HI": [
      {
        "PropertyName": "Live Servers",
        "Red": "!=49"
      },
      {
        "PropertyName": "Queueing",
        "Red": ">0"
      }
    ],
    "CB": [
      {
        "PropertyName": "Live Servers",
        "Red": "!=24"
      },
      {
        "PropertyName": "MTA_Monitor",
        "Red": "=0"
      },
      {
        "PropertyName": "Active_Search_Partition_Count",
        "Green": "<300000000",
        "Yellow": ">=3000000001 && <379999999",
        "Red": "<380000000"
      },
      {
        "PropertyName": "Queueing",
        "Red": ">0"
      }
    ],
    "CA2022": [
      {
        "PropertyName": "Live Servers",
        "Red": "!=230"
      },
      {
        "PropertyName": "TIPSyncDate",
        "Red": "olderThan7Days"
      },
      {
        "PropertyName": "MTA_Monitor",
        "Yellow": ">=1 && <=10",
        "Red": ">10"
      },
      {
        "PropertyName": "Active_Search_Partition_Count",
        "Green": "<300000000",
        "Yellow": ">=3000000001 && <379999999",
        "Red": "<380000000"
      },
      {
        "PropertyName": "Queueing",
        "Red": ">0"
      }
    ]
  }
}


#>