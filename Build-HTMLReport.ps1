function Build-HTMLReport {
    param (
        [Parameter(Mandatory = $true)]
        [array]$CustomObjects,  # Array of PowerShell custom objects to create tables from

        [Parameter(Mandatory = $true)]
        [string]$Description,   # Text from literal array (before the tables)

        [Parameter(Mandatory = $true)]
        [string]$FooterText,    # Text from second literal array (below the tables)

        [Parameter(Mandatory = $true)]
        [int]$ThresholdRed = 30,   # Set default threshold for Red
        [Parameter(Mandatory = $true)]
        [int]$ThresholdYellow = 60 # Set default threshold for Yellow
    )

    # Start HTML document structure
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
        .red { background-color: #ffcccc; }     /* Light Red */
        .yellow { background-color: #ffffcc; }  /* Light Yellow */
        .green { background-color: #ccffcc; }   /* Light Green */
    </style>
</head>
<body>
<h1>Custom HTML Report</h1>

<pre>
$Description
</pre>

<div class="container">
"@

    # Create HTML tables from custom objects and add to the HTML structure
    foreach ($object in $CustomObjects) {
        # Convert the custom object to an HTML table
        $html += @"
    <div class="table-container">
        <table>
            <tr>
"@

        # Add table headers based on properties of the first object
        foreach ($property in $object.PSObject.Properties.Name) {
            $html += "<th>$property</th>"
        }
        $html += "</tr>"

        # Add rows of data for each custom object
        foreach ($row in $object) {
            $html += "<tr>"
            foreach ($property in $row.PSObject.Properties.Name) {
                $value = $row.$property

                # Apply conditional formatting based on thresholds
                $cellClass = ""
                if ($property -eq "PercentFree") {
                    $percentValue = [int]($value -replace '[^0-9]', '')  # Extract numeric part

                    if ($percentValue -le $ThresholdRed) {
                        $cellClass = "red"
                    } elseif ($percentValue -le $ThresholdYellow) {
                        $cellClass = "yellow"
                    } else {
                        $cellClass = "green"
                    }
                }

                # Add table cell with conditional formatting
                if ($cellClass) {
                    $html += "<td class='$cellClass'>$value</td>"
                } else {
                    $html += "<td>$value</td>"
                }
            }
            $html += "</tr>"
        }
        $html += "</table></div>"
    }

    # Add Footer Text below tables
    $html += @"
</div>

<pre>
$FooterText
</pre>

</body>
</html>
"@

    # Output the HTML to a file
    $OutputPath = "G:\Users\Shawn\Desktop\CustomReport.html"
    $html | Out-File -FilePath $OutputPath -Encoding utf8

    Write-Host "HTML report generated at $OutputPath"
}

# Example usage
$Description = @"
This report contains tables generated from custom objects.
Each table represents different data points.
"@

$FooterText = @"
This is additional information that appears after the tables.
It could be notes or conclusions about the data.
"@

# Create example custom objects 
$object1 = [PSCustomObject]@{ Name = "Server1"; "TotalSize" = "576 Gb"; "UsedSpace" = "255.62 Gb"; "FreeSpace" = "321.16 Gb"; "PercentFree" = "79 %" }
$object2 = [PSCustomObject]@{ Name = "Server2"; "TotalSize" = "580 Gb"; "UsedSpace" = "224.38 Gb"; "FreeSpace" = "321.16 Gb"; "PercentFree" = "69 %" }
$object3 = [PSCustomObject]@{ Name = "Server3"; "TotalSize" = "580 Gb"; "UsedSpace" = "124.38 Gb"; "FreeSpace" = "321.16 Gb"; "PercentFree" = "38 %" }

# Build report with custom objects and thresholds
$T = Build-HTMLReport -CustomObjects @($object1, $object2, $object3) -Description $Description -FooterText $FooterText -ThresholdRed 30 -ThresholdYellow 60
