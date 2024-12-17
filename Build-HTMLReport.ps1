function Build-HTMLReport {
    param (
        [Parameter(Mandatory = $true)]
        [array]$CustomObjects,  # Array of PowerShell custom objects to create tables from

        [Parameter(Mandatory = $true)]
        [string]$Description,   # Text from literal array (before the tables)

        [Parameter(Mandatory = $true)]
        [string]$FooterText     # Text from second literal array (below the tables)
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
    foreach ($objectGroup in $CustomObjects) {
        # Extract table name from custom object (assuming TableName is the property holding the name)
        $tableName = $objectGroup[0].TableName

        $html += @"
    <div class="table-container">
        <h2>$tableName</h2>
        <table>
"@

        # If it's a ServiceNow table, display the link in the left column and leave the right column empty
        if ($tableName -eq "ServiceNow") {
            foreach ($row in $objectGroup) {
                if ($row.PSObject.Properties.Match('Link')) {
                    $link = $row.Link
                    $html += "<tr><td><a href='$link' target='_blank'>$link</a></td><td></td></tr>"
                }
            }
        } else {
            # Add rows of data for each custom object (formatted vertically with property name on the left, value on the right)
            foreach ($row in $objectGroup) {
                $properties = $row.PSObject.Properties | Where-Object { $_.Name -notin @('TableName', 'Link', 'Threshold') }
                
                foreach ($property in $properties) {
                    $propertyName = $property.Name  # Correctly capture the property name
                    $value = $row.$($property.Name)  # Capture the property value

                    # Apply color coding based on 'Threshold'
                    $cellClass = ""
                    if ($row.PSObject.Properties.Match('Threshold')) {
                        $threshold = $row.Threshold
                        switch ($threshold) {
                            "High" { $cellClass = "red" }
                            "Medium" { $cellClass = "yellow" }
                            "Low" { $cellClass = "green" }
                            default { $cellClass = "" }  # No coloring if threshold is blank or not defined
                        }
                    }

                    # Add table row: Property name on the left, value on the right
                    $html += "<tr><td class='$cellClass'>$propertyName</td><td class='$cellClass'>$value</td></tr>"
                }
            }
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


# Example custom objects with dynamic tables and Thresholds for color coding
$object1 = @(
    [PSCustomObject]@{ TableName = "MARRS"; "TotalSize" = "576 Gb"; "UsedSpace" = "255.62 Gb"; "FreeSpace" = "321.16 Gb"; "PercentFree" = "29 %"; Threshold = "Medium" },
    [PSCustomObject]@{ TableName = "MARRS"; "TotalSize" = "576 Gb"; "UsedSpace" = "300.50 Gb"; "FreeSpace" = "275.50 Gb"; "PercentFree" = "48 %"; Threshold = "Low" },
    [PSCustomObject]@{ TableName = "MARRS"; "TotalSize" = "576 Gb"; "UsedSpace" = "450.00 Gb"; "FreeSpace" = "126.00 Gb"; "PercentFree" = "21 %"; Threshold = "Medium" },
    [PSCustomObject]@{ TableName = "MARRS"; "TotalSize" = "576 Gb"; "UsedSpace" = "575.00 Gb"; "FreeSpace" = "1.00 Gb"; "PercentFree" = "0.17 %"; Threshold = "High" }
)

$object2 = @(
    [PSCustomObject]@{ TableName = "PZL"; "Count" = "213"; "Dead Count" = "10"; "TCP Check" = "Good"; "Partition" = "236541"; Threshold = "High" },
    [PSCustomObject]@{ TableName = "PZL"; "Count" = "13"; "Dead Count" = "1"; "TCP Check" = "Good"; "Partition" = "236541"; Threshold = "Low" },
    [PSCustomObject]@{ TableName = "PZL"; "Count" = "213"; "Dead Count" = "10"; "TCP Check" = "Good"; "Partition" = "236541"; Threshold = "Medium" },
    [PSCustomObject]@{ TableName = "PZL"; "Count" = "213"; "Dead Count" = "10"; "TCP Check" = "Good"; "Partition" = "236541"; Threshold = "High" }
)

$object3 = @(
    [PSCustomObject]@{ TableName = "Server 3"; Name = "Server3"; "TotalSize" = "580 Gb"; "UsedSpace" = "124.38 Gb"; "FreeSpace" = "321.16 Gb"; "PercentFree" = "38 %"; Threshold = "Low" }
)

$object4 = @(
    [PSCustomObject]@{ TableName = "Server 4"; Name = "Server4"; "TotalSize" = "580 Gb"; "UsedSpace" = "124.38 Gb"; "FreeSpace" = "321.16 Gb"; "PercentFree" = "78 %"; Threshold = "Low" }
)

$Static1 = @(
    [PSCustomObject]@{ TableName = "ServiceNow"; Name = "ServiceNow Incident Queue"; Link = "http://ServiceNow.com/Inc"; "Verified" = "" },
    [PSCustomObject]@{ TableName = "ServiceNow"; Name = "ServiceNow Change Queue"; Link = "http://ServiceNow.com/CHG" ; "Verified" = "" }
)

# Define description and footer text
$Description = "This report shows the disk usage details for multiple servers with dynamic thresholds for color coding."
$FooterText = "Generated by the PowerShell script."

# Build report with custom objects, including ServiceNow objects
$T = Build-HTMLReport -CustomObjects @($object1, $object2, $object3, $object4, $Static1) -Description $Description -FooterText $FooterText
