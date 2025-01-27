function Build-HTMLReport {
    param(
        # This array is for your “Main Items,” each PSCustomObject gets its own table (2-wide).
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]$CustomObjects,

        # Each of these are arrays of PSCustomObjects.
        # We will build exactly ONE table per array.
        [Parameter(Mandatory = $false)]
        [PSCustomObject[]]$ServiceNow,
        
        [Parameter(Mandatory = $false)]
        [PSCustomObject[]]$TASKS,
        
        [Parameter(Mandatory = $false)]
        [PSCustomObject[]]$PATCHING,

        [string]$Description,
        [string]$FooterText
    )

    #------------------------------------------------------------------
    # Helper 1: Build multiple tables (one per object) for the main array
    #           => "2-wide" layout using flex.
    #------------------------------------------------------------------
    function Build-MultiObjectTables {
        param(
            [PSCustomObject[]]$ObjArray,
            [string]$SectionHeading
        )

        if (!$ObjArray -or $ObjArray.Count -eq 0) { return "" }

        $htmlBlock = @"
    <div style="width:100%">
        <h1>$SectionHeading</h1>
    </div>
"@

        foreach ($obj in $ObjArray) {
            # If each object has TableName or Name, you can decide which to show as <h2>
            $tableHeading = $obj.Name

            $htmlBlock += @"
    <div class="table-container">
        <h2>$tableHeading</h2>
        <table>
"@

            # Build 2 columns: PropertyName | Value
            foreach ($prop in $obj.PSObject.Properties) {
                if ($prop.Name -eq 'Name' -or $prop.Name -eq 'TableName') {
                    # We skip the 'Name' property since we used it as heading.
                    # We skip 'TableName' so it doesn't appear as a column.
                    continue
                }

                $propName  = $prop.Name
                $propValue = $prop.Value

                # Convert link fields named Link or URL into <a> with the object’s Name
                if ($propName -in @('Link','URL')) {
                    $propValue = "<a href='$propValue' target='_blank'>$($obj.Name)</a>"
                }

                $htmlBlock += "<tr><td>$propName</td><td>$propValue</td></tr>"
            }
            $htmlBlock += "</table></div>"
        }

        return $htmlBlock
    }

    #------------------------------------------------------------------
    # Helper 2: Build ONE table from an array of PSCustomObjects
    #           => All objects in a single table with multiple rows.
    #           => The table heading is from the first object's TableName (if present).
    #------------------------------------------------------------------
    function Build-SingleTable {
        param(
            [PSCustomObject[]]$ObjArray
        )

        if (!$ObjArray -or $ObjArray.Count -eq 0) {
            return ""
        }

        # Use the first object's TableName (if present) as the table's heading
        $firstObj = $ObjArray[0]
        $heading = if ($firstObj.PSObject.Properties.Name -contains 'TableName') {
            $firstObj.TableName
        } else {
            # fallback
            "Untitled Table"
        }

        # Gather all properties from all objects to build the columns
        $allProps = $ObjArray | ForEach-Object {
            $_.PSObject.Properties.Name
        } | Select-Object -Unique

        # Exclude TableName from the columns themselves (since we used it as heading)
        $allProps = $allProps | Where-Object { $_ -ne 'TableName' }

        $htmlBlock = @"
    <div style="width:100%">
        <h1>$heading</h1>
    </div>
    <div class="table-container" style="width:100%"> <!-- single wide table -->
        <table>
            <tr>
"@

        # Build table headers
        foreach ($p in $allProps) {
            $htmlBlock += "<th>$p</th>"
        }
        $htmlBlock += "</tr>"

        # One row per PSCustomObject
        foreach ($obj in $ObjArray) {
            $htmlBlock += "<tr>"
            foreach ($p in $allProps) {
                $propValue = $obj.$p

                # If property is Link or URL, turn it into a clickable text using the object's Name
                if ($p -in @('Link','URL') -and $propValue) {
                    # If there's no Name property, fallback to the raw URL or something else
                    $linkText = if ($obj.PSObject.Properties.Name -contains 'Name') {
                        $obj.Name
                    } else {
                        $propValue
                    }
                    $propValue = "<a href='$propValue' target='_blank'>$linkText</a>"
                }

                $htmlBlock += "<td>$propValue</td>"
            }
            $htmlBlock += "</tr>"
        }
        
        $htmlBlock += "</table></div>"
        return $htmlBlock
    }

    #-------------------------------
    # Start building the HTML
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

    # 1) Multiple tables for the $CustomObjects array
    $html += Build-MultiObjectTables -ObjArray $CustomObjects -SectionHeading "Main Items"

    # 2) ONE table for all $ServiceNow objects
    if ($ServiceNow)
