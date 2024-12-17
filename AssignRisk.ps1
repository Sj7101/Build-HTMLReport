function Set-RiskLevel {
    param(
        [Parameter(Mandatory=$true)]
        [array]$CustomObjects,

        [Parameter(Mandatory=$true)]
        [string]$ConfigFile
    )

    # Load the config file (it contains the thresholds)
    $config = Get-Content -Path $ConfigFile | ConvertFrom-Json

    # Iterate through each PSCustomObject in the array
    foreach ($obj in $CustomObjects) {
        # Loop through each key (field) in the PSCustomObject
        foreach ($key in $obj.PSObject.Properties.Name) {
            # Check if the key exists in the thresholds section of the config
            if ($config.thresholds.ContainsKey($key)) {
                $threshold = $config.thresholds[$key]

                # Get the value from the PSCustomObject field
                $value = $obj.$key

                # Default to "None" if the value is below the threshold for any risk level
                $riskLevel = "None"

                # Apply the risk level based on the range thresholds
                if ($value -ge $threshold.Low -and $value -lt $threshold.Medium) {
                    $riskLevel = "Low"
                }
                elseif ($value -ge $threshold.Medium -and $value -lt $threshold.High) {
                    $riskLevel = "Medium"
                }
                elseif ($value -ge $threshold.High) {
                    $riskLevel = "High"
                }

                # Add a new property 'Risk Level' with the calculated risk
                $obj | Add-Member -MemberType NoteProperty -Name "Risk Level" -Value $riskLevel -Force
            }
        }
    }

    # Return the updated objects
    return $CustomObjects
}

# Example usage:
$configFilePath = "C:\path\to\config.json"
$customObjects = @(
    [PSCustomObject]@{ TableName = "CA2022"; "Live Servers" = 230; "TIPSyncDate" = "10/14/2024 6:51"; "MTAMonitor" = 12245; "Database Free Space" = 23.45; "DBQueue Status 0" = 24; "DBQueue Status 5000" = 158 },
    [PSCustomObject]@{ TableName = "CA2023"; "Live Servers" = 150; "TIPSyncDate" = "10/10/2024 9:10"; "MTAMonitor" = 9000; "Database Free Space" = 19.5; "DBQueue Status 0" = 110; "DBQueue Status 5000" = 10 }
)

$updatedObjects = Set-RiskLevel -CustomObjects $customObjects -ConfigFile $configFilePath

# Display the updated objects with Risk Level
$updatedObjects
