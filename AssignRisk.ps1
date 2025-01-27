function Set-RiskLevel {
    <#
    .SYNOPSIS
        Assigns risk levels to custom objects based on defined thresholds.

    .DESCRIPTION
        This function evaluates each property of the provided custom objects against thresholds defined in a JSON configuration file.
        It assigns a risk level ("None", "Low", "Medium", "High") based on where each property's value falls within the thresholds.
        Properties missing from an object are skipped without assigning any risk level.

    .PARAMETER CustomObjects
        An array of PSCustomObject instances to evaluate.

    .PARAMETER ConfigFile
        The file path to the JSON configuration file containing thresholds and risk directions.

    .EXAMPLE
        $updatedObjects = Set-RiskLevel -CustomObjects $customObjects -ConfigFile "C:\Path\To\Config.json"

    .NOTES
        Ensure that the configuration file includes 'RiskDirection' and 'Levels' for each property.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$CustomObjects,

        [Parameter(Mandatory = $true)]
        [string]$ConfigFile
    )

    # Enable Verbose Output
    $VerbosePreference = "Continue"

    # Load the configuration
    try {
        Write-Verbose "Loading configuration from '$ConfigFile'."
        $configContent = Get-Content -Path $ConfigFile -ErrorAction Stop
        $config = $configContent | ConvertFrom-Json
        Write-Verbose "Configuration loaded successfully."
    }
    catch {
        Throw "Failed to load or parse the configuration file at '$ConfigFile'. Error: $_"
    }

    # Validate configuration
    if (-not $config.thresholds) {
        Throw "Configuration file must contain a 'thresholds' section."
    }

    # Helper Function to Determine Risk Level
    function Get-RiskLevel {
        param(
            [decimal]$Value,
            [string]$RiskDirection,
            [hashtable]$Thresholds
        )

        Write-Verbose "Determining risk level for Value: $Value, RiskDirection: $RiskDirection"

        switch ($RiskDirection) {
            "High" {
                Write-Verbose "RiskDirection: High"
                Write-Verbose "Thresholds: High=$($Thresholds.High), Medium=$($Thresholds.Medium), Low=$($Thresholds.Low), None=$($Thresholds.None)"
                if ($Value -le $Thresholds.High) {
                    Write-Verbose "Assigned Risk Level: High"
                    return "High"
                }
                elseif ($Value -le $Thresholds.Medium) {
                    Write-Verbose "Assigned Risk Level: Medium"
                    return "Medium"
                }
                elseif ($Value -le $Thresholds.Low) {
                    Write-Verbose "Assigned Risk Level: Low"
                    return "Low"
                }
                else {
                    Write-Verbose "Assigned Risk Level: None"
                    return "None"
                }
            }
            "Low" {
                Write-Verbose "RiskDirection: Low"
                Write-Verbose "Thresholds: High=$($Thresholds.High), Medium=$($Thresholds.Medium), Low=$($Thresholds.Low), None=$($Thresholds.None)"
                if ($Value -ge $Thresholds.None) {
                    Write-Verbose "Assigned Risk Level: None"
                    return "None"
                }
                elseif ($Value -ge $Thresholds.Low) {
                    Write-Verbose "Assigned Risk Level: Low"
                    return "Low"
                }
                elseif ($Value -ge $Thresholds.Medium) {
                    Write-Verbose "Assigned Risk Level: Medium"
                    return "Medium"
                }
                elseif ($Value -le $Thresholds.High) {
                    Write-Verbose "Assigned Risk Level: High"
                    return "High"
                }
                else {
                    Write-Verbose "Assigned Risk Level: Unknown"
                    return "Unknown"
                }
            }
            default {
                Write-Warning "Unknown RiskDirection '$RiskDirection'. Assigning 'None'."
                return "None"
            }
        }
    }

    # Iterate through each custom object
    foreach ($obj in $CustomObjects) {
        Write-Verbose "Processing object: $($obj.TableName)"

        # Iterate through each property defined in the config
        foreach ($key in $config.thresholds.PSObject.Properties.Name) {
            # Check if the object has this property
            if ($obj.PSObject.Properties.Name -contains $key) {
                Write-Verbose "Evaluating property: $key"

                $thresholdConfig = $config.thresholds[$key]

                # Output the entire thresholdConfig for debugging
                Write-Verbose "Threshold Config for '$key': $(ConvertTo-Json $thresholdConfig -Depth 5)"

                # Retrieve RiskDirection and Levels
                $riskDirection = $thresholdConfig.RiskDirection
                $thresholds = $thresholdConfig.Levels
                $value = $obj.$key

                # Initialize risk level
                $riskLevel = "Unknown"

                # Handle TIPSyncDate separately if needed
                if ($key -eq "TIPSyncDate" -and $value -match '^\d{1,2}/\d{1,2}/\d{4}\s\d{1,2}:\d{2}') {
                    try {
                        # Convert the string to DateTime
                        $dateValue = [datetime]$value

                        # Calculate the difference in days from current date
                        $daysAgo = (New-TimeSpan -Start $dateValue -End (Get-Date)).Days

                        Write-Verbose "TIPSyncDate ($value) is $daysAgo days ago."

                        # Assign risk levels based on days since the date
                        if ($daysAgo -le $thresholds.None) {
                            $riskLevel = "None"
                        }
                        elseif ($daysAgo -le $thresholds.Low) {
                            $riskLevel = "Low"
                        }
                        elseif ($daysAgo -le $thresholds.Medium) {
                            $riskLevel = "Medium"
                        }
                        else {
                            $riskLevel = "High"
                        }

                        # Format the DateTime object back to a string in the desired format
                        $obj.TIPSyncDate = $dateValue.ToString("MM/dd/yyyy HH:mm")
                    }
                    catch {
                        Write-Warning "Failed to parse date for '$key' with value '$value'. Error: $_"
                        $riskLevel = "Invalid Date"
                    }
                }
                else {
                    # For fields that might be percentages like "Database Free Space", remove "%" and convert to decimal
                    if ($value -match '%') {
                        $value = $value -replace '%', ''  # Remove percent sign if present
                        Write-Verbose "Converted '$key' value to decimal by removing '%': $value"
                    }

                    # Attempt to convert to decimal
                    try {
                        $numericValue = [decimal]$value
                        Write-Verbose "Converted '$key' value to decimal: $numericValue"
                    }
                    catch {
                        Write-Warning "Failed to convert value for '$key' with value '$value' to decimal. Error: $_"
                        $riskLevel = "Invalid Value"
                        # Assign and skip further processing
                        $statusPropertyName = "Risk Level for " + $key
                        $obj | Add-Member -MemberType NoteProperty -Name $statusPropertyName -Value $riskLevel -Force
                        continue
                    }

                    # Determine risk level based on risk direction using the helper function
                    $riskLevel = Get-RiskLevel -Value $numericValue -RiskDirection $riskDirection -Thresholds $thresholds
                }

                # Add a new property 'Risk Level for [PropertyName]' with the calculated risk
                $statusPropertyName = "Risk Level for " + $key
                $obj | Add-Member -MemberType NoteProperty -Name $statusPropertyName -Value $riskLevel -Force

                Write-Verbose "'$key' Risk Level: $riskLevel"
            }
            else {
                Write-Verbose "Property '$key' not found in object '$($obj.TableName)'. Skipping."
                # Optionally, you can add a note or do nothing
            }
        }
    }

    # Return the updated objects
    return $CustomObjects
}




# Example usage:
$configFilePath = "G:\Code\Build-HTMLReport\config.json"
# Define the custom objects with all properties
$customObjects = @(
    [PSCustomObject]@{ 
        TableName = "CA2022"; 
        "Live Servers" = 230; 
        "TIPSyncDate" = "10/1/2024 16:51"; 
        "MTAMonitor" = 12245; 
        "Database Free Space" = "23.45%"; 
        "DBQueue Status 0" = 24; 
        "DBQueue Status 5000" = 158;
        "DBQueue Status 284" = 10;
        "DBQueue Status 286" = 220 
    },
    [PSCustomObject]@{ 
        TableName = "MARRS"; 
        "Live Servers" = 4; 
        "TIPSyncDate" = "10/14/2024 6:51"; 
        "MTAMonitor" = 15245; 
        "Database Free Space" = "13.45%"; 
        "DBQueue Status 0" = 94; 
        "DBQueue Status 5000" = 548;
        "DBQueue Status 284" = 110;
        "DBQueue Status 286" = 60 
    },
    [PSCustomObject]@{ 
        TableName = "PA"; 
        "Live Servers" = 101; 
        "TIPSyncDate" = "12/17/2024 9:10"; 
        "MTAMonitor" = 4250; 
        "Database Free Space" = "70.00%"; 
        "DBQueue Status 0" = 0; 
        "DBQueue Status 5000" = 0;
        "DBQueue Status 284" = 0;
        "DBQueue Status 286" = 0 
    },
    [PSCustomObject]@{ 
        TableName = "CA2018"; 
        "Live Servers" = 400; 
        "Database Free Space" = "13.45%"
    }
)

# Path to the configuration file
$configFilePath = "G:\Code\Build-HTMLReport\config.json"

# Invoke the function
$updatedObjects = Set-RiskLevel -CustomObjects $customObjects -ConfigFile $configFilePath -Verbose

# Display the updated objects with Risk Levels for each property
$updatedObjects | Select-Object `
    TableName, `
    "Live Servers", `
    "Risk Level for Live Servers", `
    "MTAMonitor", `
    "Risk Level for MTAMonitor", `
    "Database Free Space", `
    "Risk Level for Database Free Space", `
    "DBQueue Status 0", `
    "Risk Level for DBQueue Status 0", `
    "DBQueue Status 5000", `
    "Risk Level for DBQueue Status 5000", `
    "DBQueue Status 284", `
    "Risk Level for DBQueue Status 284", `
    "DBQueue Status 286", `
    "Risk Level for DBQueue Status 286" | Format-Table -AutoSize