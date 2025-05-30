<#
.SYNOPSIS
    Tests Microsoft Store (msstore) availability of the "Company Portal" app across all supported geographical regions using their GeoID values.

.DESCRIPTION
    This script automates region-based testing by:
    - Importing a list of GeoIDs and region names from a CSV file.
    - Temporarily setting the system's home location (GeoID) to each region.
    - Using `winget` to search for the "Company Portal" app in the msstore source.
    - Recording whether the app is available in that region.
    - Restoring the original system GeoID after testing.

    The final results are exported to a CSV file for analysis.

.PARAMETER geoListPath
    The path to the CSV file containing GeoID and region name pairs (semicolon-delimited).

.OUTPUTS
    A CSV file named 'CompanyPortalAvailability_Result.csv' listing each GeoID, region, and whether the app was found.

.NOTES
    Author: Kevin Schouten
    Version: v1.0
    Requires: PowerShell 5.1+, Winget with msstore source enabled
    
    !!Use at your own risk!!

.LINK
    https://learn.microsoft.com/en-us/windows/win32/intl/table-of-geographical-locations
#>

# created a CSV from: https://learn.microsoft.com/en-us/windows/win32/intl/table-of-geographical-locations
$geoListPath = "GeoID_List.csv"

# Import the CSV file
$geoList = Import-Csv -Path $geoListPath -Delimiter ";"

# Array to store results
$results = @()

#store original GeoID
$OriginalGeo = Get-WinHomeLocation

Write-Host -ForegroundColor Yellow "Storing Original GeoID: $($OriginalGeo.GeoId)"

# process each region
foreach ($geo in $geoList) {
    $geoID = [int]$geo.GEOID
    $region = $geo.Location

    $currentcount ++
    Write-Host "Testing Region: $region (GeoID: $geoID) - Progress: $currentcount / $($geolist.count)"

    try {
        # Set region
        Set-WinHomeLocation -GeoId $geoID
        Start-Sleep -Seconds 1

        # Run winget search
        $output = $null
        $output = winget search 'Company Portal' --source msstore

        #Validate the output
        if ($output -match "9WZDNCRFJ3PZ") {
            $found = "Yes"
        } else {
            $found = "No"
        }
    } catch {
        $found = "Error"
        Write-Warning "Error during processing GeoID $geoID - $region : $_"
    }

    $results += [PSCustomObject]@{
        GEOID   = $geoID
        Region  = $region
        Found   = $found
    }
}

#Restore Orginial GEOID
Write-Host -ForegroundColor Yellow "Restoring Original GeoID: $($OriginalGeo.GeoId)"
Set-WinHomeLocation -GeoID $OriginalGeo.GeoID

# Export the results
$outputPath = "CompanyPortalAvailability_Result.csv"
$results | Export-Csv -Path $outputPath -NoTypeInformation -Delimiter ";"
Write-Host "Results exported to $outputPath"
