
$DomainName = $env:USERDNSDOMAIN  
$GPOsInDomain = Get-GPO -All -Domain $DomainName 
# $GPOSpecific = Get-GPO -Name  - Specific GPO to target
$ScriptPath = "C:\Users\ABaquilo\OneDrive - City National Bank\Documents\GPO Scripts\GPO Output"


if (!(get-module -name GroupPolicy)) {Import-Module -Name GroupPolicy -Force}


$GPODrivePaths = foreach ($GPOName in $GPOsInDomain) {
            
        [xml]$GPOreport = Get-GPOReport -Name $GPOName.DisplayName -ReportType Xml      
        $driveMapSettings = (($GPOreport.GPO.User.ExtensionData | Where-Object Name -eq "Drive Maps").Extension.DriveMapSettings).drive
        
        foreach ($i in $GPOName ) {
                [PSCustomObject] @{
                    "GPO Name" = $i.DisplayName  -join ';' # GPO name
                    "Is Enabled" = $i.GpoStatus  -join ';' # Verify GPO enabled
                    "Path" = $i.SOMPath  -join ';' # GPO Link path
                    "Drive Letter" = $driveMapSettings.Properties.Letter  -join ';' # Mapped Drive letter
                    "Drive Path" = $driveMapSettings.Properties.Path  -join ';' # Mapped Drive Path
                    "Order" = $driveMapSettings.GPOSettingOrder  -join ';' # Order number
                    "Creation Time" = $i.CreationTime   -join ';' # Time GPO was created
                    "Modification Time" = $i.ModificationTime  -join ';' # Time GPO was modified
                    "Owner" = $i.Owner   -join ';' # GPO Owner
                    "Domain Name" = $i.DomainName -join ';' # Domain Name
                }
        }

        $driveMapSettings = $null
}


Write-host "Completed running the GPO drive map script at: $(get-date -Format G)."
$GPODrivePaths | Sort-Object Name
#$GPODrivePaths | Sort-Object Name | Export-Excel -KillExcel -Path $ScriptPath\GPODriveMaps.xlsx -Append -AutoSize
$GPODrivePaths | Export-Csv -path "$ScriptPath\GPODriveMaps.csv" -NoTypeInformation
