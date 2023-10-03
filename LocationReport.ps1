# This is used to generate a .xlsx file containing location node information
# Location: /sitecore/system/Modules/PowerShell/Script Library/SPE/Tools/Data Management/Toolbox/Data Management/Location Report
# Last updated: 27 July 2022

Import-Function -Name ConvertTo-Xlsx
f
$reportDate = Get-Date -Format "MM-dd-yyyy_HH-mm-ss"

$path = "master:\content\KelseySeybold\KelseySeyboldGlobal\Data\Locations"

[byte[]] $data = Get-ChildItem -Path $path -Recurse | 
    Select-Object -Property "External ID", LocationName, PhoneNumber, AddressLine1, City, ZipCode, Latitude, Longitude, MetaDescription, DepartmentID, 
        @{Name='Date Created'; Expression={
            $date = $_["__Created"]
            $formatDate = [System.Text.StringBuilder]::new()
            [void]$formatDate.Append($date.substring(0,4)).Append("-").Append($date.SubString(4,2)).Append("-").Append($date.SubString(6,2))
            $formatDate.ToString()
        }} | 
    ConvertTo-Xlsx

Out-Download -Name $reportDate" LocationReport".xlsx -InputObject $data
