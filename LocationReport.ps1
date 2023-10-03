<#
Script Name:    LocationReport
Version:        1.1.0
Author:         Jared Caraway (jared.caraway@kelsey-seybold.com)
Date Updated:   10/03/2023
Change Log:     10/03/23:   Initial version. Added date of creation to output; added timestamp to filename.

Description:    This script iterates through all location pages and exports a spreadsheet in
                XLSX format.
#>

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
