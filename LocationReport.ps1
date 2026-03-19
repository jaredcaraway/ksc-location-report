<#
Script Name:    LocationReport
Version:        1.2.0
Author:         Jared Caraway (jared.caraway@kelsey-seybold.com)
Date Updated:   03/19/2026
Change Log:     10/03/23:   Initial version. Added date of creation to output; added timestamp to filename.
                03/19/26:   Added campus building attributes; filtered to published locations only.

Description:    This script iterates through all location pages and exports a spreadsheet in
                XLSX format. Only published locations are included.
#>

Import-Function -Name ConvertTo-Xlsx

$reportDate = Get-Date -Format "MM-dd-yyyy_HH-mm-ss"

$path = "master:\content\KelseySeybold\KelseySeyboldGlobal\Data\Locations"

[byte[]] $data = Get-ChildItem -Path $path -Recurse |
    Where-Object { $_.Publishing.IsPublishable((Get-Date), $true) } |
    Select-Object -Property "External ID", LocationName, PhoneNumber, AddressLine1, City, ZipCode, Latitude, Longitude, MetaDescription, DepartmentID,
        CampusFacilityDirectionAndParking, CampusFacilityPhoneNumber, CampusFacilityExternalId,
        @{Name='Date Created'; Expression={
            $date = $_["__Created"]
            $formatDate = [System.Text.StringBuilder]::new()
            [void]$formatDate.Append($date.substring(0,4)).Append("-").Append($date.SubString(4,2)).Append("-").Append($date.SubString(6,2))
            $formatDate.ToString()
        }} |
    ConvertTo-Xlsx

Out-Download -Name $reportDate" LocationReport".xlsx -InputObject $data
