<#
 Script Name : LocationsReport
 Version     : 2.2.0
 Date Updated: 03/19/2026
 Description : Exports Location items (NO RECURSION)
               - Only includes published locations (__Never publish != "1")
               - Outputs media/link fields as plain public URLs
               - Normalizes /sitecore/shell and relative media URLs to https://www.kelsey-seybold.com
               - Adds **LocationUrl** using canonical pattern:
                 https://www.kelsey-seybold.com/find-a-location/[item-name-slug]
               - Campus locations expand into one row per building in Campus Facilities
                 (e.g. "Summer Creek Campus - Building A")
               - Streamlined column set (removed image, Google tour, directions, and coordinate fields)
               - LocationName and LocationUrl render as clickable hyperlinks via EPPlus post-processing
#>

Import-Function -Name ConvertTo-Xlsx

# ------------------ URL Normalizers -----------------------

# Normalize any Sitecore UI/relative media URL to the public host
function Convert-ToPublicUrl {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) { return "" }

    $u = $Url.Trim()

    # If already absolute, keep as-is
    if ($u -match '^(?i)https?://') { return $u }

    # Collapse duplicate slashes (defensive)
    $u = $u -replace '//+', '/'

    # /sitecore/shell/-/media/... -> https://www.kelsey-seybold.com/-/media/...
    if ($u -like '/sitecore/shell/-/media*') {
        return $u -replace '^/sitecore/shell', 'https://www.kelsey-seybold.com'
    }

    # /-/media/... -> prefix with public host
    if ($u -like '/-/media*') {
        return 'https://www.kelsey-seybold.com' + $u
    }

    # Any other rooted path -> prefix host (safe default)
    if ($u.StartsWith('/')) {
        return 'https://www.kelsey-seybold.com' + $u
    }

    return $u
}

# Build the canonical Location URL from the item name
# Result: https://www.kelsey-seybold.com/find-a-location/[item-name-slug]
function Convert-ItemNameToLocationUrl {
    param([string]$ItemName)

    if ([string]::IsNullOrWhiteSpace($ItemName)) { return "" }

    # Normalize/slugify:
    # - Lowercase
    # - Replace & with 'and'
    # - Drop non letters/digits/spaces/hyphens
    # - Collapse whitespace -> single hyphen
    # - Collapse multiple hyphens
    # - Trim leading/trailing hyphens
    $s = $ItemName.Trim()

    # (Optional) remove diacritics for safety
    try {
        $sNorm = $s.Normalize([Text.NormalizationForm]::FormD)
        $sb = New-Object System.Text.StringBuilder
        foreach ($ch in $sNorm.ToCharArray()) {
            if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
                [void]$sb.Append($ch)
            }
        }
        $s = $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
    } catch { }

    $slug = $s.ToLowerInvariant()
    $slug = $slug -replace '[–—]', '-'         # en/em dashes -> hyphen
    $slug = $slug -replace '&', 'and'          # & -> and
    $slug = $slug -replace '[^a-z0-9\s-]', ''  # remove other punctuation
    $slug = $slug -replace '\s+', '-'          # whitespace -> hyphen
    $slug = $slug -replace '-+', '-'           # collapse hyphens
    $slug = $slug.Trim('-')

    return "https://www.kelsey-seybold.com/find-a-location/$slug"
}

# ------------------ Field Helpers -------------------------

function Get-MediaUrl {
    param($Item, [string]$FieldName)

    $field = $Item.Fields[$FieldName]
    if ($field -eq $null) { return "" }

    # ImageField
    try {
        $img = New-Object Sitecore.Data.Fields.ImageField($field, $Item.Database)
        if ($img -and $img.MediaItem) {
            $url = [Sitecore.Resources.Media.MediaManager]::GetMediaUrl($img.MediaItem)
            return Convert-ToPublicUrl -Url $url
        }
    } catch {}

    # FileField
    try {
        $file = New-Object Sitecore.Data.Fields.FileField($field, $Item.Database)
        if ($file -and $file.MediaItem) {
            $url = [Sitecore.Resources.Media.MediaManager]::GetMediaUrl($file.MediaItem)
            return Convert-ToPublicUrl -Url $url
        }
    } catch {}

    # Raw <image>/<file> XML fallback (mediaid)
    $raw = $Item[$FieldName]
    if ([string]::IsNullOrWhiteSpace($raw)) { return "" }
    try {
        [xml]$xml = $raw
        $mid = $null
        if ($xml.image) { $mid = $xml.image.mediaid }
        elseif ($xml.file) { $mid = $xml.file.mediaid }
        if ($mid) {
            $mi = Get-Item -Path "master:" -ID $mid -ErrorAction SilentlyContinue
            if ($mi) {
                $url = [Sitecore.Resources.Media.MediaManager]::GetMediaUrl($mi)
                return Convert-ToPublicUrl -Url $url
            }
        }
    } catch {}

    # If the raw value already looks like a URL/path, normalize it
    return Convert-ToPublicUrl -Url $raw
}

function Get-LinkUrl {
    param($Item, [string]$FieldName)

    # Output plain URL (no Excel HYPERLINK) and don't use LinkManager
    $raw = $Item[$FieldName]
    if ([string]::IsNullOrWhiteSpace($raw)) { return "" }

    # If XML, prefer <link url="">
    try {
        [xml]$xml = $raw
        if ($xml.link.url) {
            return Convert-ToPublicUrl -Url ($xml.link.url)
        }
    } catch {
        # Not XML; fall through to raw
    }

    return Convert-ToPublicUrl -Url $raw
}

function Get-MultilistNames {
    param($Item, [string]$FieldName)
    try {
        $ml = New-Object Sitecore.Data.Fields.MultilistField($Item.Fields[$FieldName])
        if ($ml) { return ($ml.GetItems().DisplayName -join ", ") }
    } catch {}

    # Fallback: raw newline-separated IDs
    $raw = $Item[$FieldName]
    if ([string]::IsNullOrWhiteSpace($raw)) { return "" }

    $ids = $raw -split "`n"
    $names = foreach ($id in $ids) {
        $trim = $id.Trim()
        if ($trim) {
            $it = Get-Item -Path "master:" -ID $trim -ErrorAction SilentlyContinue
            if ($it) { $it.DisplayName }
        }
    }
    return ($names -join ", ")
}

function Get-DroplinkName {
    param($Item, [string]$FieldName)
    try {
        $rf = New-Object Sitecore.Data.Fields.ReferenceField($Item.Fields[$FieldName])
        if ($rf -and $rf.TargetItem) { return $rf.TargetItem.DisplayName }
    } catch {}

    $raw = $Item[$FieldName]
    if ([string]::IsNullOrWhiteSpace($raw)) { return "" }

    $ti = Get-Item -Path "master:" -ID $raw -ErrorAction SilentlyContinue
    if ($ti) { return $ti.DisplayName }

    return ""
}

function Get-Checkbox {
    param($Item, [string]$FieldName)
    try {
        $cb = New-Object Sitecore.Data.Fields.CheckboxField($Item.Fields[$FieldName])
        return [bool]$cb.Checked
    } catch {}
    return ($Item[$FieldName] -eq "1")
}

# ------------------ Field Order (as provided) -------------------------

$fields = @(
 "LocationName","BannerImage","AddressLine1","AddressLine2",
 "City","State","ZipCode","PhoneNumber","AboutThisLocation",
 "MedicalSpecialties","Services","Accreditations","Amenities","Hours",
 "ParkingInformationPDF","DepartmentID","VisitTypes","GetDirections","ThumbnailImage",
 "ClinicStatusOverlay","LocationType","LinkOne","LinkTwo","LinkThree",
 "Can Schedule Appointments","SchedulePediAsAdult","SadtSpecialties","Our Experts",
 "AccordionTitle","AccordionContent","AlertStrip","Campus Location",
 "Campus Facilities","Campus Short Description","GoalNameForLocation",
 "EidForLocation"
)

# ------------------ Main (NO RECURSION) -------------------

$path = "master:\content\KelseySeybold\KelseySeyboldGlobal\Data\Locations"
Write-Host "Starting Locations report (NO RECURSION) from: $path"

$items = Get-ChildItem -Path $path

$report    = @()
$index     = 0
$total     = $items.Count
$processed = 0
$skipped   = 0
$failed    = 0

foreach ($item in $items) {
    $index++
    Write-Progress -Activity "Exporting Locations" `
                   -Status "Processing $($item.Name) ($index of $total)" `
                   -PercentComplete (($index / [math]::Max($total,1)) * 100)

    try {
        # Simple filter: LocationName present or template name contains "Location"
        $isLocation = $item.Fields["LocationName"] -ne $null -or $item.TemplateName -like "*Location*"
        if (-not $isLocation) { $skipped++; continue }

        # Only include published locations
        if ($item["__Never publish"] -eq "1") { $skipped++; continue }

        Write-Host "Processing: $($item.Paths.Path)"

        $row = [ordered]@{}

        foreach ($f in $fields) {
            switch ($f) {
                # Images → media URL normalized to public host
                { $_ -in @("BannerImage","ThumbnailImage","ImageOne","ImageTwo","ImageThree",
                           "ImageFour","ImageFive","ImageSix","ImageSeven","ImageEight",
                           "ImageNine","ImageTen") } {
                    $row[$f] = Get-MediaUrl -Item $item -FieldName $f
                    break
                }

                # Files → media URL normalized to public host
                { $_ -in @("DrivingDirectionsPDF","ParkingInformationPDF") } {
                    $row[$f] = Get-MediaUrl -Item $item -FieldName $f
                    break
                }

                # General Links → raw/xml url normalized to public host
                { $_ -in @("DrivingDirectionsLink","ParkingInformationLink",
                           "GetDirections","LinkOne","LinkTwo","LinkThree") } {
                    $row[$f] = Get-LinkUrl -Item $item -FieldName $f
                    break
                }

                # Multilist
                { $_ -in @("MedicalSpecialties","Amenities","VisitTypes","LocationType",
                           "SadtSpecialties","Campus Facilities") } {
                    $row[$f] = Get-MultilistNames -Item $item -FieldName $f
                    break
                }

                # Droplinks
                { $_ -in @("Our Experts","AlertStrip") } {
                    $row[$f] = Get-DroplinkName -Item $item -FieldName $f
                    break
                }

                # Checkboxes
                { $_ -in @("Can Schedule Appointments","SchedulePediAsAdult","Campus Location") } {
                    $row[$f] = Get-Checkbox -Item $item -FieldName $f
                    break
                }

                default {
                    $row[$f] = $item[$f]
                }
            }
        }

        # Canonical Location URL based on Item Name
        $locationUrl = Convert-ItemNameToLocationUrl -ItemName $item.Name
        $row["LocationUrl"] = $locationUrl

        # --- Campus expansion: one row per building ---
        $isCampus = Get-Checkbox -Item $item -FieldName "Campus Location"

        if ($isCampus) {
            $facilityItems = @()
            try {
                $ml = New-Object Sitecore.Data.Fields.MultilistField($item.Fields["Campus Facilities"])
                if ($ml) { $facilityItems = $ml.GetItems() }
            } catch {}

            if ($facilityItems.Count -gt 0) {
                foreach ($facility in $facilityItems) {
                    $campusRow = [ordered]@{}
                    foreach ($key in $row.Keys) { $campusRow[$key] = $row[$key] }

                    # Blank out all inherited fields — building rows only show facility-specific data
                    $facilityOnlyFields = @("LocationName","BannerImage","AddressLine1","AddressLine2",
                                            "City","State","ZipCode","GetDirections","Campus Facilities",
                                            "Campus Location","LocationUrl")
                    foreach ($key in $row.Keys) {
                        if ($key -notin $facilityOnlyFields) {
                            $campusRow[$key] = ""
                        }
                    }

                    # Override LocationName with facility name
                    $facilityName = $facility["CampusFacilityName"]
                    if (-not [string]::IsNullOrWhiteSpace($facilityName)) {
                        $campusRow["LocationName"] = "$($row['LocationName']) - $facilityName"
                    } else {
                        $campusRow["LocationName"] = "$($row['LocationName']) - $($facility.DisplayName)"
                    }
                    # Single facility name for this row
                    $campusRow["Campus Facilities"] = $facility.DisplayName

                    # Map CampusFacility fields to report columns
                    $campusRow["AddressLine1"] = $facility["CampusFacilityAddressLineOne"]
                    $campusRow["AddressLine2"] = $facility["CampusFacilityAddressLineTwo"]
                    $campusRow["City"]         = $facility["CampusFacilityCity"]
                    $campusRow["State"]        = $facility["CampusFacilityState"]
                    $campusRow["ZipCode"]      = $facility["CampusFacilityZipCode"]

                    # BannerImage from facility (Image field)
                    $campusRow["BannerImage"] = Get-MediaUrl -Item $facility -FieldName "CampusFacilityBannerImage"

                    # GetDirections from facility (General Link field)
                    $campusRow["GetDirections"] = Get-LinkUrl -Item $facility -FieldName "CampusFacilityGetDirectionsLink"

                    $report += New-Object psobject -Property $campusRow
                }
                $processed++
            }
            else {
                # Campus with no facilities — emit single row as-is
                $report += New-Object psobject -Property $row
                $processed++
            }
        }
        else {
            $report += New-Object psobject -Property $row
            $processed++
        }
    }
    catch {
        Write-Warning "FAILED on item $($item.Paths.Path) -- $($_.Exception.Message)"
        $failed++
    }
}

Write-Host "Extraction finished. Processed: $processed, Skipped: $skipped, Failed: $failed."

# ------------------ Export & Download ---------------------

# Ensure LocationName is first, LocationUrl second, then remaining fields.
$first = @("LocationName")
$remaining = $fields | Where-Object { $_ -notin $first }
$exportColumns = $first + @("LocationUrl") + $remaining

Write-Host "Building XLSX..."
[byte[]]$xlsx = $report |
    Select-Object -Property $exportColumns |
    ConvertTo-Xlsx

# Post-process: set native EPPlus hyperlinks on LocationName and LocationUrl columns
Write-Host "Applying hyperlinks..."
$ms = New-Object System.IO.MemoryStream(,$xlsx)
$pkg = New-Object OfficeOpenXml.ExcelPackage($ms)
$ws = $pkg.Workbook.Worksheets[1]

$nameCol = $null; $urlCol = $null
for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
    switch ($ws.Cells[1,$c].Value) {
        "LocationName" { $nameCol = $c }
        "LocationUrl"  { $urlCol  = $c }
    }
}

if ($nameCol -and $urlCol) {
    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $url = $ws.Cells[$r,$urlCol].Value
        if (-not [string]::IsNullOrWhiteSpace($url)) {
            $uri = [Uri]$url
            $ws.Cells[$r,$nameCol].Hyperlink = $uri
            $ws.Cells[$r,$urlCol].Hyperlink  = $uri
        }
    }
}

$outMs = New-Object System.IO.MemoryStream
$pkg.SaveAs($outMs)
[byte[]]$xlsx = $outMs.ToArray()
$pkg.Dispose(); $ms.Dispose(); $outMs.Dispose()

$stamp = Get-Date -Format "MM-dd-yyyy_HH-mm-ss"
$filename = "$stamp`_LocationsReport.xlsx"

Write-Host "Download starting: $filename"
Out-Download -Name $filename -InputObject $xlsx
