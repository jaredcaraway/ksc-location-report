# KSC Location Report

A Sitecore PowerShell Extensions (SPE) script that exports Kelsey-Seybold Clinic location data to an Excel (.xlsx) report.

## What It Does

- Reads Location items from the Sitecore content tree (no recursion)
- Filters out unpublished locations (`__Never publish = "1"`)
- Normalizes Sitecore media/link URLs to public `kelsey-seybold.com` URLs
- Generates a canonical `LocationUrl` for each location based on item name
- Expands campus locations into one row per building from Campus Facilities
- Exports the result as a timestamped `.xlsx` download

## Usage

Run `LocationReport.ps1` from the Sitecore PowerShell Extensions console or via SPE remoting. The script will process all location items and prompt a file download when complete.

## Column Output

The report includes fields such as LocationName, LocationUrl, BannerImage, address fields, phone, hours, medical specialties, services, amenities, visit types, and more. See the `$fields` array in the script for the full list.
