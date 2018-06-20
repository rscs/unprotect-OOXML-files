# Unprotect Spreadsheets

Use this bash script on a *nix machine to unprotect workbooks, worksheets, and VBA code in XLSX and XLSM files. After running, all worksheets will be visible and unprotected.

*NOTE: This script will not work on files protected by an encrypted "Password to open" or "Password to modify" restriction. It will only remove workbook protection and worksheet protections from unencrypted files.*

#### Requirements

The following tools must be installed on your system and available in your path:
* mktemp
* zip and unzip
* hexdump (only required if removing VBA password/restrictions)
* xxd (only required if removing VBA password/restrictions)
