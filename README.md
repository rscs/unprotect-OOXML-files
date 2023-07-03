# Unprotect Office Open XML (OOXML) Files

## Overview

Use these base scripts on a *nix machine to unprotected workbooks, worksheets, VBA code, and documents in OOXML files.

These scripts will not work on files protected by an encrypted "Password to open" or "Password to modify" restriction.  It will only remove protections from unencrypted files.*

The scripts will not modify the existing file.  A new file will be created with the phrase `-Unprotected` appended to the original filename.

## Requirements

The following tools must be installed on your system and available in your path:
* mktemp
* zip and unzip
* hexdump (only required if removing VBA password/restrictions)
* xxd (only required if removing VBA password/restrictions)

## Spreadsheets (XLSX and XLSM files)

Use the `spreadsheet.sh` script for spreadsheet files.  After running, all worksheets will be visible and unprotected.

## Documents (DOCX files)

Use the `document.sh` script for document files.  After running, the document will be unprotected.
