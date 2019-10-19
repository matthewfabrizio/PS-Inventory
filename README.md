<h3 align="center">Get-Inventory</h3>

<div align="center">

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

</div>

## About
Fetch computer information and output to respective JSON file. Afterwards, build a reporting dashboard with PSWriteHTML.

## Getting Started
Since this is a small script, a simple `git clone` will get you a copy of the script up and running on your local machine for development and testing purposes.

This script relies on the
[ActiveDirectory](https://docs.microsoft.com/en-us/powershell/module/addsadministration/?view=win10-ps).
and
[PSWriteHTML](https://github.com/EvotecIT/PSWriteHTML)
modules.

```ps
Import-Module ActiveDirectory
Import-Module PSWriteHTML
```

It is also advantageous to install the RSAT tools for Active Directory.

```ps
Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability -Online
```

## Usage

To start, run `.\PS-Inventory.ps1`. The text based menu will help you with the rest.

Once you have at least one computer scanned, run the `Build-Report.ps1` script. This will generate a report called `report.html` in the script root and will open automatically once finished. You can send this report to whoever via one single file!

If you want to modify the report builder, open the file and change a few configuration variables at the top (mainly for colors and ignoring some warnings).

Enjoy!

### Getting Help
The `Help` parameter will get you started on how the script behaves.

```ps
.\PS-Inventory.ps1 -Help
```

## Future Implementations

- [ ] Scan all computers in a specified OU
- [ ] Add report building into main script with parameter

## Resources

|   |
|---|
| [Parameters](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_advanced_parameters?view=powershell-6) |
