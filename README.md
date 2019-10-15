<h3 align="center">Get-Inventory</h3>

<div align="center">

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](../../LICENSE)

</div>

## About
Fetch computer information and output to respective JSON file.

## Getting Started
Since this is a small script, a simple `git clone` will get you a copy of the script up and running on your local machine for development and testing purposes.

This script relies on the ActiveDirectory
[module](https://docs.microsoft.com/en-us/powershell/module/addsadministration/?view=win10-ps).

```ps
Import-Module ActiveDirectory
```

## Usage

### Getting Help
The `Help` parameter will get you started on how the script behaves.

```ps
.\Get-ACDCPowerCfg.ps1 -Help
```

## Future Implementations

- [ ] Scan all computers in a specified OU

## Resources

|   |
|---|
| [Parameters](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_advanced_parameters?view=powershell-6) |
| [Get-CimClass](https://docs.microsoft.com/en-us/powershell/module/cimcmdlets/get-cimclass?view=powershell-6) |
| [Get-ADComputer](https://docs.microsoft.com/en-us/powershell/module/addsadministration/get-adcomputer?view=win10-ps) |
