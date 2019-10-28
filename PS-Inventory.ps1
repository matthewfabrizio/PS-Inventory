<#
.SYNOPSIS
Fetch computer information.

.DESCRIPTION
Fetch computer information and output to respective JSON file.

.PARAMETER Help
Prints help information.

.EXAMPLE
./PS-Inventory.ps1 -Help

.NOTES
===========================================================================
    Created on:   	2019-09-18
	Author:    	    Matthew Fabrizio
	Organization: 	*** 
    Filename:     	PS-Inventory.ps1
===========================================================================

.LINK
https://github.com/matthewfabrizio
#>

# Requires -Module ActiveDirectory

# Prints help information.
[CmdletBinding()] param ( [Parameter()] [switch] $Help )

<#
    USER CONFIGURATION VARIABLES
    Ex: $DeviceOUs = 'OU1', 'OU2', ... 'OUn'
#>
$DeviceOUs = 'OU1', 'OU2'
$DomainDN = (Get-ADDomain).DistinguishedName

function Get-Help() {
    Clear-Host

    "
    Welcome to PS-Inventory!

    AD Query Scan:
        The AD Query menu option allows a user to scan any number of devices starting with a certain string of text based on caret (^) anchor.
        For instance, say you want to scan all computers starting with a specific name of 'LAB-' and there are 20 computers ranging from LAB-01 to LAB-20
        The scan will ping all computers starting with that name and return the results, as well as, output all info to respective JSON files.

    OU Query Scan:
        The OU Query menu option allows a user to scan any number of devices in a specific OU.
        Because of how different/complex AD environments can be, there needs to be an easy way to include what you want.
        The best method was to add a DeviceOUs variable to the user configuration (can find it above the Get-Help function)
        Make a list of what OUs contain your devices, select OU Query, and the script will handle the rest.

    Single and Loop Scan are pretty self explanatory; only requiring a hostname to be entered.

    Excel Formatting:
        This script automatically copies all scanned devices to the clipboard in tab delimited format.
        This allows for simple pasting into Excel.

    Device Type:
        Device type used to be very Dell specific, but now because of Win32_SystemEnclosure we can make smart decisions!
        I have no idea why I put this into the help section to begin with \_($(0x30C4|%{[char]$_}))_/
    "
}

function Show-Menu {
    Clear-Host

    "`n----------- MENU -----------"
    "[A] : AD Query Scan"
    "[O] : OU Query Scan"
    "[L] : Loop Scan"
    "[S] : Single Scan"
    "[Q] : Quit"
    "----------------------------`n"
}

function Get-DeviceInfo() {
    param (
        # Computer(s) to be scanned
        [Parameter(Mandatory)]
        [string[]]
        $ComputerName
    )
    
    # Demolish anything in the clipboard
    $Null | Set-Clipboard

    $Online = [System.Collections.ArrayList]::New()
    $DeviceArray = [System.Collections.ArrayList]::New()

    foreach ($Computer in $ComputerName.ToUpper()) {
        Write-Progress "Pinging $Computer"

        if (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
            Write-Host "Connection to [$Computer] successful." -ForegroundColor Green
            [Void]$Online.Add($Computer)
        }
        else { Write-Warning -Message "$Computer unavailable." }
    }

    foreach ($Computer in $Online) {
        try {
            # Load relevant WMI classes
            $Win32_OperatingSystem = (Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem -ErrorAction Stop)
            $Win32_ComputerSystem = (Get-WmiObject -ComputerName $Computer -Class Win32_ComputerSystem -ErrorAction Stop)
            $Win32_Bios = (Get-WmiObject -ComputerName $Computer -Class Win32_Bios -ErrorAction Stop)
            $Win32_PhysicalMemory = (Get-WmiObject -ComputerName $Computer -Class Win32_PhysicalMemory -ErrorAction Stop)
            $Win32_NetworkAdapter = (Get-WmiObject -ComputerName $Computer -Class Win32_NetworkAdapter | Where-Object { $_.Description -notmatch 'wan miniport|microsoft isatap adapter|bluetooth|juniper|ras async adapter|virtual|apple|miniport|tunnel|debug|advanced-n|wireless-n|ndis' } -ErrorAction Stop)
            $AV = (Get-WmiObject -ComputerName $Computer -Namespace "root\SecurityCenter2" -Query "SELECT * FROM AntiVirusProduct").displayName
            $Win32_SystemEnclosure = (Get-WmiObject -ComputerName $Computer -ClassName Win32_SystemEnclosure -Namespace 'root\CIMV2' -Property ChassisTypes).ChassisTypes

            # Manufacturer / Physical
            $Manufacturer = $Win32_ComputerSystem.Manufacturer
            $Model = $Win32_ComputerSystem.Model
            $Serial = $Win32_Bios.SerialNumber
            $Memory = $Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum | ForEach-Object { "{0:N2}" -f ([math]::round(($_.Sum / 1GB), 2)) }
            
            # NetBIOS
            $Hostname = $Win32_OperatingSystem.CSName

            # Operating System
            $Edition = $Win32_OperatingSystem.Caption
            $OS = $Win32_OperatingSystem.Version
            
            # Lifespan
            $Age = $Win32_Bios
            $ReimageDate = $Win32_OperatingSystem

            # Networking
            $Domain = $Win32_ComputerSystem.Domain
            $EthMAC = ($Win32_NetworkAdapter | Where-Object { $_.NetConnectionID -like '*Ethernet*' }).MACAddress
            $WlpMAC = ($Win32_NetworkAdapter | Where-Object { $_.NetConnectionID -like '*Wi-Fi*' }).MACAddress

            switch ($OS) {
                '10.0.10240' { $OS = "1507" }
                '10.0.10586' { $OS = "1511" }
                '10.0.14393' { $OS = "1607" }
                '10.0.15063' { $OS = "1703" }
                '10.0.16299' { $OS = "1709" }
                '10.0.17134' { $OS = "1803" }
                '10.0.17763' { $OS = "1809" }
                '10.0.18362' { $OS = "1903" }
            }

            # Mainly for Dell devices to remove Inc.
            $Manufacturer = [string]$Manufacturer.replace("Inc.", "")
            $Make = $Manufacturer + $Model

            # Chassis type source values : https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-systemenclosure
            switch ($Win32_SystemEnclosure) {
                { @(3..7) -contains $PSItem } { $Type = "Desktop" }
                { @(8..16) -contains $PSItem } { $Type = "Laptop" }
                Default { "Unknown" }
            }

            if ($AV.Count -gt 1) { $AntivirusProduct = $AV | Where-Object -FilterScript { $PSItem -ne 'Windows Defender' } }
            else { $AntivirusProduct = $AV }

            # Calculate age and reimage date
            $Age = (New-TimeSpan -Start ($Age.ConvertToDateTime($Age.ReleaseDate).ToShortDateString()) -End $(Get-Date)).Days / 365
            $ReimageDate = ($ReimageDate.ConvertToDateTime($ReimageDate.InstallDate).ToString("MM/dd/yyyy"))

            <# Calculate Decimal Age to Date #>
            $CalculatedAge = ($Age * 365.25)
            $CalculatedAge = (Get-Date).AddDays(-$CalculatedAge).ToString("MM/dd/yyyy")

            # Calculate Excel formula based off TODAYs date
            $ExcelAge = "=ROUND(YEARFRAC(`"$CalculatedAge`", TODAY()), 2)"

            # Big ol object dump
            $Properties = [PSCustomObject]@{
                AntiVirusProduct = $AntivirusProduct
                Hostname         = $Hostname
                Device           = $Make
                Type             = $Type
                Serial           = $Serial
                Edition          = $Edition
                OS               = $OS
                Memory           = $Memory
                Domain           = $Domain
                DecimalAge       = $Age
                Age              = $CalculatedAge
                Reimaged         = $ReimageDate
                ExcelAge         = $ExcelAge
            }

            # Add MAC if exists to object
            if ($EthMAC) { $Properties | Add-Member -NotePropertyName EthMAC -NotePropertyValue $EthMAC }
            if ($WlpMAC) { $Properties | Add-Member -NotePropertyName WlpMAC -NotePropertyValue $WlpMAC }

            # Store computer info in DeviceArray, append additional devices
            [Void]$DeviceArray.Add($Properties)

            $Hosts = (Get-ChildItem -Path "$PSScriptRoot\Hosts").FullName
            if ($Hosts.Count -gt 0) {
                $HostsCheck = Get-Content -Path $Hosts -Raw | ConvertFrom-Json

                foreach ($Item in $HostsCheck) {
                    if ($Properties.Serial -eq $Item.Serial) {
                        Write-Verbose -Message "Removing duplicate entry"
                        Remove-Item -Path "$PSScriptRoot\Hosts\$($Item.Hostname).json"
                    }
                }
            }
            

            $Properties | ConvertTo-Json | Out-File "$PSScriptRoot\Hosts\$Computer.json"

            <# Prep specific properties for Excel pasting; tab delimited #>
            ($Properties | Select-Object * | ForEach-Object {
                    $_.Device
                    $_.Type
                    $_.Hostname
                    $_.Serial
                    $_.Edition
                    $_.OS
                    $_.ExcelAge
                    $_.Reimaged
                }) -join "`t" | Set-Clipboard -Append
        }
        catch [System.Exception] {
            $ExceptionLineNumber = $PSItem.InvocationInfo.ScriptLineNumber
            $ExceptionLineContent = (Get-Content (Split-Path $MyInvocation.ScriptName -Leaf) -TotalCount $ExceptionLineNumber)[-1]
            $ExceptionMessage = $PSItem.Exception.Message

            Write-Warning -Message "[ERROR] : Device [$Computer]
            $ExceptionMessage
            Exception caught on line $ExceptionLineNumber
            $ExceptionLineContent"
        }
    }
    
    <# Spicy STDOUT #>
    $DeviceArray | Sort-Object -Property Hostname | Format-Table -AutoSize
}

# If help, help please
if ($Help) { Get-Help; exit }

<# If there is no Hosts directory to store JSON, create it #>
if (!(Test-Path -Path $PSScriptRoot\Hosts)) {
    New-Item -Path $PSScriptRoot -Name "Hosts" -ItemType Container > $null
}

do {
    Show-Menu
    $Choice = Read-Host "Make a selection"

    switch ($Choice) {
        <# Prompt for AD search terms - only starting characters #>
        'a' {
            Write-Host "`n**Note: AD search terms comply with ^ anchor (Run $($MyInvocation.MyCommand.Name) -Help) for more info**`n" -foregroundColor Yellow
            $ADFilter = Read-Host "What computer would you like to search for?"
            $ADQuery = (Get-ADComputer -Filter "Name -like '$ADFilter*'" | Select-Object -ExpandProperty Name) -join ","
            $ADQuery = $ADQuery.Split(",").Trim(" ")
            Get-DeviceInfo -ComputerName $ADQuery
        }
        'o' {
            # Calculate the OUs from $DeviceOUs
            $OUList = [System.Collections.Generic.List[psobject]]::new()
            foreach ($OU in $DeviceOUs) { $OUList += New-Object -TypeName psobject -Property @{ OU = $OU } }

            if (!$DeviceOUs) { Write-Warning -Message "You need to setup your DeviceOUs variable at the top of $($MyInvocation.MyCommand.Name)"; exit }

            # modify example from https://stackoverflow.com/questions/55152044/making-a-dynamic-menu-in-powershell
            Clear-Host
            Write-Host "Choose an OU to scan" -ForegroundColor Yellow
            Write-Host "**Organizational Units can be defined in the DeviceOUs variable at the top of the script**`n" -ForegroundColor Yellow

            "----------- OU -----------"

            foreach ($MenuItem in $OUList) { '{0} - {1}' -f ($OUList.IndexOf($MenuItem) + 1), $MenuItem.OU }

            "----------------------------`n"

            $ChoiceValid = $false
            while (!$ChoiceValid) {
                $Choice = Read-Host 'Make a selection'
                if ($Choice -notin 1..$OUList.Count) { Write-Warning -Message ('Your choice [ {0} ] is not valid.' -f $Choice) }
                else { $ChoiceValid = $true }
            }

            # Once access to AD test on Get-ADOrganizationUnit and pass to Get-DeviceInfo
            # Will probably have to filter out bogus OUs
            # Match this portion with manual OU switch on other workspace
            $SelectedOU = $OUList.OU[$Choice - 1]
            $SearchBase = "OU=Computers,OU=$SelectedOU,$DomainDN"
            Read-Host "What OU do you want to scan in $SearchBase"
            $OUQuery = (Get-ADComputer -Filter * -SearchBase $SearchBase).Name
            $OUQuery = $OUQuery.Split(",").Trim(" ")
            $OUQuery
        }
        <# [l|L] Prompt for a computer to scan; exit on SIGINT #>
        <# [s|S] Prompt for a computer to scan; exit once complete #>
        { @('l', 's') -contains $PSItem } {
            $HostnameExists = $false

            while (!$HostnameExists) {
                Write-Host "`nType 'stop|Stop|Ctrl+c' to exit.`n" -ForegroundColor Yellow

                # Get device from user; quit on stop
                $Computers = Read-Host "What computer would you like to scan?"
                if ($Computers -contains 'stop') { exit }

                # Combine all computers together
                $Computers = $Computers.Split(",").Trim(" ")
                Get-DeviceInfo -ComputerName $Computers

                # If user selected single scan, auto set to $true
                if ($PSItem -eq 's') { $HostnameExists = $true }
            }
        }
        'q' { Clear-Host; exit }
        default { Write-Warning -Message "Invalid menu choice" }  
    }
} while ($Choice -eq 'q')
