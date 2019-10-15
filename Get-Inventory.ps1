<#
.SYNOPSIS
Fetch computer information.

.DESCRIPTION
Fetch computer information and output to respective JSON file.

.PARAMETER Help
Prints help information.

.EXAMPLE
./Inventory.ps1 -Help

.NOTES
===========================================================================
    Created on:   	9/18/2019
    Updated on:     ***
	Author:    	    Matthew Fabrizio
	Organization: 	*** 
	Filename:     	Get-Inventory.ps1
===========================================================================
#>

# Prints help information.
param(
    [Parameter(Mandatory=$false,
           Position=0,
           HelpMessage="Prints help information.")]
    [switch]
    $Help
)

function Get-Help() {
    "
    AD Query Scan:
        The AD Query menu option allows a user to scan any number of devices starting with a certain string of text.
        For instance, say you want to scan all computers starting with a specific name of 'LAB-' and there are 20 computers ranging from LAB-01 to LAB-20
        The scan will ping all computers starting with that name and return the results, as well as, output all info to respective JSON files.

    Single and Loop Scan are pretty self explanatory; only requiring a hostname to be entered.

    Error Logging:
        If the script runs into an error, the information will be logged to a directory called Log in $PSScriptRoot
        This will contain information specific to what line of code actually failed.

    Excel Formatting:
        This script automatically copies all scanned devices to the clipboard in tab delimited format.
        This allows for simple pasting into Excel.

    Device Type:
        The device type is generated based on device name. The script is catered towards Dell devices and if anything isn't an Optiplex, it's deemed a laptop.
        It's not hard to change and other desktop models can be added if you really want, but overall it's sort of useless.
    "
}

function Write-LogEntry {
    param(
        [parameter(Mandatory=$true, HelpMessage="Value of log entry to be added.")]
        [ValidateNotNullOrEmpty()]
        [string] $Value,

        [parameter(Mandatory=$false, HelpMessage="Name of the log file.")]
        [ValidateNotNullOrEmpty()]
        [string] $FileName = "error.log",

        [Parameter(Mandatory=$false, HelpMessage="Prevents logging to STDOUT")]
        [switch] $OutNull
    )
    
    <# If no Log directory exists, create it #>
    if (!(Test-Path "$PSScriptRoot\Log" -PathType Container)) {
        New-Item -Path "$PSScriptRoot" -Name "Log" -ItemType Directory | Out-Null
    }

    <# Determine log file location #>
    $LogFilePath = Join-Path -Path "$PSScriptRoot" -ChildPath "Log\$($FileName)"

    <# Add value to log file #>
    try {
        if (!$OutNull) { Write-Host $Value }
        Out-File -InputObject "[$(Get-Date -Format g)] => $Value" -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
    }
    catch [System.Exception] {
        Write-Warning -Message "Unable to append log entry to $FileName file"
    }
}

function Get-Menu {
    Clear-Host

    Write-Host "`n----------- MENU -----------"
    Write-Host "[A] : AD Query Scan"
    Write-Host "[L] : Loop Scan"
    Write-Host "[S] : Single Scan"
    Write-Host "[Q] : Quit"
    Write-Host "----------------------------`n"
}

function Get-DeviceInfo() {
    param (
        # Computer(s) to be scanned
        [Parameter(Mandatory=$true)]
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
        else { Write-LogEntry "$Computer unavailable." }
    }

    foreach ($Computer in $Online) {
        try {
            # Load relevant WMI classes
            $Win32_OperatingSystem = (Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem -ErrorAction Stop)
            $Win32_ComputerSystem  = (Get-WmiObject -ComputerName $Computer -Class Win32_ComputerSystem -ErrorAction Stop)
            $Win32_Bios            = (Get-WmiObject -ComputerName $Computer -Class Win32_Bios -ErrorAction Stop)
            $Win32_PhysicalMemory  = (Get-WmiObject -ComputerName $Computer -Class Win32_PhysicalMemory -ErrorAction Stop)
            # $Win32_NetworkAdapterConfiguration = (Get-WmiObject -ComputerName $Computer -Class Win32_NetworkAdapterConfiguration -ErrorAction Stop)
            $Win32_NetworkAdapter = (Get-WmiObject -ComputerName $Computer -Class Win32_NetworkAdapter | Where-Object { $_.Description -notmatch 'wan miniport|microsoft isatap adapter|bluetooth|juniper|ras async adapter|virtual|apple|miniport|tunnel|debug|advanced-n|wireless-n|ndis' } -ErrorAction Stop)
            
            # Manufacturer / Physical
            $Manufacturer = $Win32_ComputerSystem.Manufacturer
            $Model        = $Win32_ComputerSystem.Model
            $Serial       = $Win32_Bios.SerialNumber
            $Memory       = $Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum | ForEach-Object { "{0:N2}" -f ([math]::round(($_.Sum / 1GB),2)) }
            
            # NetBIOS
            $Hostname     = $Win32_OperatingSystem.CSName

            # Operating System
            $Edition      = $Win32_OperatingSystem.Caption
            $OS           = $Win32_OperatingSystem.Version
            
            # Lifespan
            $Age          = $Win32_Bios
            $ReimageDate  = $Win32_OperatingSystem

            # Networking
            $Domain       = $Win32_ComputerSystem.Domain
            $EthMAC       = ($Win32_NetworkAdapter | Where-Object {$_.NetConnectionID -like '*Ethernet*'}).MACAddress
            $WlpMAC       = ($Win32_NetworkAdapter | Where-Object {$_.NetConnectionID -like '*Wi-Fi*'}).MACAddress

            <# AD Queries #>
            # $Hostname = Get-ADComputer $Computer -Properties CN | Select-Object -ExpandProperty CN
            # $OS = Get-ADComputer $Computer -Properties OperatingSystem | Select-Object -ExpandProperty OperatingSystem
            # $IP = Get-ADComputer $Computer -Properties IPv4Address | Select-Object -ExpandProperty IPv4Address

            switch($OS){
                '10.0.10240' {$OS="1507"}
                '10.0.10586' {$OS="1511"}
                '10.0.14393' {$OS="1607"}
                '10.0.15063' {$OS="1703"}
                '10.0.16299' {$OS="1709"}
                '10.0.17134' {$OS="1803"}
                '10.0.17763' {$OS="1809"}
                '10.0.18362' {$OS="1903"}
            }

            # Mainly for Dell devices to remove Inc.
            $Manufacturer = [string]$Manufacturer.replace("Inc.","")
            $Make = $Manufacturer + $Model

            # Mainly for Dell devices, can add more desktop options if necessary
            if ($Model -like '*Optiplex*') { $Type = "Desktop" } 
            else { $Type = "Laptop" }

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
                Hostname    = $Hostname
                Device      = $Make
                Type        = $Type
                Serial      = $Serial
                Edition     = $Edition
                OS          = $OS
                Memory      = $Memory
                Domain      = $Domain
                DecimalAge  = $Age
                Age         = $CalculatedAge
                Reimaged    = $ReimageDate
                ExcelAge    = $ExcelAge
            }

            # Add MAC if exists to object
            if ($EthMAC) { $Properties | Add-Member -NotePropertyName EthMAC -NotePropertyValue $EthMAC }
            if ($WlpMAC) { $Properties | Add-Member -NotePropertyName WlpMAC -NotePropertyValue $WlpMAC }

            # Store computer info in DeviceArray, append additional devices
            [Void]$DeviceArray.Add($Properties)

            # Output computer data to JSON in Hosts directory
            $Properties | ConvertTo-Json | Out-File $PSScriptRoot\Hosts\$Computer.json

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
            $ExceptionLineNumber  = $PSItem.InvocationInfo.ScriptLineNumber
            $ExceptionLineContent = (Get-Content (Split-Path $MyInvocation.ScriptName -Leaf) -TotalCount $ExceptionLineNumber)[-1]
            $ExceptionMessage     = $PSItem.Exception.Message

            Write-LogEntry "[ERROR] : Device [$Computer]
            $ExceptionMessage
            Exception caugh on line $ExceptionLineNumber
            $ExceptionLineContent"
        }
    }
    
    <# Spicy STDOUT #>
    $DeviceArray
}

# If help, help please
if ($Help) { Get-Help; exit }

<# Remove Log directory if it exists; generate on error #>
if (Test-Path -Path "$PSScriptRoot\Log") {
    Remove-Item -Path "$PSScriptRoot\Log" -Recurse
}

<# If there is no Hosts directory to store JSON, create it #>
if (!(Test-Path -Path $PSScriptRoot\Hosts)) {
    New-Item -Path $PSScriptRoot -Name "Hosts" -ItemType Container > $null
}

do {
    <# Print the menu; Select the choice #>
    Get-Menu
    $Choice = Read-Host "Make a selection"

    switch ($Choice) {
        <# Prompt for AD search terms - only starting characters #>
        'a' { 
            $ADFilter = Read-Host "What computer would you like to search for?"
            $ADQuery = (Get-ADComputer -Filter "Name -like '$ADFilter*'" | Select-Object -ExpandProperty Name) -join ","
            $ADQuery = $ADQuery.Split(",").Trim(" ")
            Get-DeviceInfo -ComputerName $ADQuery
        }
        <# [l|L] Prompt for a computer to scan; exit on SIGINT #>
        <# [s|S] Prompt for a computer to scan; exit once complete #>
        { @('l','s') -contains $PSItem } {
            $hostnameExists = $false
            while (!$hostnameExists) {
                Write-Host "`nType 'stop|Stop|Ctrl+c' to exit.`n" -ForegroundColor Yellow
                $Computers = Read-Host "What computer would you like to scan?"
                if ($Computers -contains 'stop') { exit }
                $Computers = $Computers.Split(",").Trim(" ")
                Get-DeviceInfo -ComputerName $Computers
                if ($PSItem -eq 's') { $hostnameExists=$true }
            }
        }
        <# Log user quit prompt; Clear the screen; Exit the application #>
        'q' { Clear-Host; exit }
        default { Write-Host "Invalid menu choice" -ForegroundColor Red }  
    }
} while ($Choice -eq 'q')