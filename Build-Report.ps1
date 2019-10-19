<# USER CONFIGURATION VARIABLES #>
$MinimumOSBuild = 1803 # Lowest build before conditional formatting is applied
$MaximumAge     = 5.9  # Maximum age before conditional formatting is applied
$UnsupportedOS  = 'Microsoft Windows 7 Professional' # Problem OS to conditionally format

# Domain applies to $DomainWarning, also possible to use (Get-ADDomain).Forest
$Domain         = 'YOURDOMAIN.org'

# If you don't want any warnings on the Devices tab, set values to $false
$MinimumOSBuildWarning  = $true  # Will warn if any devices are running builds lower than $MinimumOSBuild (above)
$UnsupportedOSWarning   = $true  # Will warn if any devices are running $UnsupportedOS (above)
$DomainWarning          = $false # Will warn if any devices are on a WORKGROUP

# Colors for Reports tab
$TotalDevicesColor = 'Green'
$AntivirusReportColor = 'Orange'
$WindowsEditionReportColor = 'Blue'
$WindowsBuildReportColor = 'Red'
############################################################################################################

# PSWriteHTML Module sources:
# https://evotec.xyz/easy-way-to-create-diagrams-using-powershell-and-pswritehtml/
# https://github.com/EvotecIT/PSWriteHTML

$Hosts   = (Get-ChildItem -Path "$PSScriptRoot\Hosts").FullName
$Content = Get-Content -Path $Hosts -Raw | ConvertFrom-Json

$Collection = [System.Collections.Generic.List[psobject]]::new()

foreach ($Item in $Content) {
    $Collection.Add([PSCustomObject]@{
        Hostname            = $Item.Hostname
        Device              = $Item.Device
        Type                = $Item.Type
        'Serial Number'     = $Item.Serial
        'Windows Edition'   = $Item.Edition
        'Windows Build'     = $Item.OS
        Memory              = $Item.Memory
        Domain              = $Item.Domain
        'Decimal Age'       = $Item.DecimalAge
        Age                 = $Item.Age
        'Reimaged Date'     = $Item.Reimaged
        'Excel Age Formula' = $Item.ExcelAge
        Antivirus           = $Item.AntiVirusProduct
    })
}

# Hastable enumertaion source : https://mjolinor.wordpress.com/2012/01/29/powershell-hash-tables-as-accumulators/

#region Counters
# These counters will break on legacy versions of PS-Inventory
# If you make core changes to what PS-Inventory outputs, then godspeed
$SerialCount            = [ordered]@{}
$DomainCount            = [ordered]@{}
$WindowsBuildCount      = [ordered]@{}
$AntivirusCount         = [ordered]@{}
$WindowsEditionCount    = [ordered]@{}
$DeviceTypeCount        = [ordered]@{}

$Collection | ForEach-Object {
    $SerialCount[$PSItem.'Serial Number']++
    $DomainCount[$PSItem.Domain]++
    $AntivirusCount[$PSItem.Antivirus]++
    $WindowsEditionCount[$PSItem.'Windows Edition']++
    $DeviceTypeCount[$PSItem.Type]++

    if ($PSItem.'Windows Build' -eq '6.3.9600') {
        $WindowsBuildCount['Windows 8.1 Pro']++
    }
    elseif ($PSItem.'Windows Build' -eq '6.1.7601') {
        $WindowsBuildCount['Windows 7 Professional']++
    }
    else {
        $WindowsBuildCount[$PSItem.'Windows Build']++  
    }
}

New-HTML -Name "Inventory Report" -FilePath "$PSScriptRoot\Report.html" -ShowHTML {
    New-HTMLTab -Name "Reports" {
        # Duplicate serial number warning
        foreach ($Serial in $SerialCount.GetEnumerator()) {
            if ($Serial.Value -eq 2) {
                New-HTMLPanel -Invisible {
                    New-HTMLToast -TextHeader "Duplicate Serial Numbers" -Text "Serial number [$($Serial.Name)] found twice in host files." -BarColorLeft OrangeRed -IconSolid info-circle -IconColor OrangeRed
                }
            }
        }

        # Start device reporting
        New-HTMLSection -Invisible {
            New-HTMLPanel -Invisible {
                New-HTMLChart -Title "Total Devices" {
                    New-ChartBarOptions -Vertical
                    New-ChartLegend -Names "Number of Devices" -Color $TotalDevicesColor
                    New-ChartBar -Name "Total Devices" -Value $Collection.Count
                }
            }

            New-HTMLPanel -Invisible {
                New-HTMLChart -Title "Device Types" {
                    New-ChartBarOptions -Vertical
                    New-ChartLegend -Names "Device Types" 
                    foreach ($Type in $DeviceTypeCount.GetEnumerator()) {
                        New-ChartPie -Name $Type.Name -Value $Type.Value
                    }
                }
            }
        } 

        New-HTMLChart -Title "Antivirus Report" {
            New-ChartLegend -Names "Antivirus Product Report" -Color $AntivirusReportColor
            foreach ($Antivirus in $AntivirusCount.GetEnumerator()) {
                New-ChartBar -Name $Antivirus.Name -Value $Antivirus.Value
            }
        }
        New-HTMLChart -Title "Windows Edition Report" {
            New-ChartBarOptions -Vertical
            New-ChartLegend -Names "Windows Edition" -Color $WindowsEditionReportColor
            foreach ($Edition in $WindowsEditionCount.GetEnumerator()) {
                New-ChartBar -Name $Edition.Name -Value $Edition.Value
            }
        }
        New-HTMLSection -Invisible {
            New-HTMLPanel -Invisible {
                New-HTMLChart -Title "Windows Build Report" {
                    New-ChartBarOptions -Vertical
                    New-ChartLegend -Names "Windows Build" -Color $WindowsBuildReportColor
                    foreach ($Build in $WindowsBuildCount.GetEnumerator()) {
                        New-ChartBar -Name $Build.Name -Value $Build.Value
                    }
                }
            }
        } 
    }
    
    New-HTMLTab -Name "Devices" {
        if ($UnsupportedOSWarning) {
            if ($WindowsBuildCount.'Windows 7 Professional' -gt 0) {
                New-HTMLToast -TextHeader "Windows 7 Support Ending Soon!" -Text "You have devices still running Windows 7.<br>
                <a href='https://support.microsoft.com/en-us/help/4057281/windows-7-support-will-end-on-january-14-2020 target='_blank''>Windows 7 EOL</a>" -BarColorLeft OrangeRed -IconSolid info-circle -IconColor OrangeRed
            }
        }
        
        if ($MinimumOSBuildWarning) {
            if ($WindowsBuildCount.GetEnumerator() | Where-Object {$_.Name -lt $MinimumOSBuild} ) {
                New-HTMLToast -TextHeader "Out of date devices!" -Text "You have devices still running builds lower than $MinimumOSBuild." -BarColorLeft OrangeRed -IconSolid info-circle -IconColor OrangeRed
            }
        }
        
        if ($DomainWarning) {
            if ($DomainCount.GetEnumerator() | Where-Object {$_.Name -eq 'WORKGROUP'}) {
                New-HTMLToast -TextHeader "Devices not joined to $Domain" -Text "You have devices still joined to a WORKGROUP." -BarColorLeft OrangeRed -IconSolid info-circle -IconColor OrangeRed
            }
        }
        
        New-HTMLSection -Name "All Device Information" {
            New-HTMLTable -HideFooter -DataTable $Collection {
                New-TableCondition -Name 'Decimal Age' -ComparisonType number -Operator gt -Value $MaximumAge -Color White -BackgroundColor Red -Row
                New-TableCondition -Name 'Windows Build' -ComparisonType number -Operator lt -Value $MinimumOSBuild -Color White -BackgroundColor Orange
                New-TableCondition -Name 'Windows Edition' -ComparisonType string -Operator eq -Value $UnsupportedOS -Color White -BackgroundColor Yellow
            }
        }
    }

    New-HTMLTab -Name "Reimaged Information" {
        New-HTMLCalendar {
            foreach ($Item in $Collection) {
                New-CalendarEvent -StartDate $Item.'Reimaged Date' -Title "$($Item.Hostname) reimaged"
            }
        }
    }
}