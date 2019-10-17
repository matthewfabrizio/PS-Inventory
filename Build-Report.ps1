$Hosts   = (Get-ChildItem -Path "$PSScriptRoot\Hosts").FullName
$Content = Get-Content -Path $Hosts -Raw | ConvertFrom-Json

$Collection = [System.Collections.Generic.List[psobject]]::new()

foreach ($Item in $Content) {
    $Collection.Add([PSCustomObject]@{
        Hostname = $Item.Hostname
        Device = $Item.Device
        Type = $Item.Type
        'Serial Number' = $Item.Serial
        'Windows Edition' = $Item.Edition
        'Windows Build' = $Item.OS
        Memory = $Item.Memory
        Domain = $Item.Domain
        'Decimal Age' = $Item.DecimalAge
        Age = $Item.Age
        'Reimaged Date' = $Item.Reimaged
        'Excel Age Formula' = $Item.ExcelAge
        Symantec = $Item.Symantec
    })
}

# Hastable enumertaion source : https://mjolinor.wordpress.com/2012/01/29/powershell-hash-tables-as-accumulators/

#region OS Build Version Count

$WindowsBuildCount = [ordered]@{}

$Collection | ForEach-Object {
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

#endregion

# $AntivirusCount = @{}

# $Collection | ForEach-Object {
#     if ($AV -eq $true) {
#         $AntivirusCount[$PSItem.Symantec]++
#     }
# }

$SymantecCount = 0
$OtherCount = 0
foreach ($AV in $Collection.Symantec) {
    if ($AV -eq $true) {
        $SymantecCount++
    }
    else {
        $OtherCount++
    }
}

#region OS Edition Count

$WindowsEditionCount = @{}

$Collection | ForEach-Object {
    $WindowsEditionCount[$PSItem.'Windows Edition']++  
}

#endregion

New-HTML -Name "Inventory Report" -FilePath "$PSScriptRoot\Report.html" -ShowHTML {
    New-HTMLTab -Name "Reports" {
        New-HTMLChart -Title "Total Devices" {
            New-ChartLegend -Name "Number of Devices"
            New-ChartBar -Name "Total Devices" -Value $Collection.Count
        }
        New-HTMLChart -Title "Antivirus Report" {
            New-ChartBarOptions -Vertical
            New-ChartLegend -Names "Antivirus"
            New-ChartBar -Name "Symantec" -Value $SymantecCount
            New-ChartBar -Name "Other" -Value $OtherCount
        }
        New-HTMLChart -Title "Windows Edition Report" {
            New-ChartBarOptions -Vertical
            New-ChartLegend -Names "Windows Edition"
            foreach ($Edition in $WindowsEditionCount.GetEnumerator()) {
                New-ChartBar -Name $Edition.Name -Value $Edition.Value
            }
        }
        New-HTMLSection -Invisible {
            New-HTMLPanel -Invisible {
                New-HTMLChart -Title "Windows Build Report" {
                    New-ChartBarOptions -Vertical
                    New-ChartLegend -Names "Windows Build" 
                    foreach ($Build in $WindowsBuildCount.GetEnumerator()) {
                        New-ChartBar -Name $Build.Name -Value $Build.Value
                    }
                }
            }
        } 
    }
    
    New-HTMLTab -Name "Devices" {
        New-HTMLSection -Name "All Device Information" {
            New-HTMLTable -HideFooter -DataTable $Collection {
                New-TableCondition -Name 'Decimal Age' -ComparisonType number -Operator gt -Value 6 -Color White -BackgroundColor Red -Row
                New-TableCondition -Name 'Windows Build' -ComparisonType number -Operator lt -Value 1803 -Color White -BackgroundColor Orange
                New-TableCondition -Name 'Windows Edition' -ComparisonType string -Operator eq -Value "Microsoft Windows 7 Professional" -Color White -BackgroundColor BurlyWood
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