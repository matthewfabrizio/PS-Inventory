$Hosts = (Get-ChildItem -Path "$PSScriptRoot\Hosts").FullName

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

# god I hate this so much. I'll fix it later
$Build1507 = 0
$Build1511 = 0
$Build1607 = 0
$Build1703 = 0
$Build1709 = 0
$Build1803 = 0
$Build1809 = 0
$Build1903 = 0
foreach ($Build in $Collection.'Windows Build') {
    switch ($Build) {
        '1507' { $Build1507++ }
        '1511' { $Build1511++ }
        '1607' { $Build1607++ }
        '1703' { $Build1703++ }
        '1709' { $Build1709++ }
        '1803' { $Build1803++ }
        '1809' { $Build1809++ }
        '1903' { $Build1903++ }
        Default {}
    }
}

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

$Windows7Count = 0
$Windows8Count = 0
$Windows10Count = 0
foreach ($OS in $Collection.'Windows Edition') {
    if ($OS -match 'Windows 7') {
        $Windows7Count++
    }
    elseif ($OS -match "Windows 8") {
        $Windows8Count++
    }
    elseif ($OS -match "Windows 10") {
        $Windows10Count++
    }
}

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
            New-ChartBar -Name "Windows 7 Professional" -Value $Windows7Count
            New-ChartBar -Name "Windows 8" -Value $Windows8Count
            New-ChartBar -Name "Windows 10" -Value $Windows10Count
        }
        New-HTMLSection -Invisible {
            New-HTMLPanel -Invisible {
                New-HTMLChart -Title "Windows Build Report" {
                    New-ChartBarOptions -Vertical
                    New-ChartLegend -Names "Windows Build" 
                    New-ChartBar -Name "1507" -Value $Build1507
                    New-ChartBar -Name "1511" -Value $Build1511
                    New-ChartBar -Name "1607" -Value $Build1607
                    New-ChartBar -Name "1703" -Value $Build1703
                    New-ChartBar -Name "1709" -Value $Build1709
                    New-ChartBar -Name "1803" -Value $Build1803
                    New-ChartBar -Name "1809" -Value $Build1809
                    New-ChartBar -Name "1903" -Value $Build1903
                }
            }
            New-HTMLPanel -Invisible {
                New-HTMLChart -Title "Windows Build Report" {
                    New-ChartBarOptions -Vertical
                    New-ChartLegend -Names "Windows Build" 
                    New-ChartPie -Name "1507" -Value $Build1507
                    New-ChartPie -Name "1511" -Value $Build1511
                    New-ChartPie -Name "1607" -Value $Build1607
                    New-ChartPie -Name "1703" -Value $Build1703
                    New-ChartPie -Name "1709" -Value $Build1709
                    New-ChartPie -Name "1803" -Value $Build1803
                    New-ChartPie -Name "1809" -Value $Build1809
                    New-ChartPie -Name "1903" -Value $Build1903
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