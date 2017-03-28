<#
.SYNOPSIS
    AD cleanup script created to cleanup inactive stale computers off of a OU and sub OU's after a time of inactivity.
.DESCRIPTION
    Script has three functions: 
    [Get-StaleADComputers] - Finds all the potential stale computers on the network inactive for 90 days and exports the computers to a .CSV file.
    [Disable-staleADComputers] - Disables all stale AD computers that have been inactive within 90 days and export the log of the computers disabled.
    [Delete-disableADComputers] - Deletes all disabled stale AD computers that have been inactive within 120 days and export the log of the computers deleted.

.NOTES
    File Name      : Invoke-CleanCOSAD.ps1
    Author         : Tyler Hardy (tylerhardy@weber.edu)
    Prerequisite   : PowerShell V3, RSAT
    Copyright 2017 - Tyler Hardy
#>

function Invoke-CleanCOSAD {
    [CmdletBinding()]
    param(
        [Parameter()]
        # Change these variables to your enivronment
        [int]$disableDay=90,
        [int]$deleteDay=120,
        [string]$searchOU="OU=Corporate Computers,DC=corp,DC=agricorp,DC=com",
        [string]$excludeOU1="*OU=servers,OU=Corporate Computers,DC=corp,DC=agricorp,DC=com",
        [string]$excludeOU2="*OU=vm,OU=Corporate Computers,DC=corp,DC=agricorp,DC=com",
        [string]$default_log="$env:userprofile\Documents\stale_computer_report_$((Get-Date).ToString('MM-dd-yyyy')).csv",
        [string]$disable_log="$env:userprofile\Documents\disabled_stale_computer_report_$((Get-Date).ToString('MM-dd-yyyy')).csv",
        [string]$delete_log="$env:userprofile\Documents\deleted_disabled_stale_computer_report_$((Get-Date).ToString('MM-dd-yyyy')).csv"
    )
    PROCESS {
        #####################################
        # Functions
        #####################################
        function Show-Menu {
            param (
                [int]$disableDay,
                [int]$deleteDay
            )
            $Today = Get-Date
            $compareDateDisable = [DateTime]::Today.AddDays(-($disableDay))
            $compareDateDelete = [DateTime]::Today.AddDays(-($deleteDay))
            Clear-Host
            Write-Host "`n================ Clean COS AD Menu v1.0 ================`n"
            Write-Host "Todays date: "$Today.ToShortDateString()            
            Write-Host "Disable cutoff date: "$compareDateDisable.ToShortDateString()           
            Write-Host "Delete cutoff date: "$compareDateDelete.ToShortDateString()             
            Write-Host "`n========================================================`n"            
            Write-Host "1: Press '1' To run report (Get-staleADComputers)."
            Write-Host "2: Press '2' To disable stale computers (Disable-staleADComputers)."
            Write-Host "3: Press '3' To delete disabled computers (Remove-disableADComputers)."
            Write-Host "Q: Press 'Q' to quit."
        }
        function Get-staleADComputers {
            [CmdletBinding()]
            param(
                [int]$disableDay,
                [string]$searchOU,
                [string]$excludeOU1,
                [string]$excludeOU2,
                [string]$default_log
            )
            PROCESS {
                # Local Variables
                $compareDate = [DateTime]::Today.AddDays(-($disableDay))
                $results = $null
                try {
                    # Build the report
                    $stalePCReport = Get-ADComputer -Filter {(isCriticalSystemObject -eq $False)} -Properties Name,PwdLastSet,WhenChanged,SamAccountName,LastLogonTimeStamp,Enabled,IPv4Address,`
                    operatingsystem,operatingsystemversion,serviceprincipalname -SearchScope Subtree -SearchBase $SearchOU -ErrorAction Stop |
                    Where-Object {($_.DistinguishedName -notlike $excludeOU1) -and ($_.DistinguishedName -notlike $excludeOU2)} |
                    Select-Object Name,operatingsystem,operatingsystemversion,Enabled,@{Name="PwdLastSet";Expression={[datetime]::FromFileTime($_.PwdLastSet)}},`
                    @{Name="LastLogonTimeStamp";Expression={[datetime]::FromFileTime($_.LastLogonTimeStamp)}},WhenChanged,IPv4Address, `
                    @{Name="Stale";Expression={if((($_.pwdLastSet -lt $compareDate.ToFileTimeUTC()) -and ($_.pwdLastSet -ne 0)`
                    -and ($_.LastLogonTimeStamp -lt $compareDate.ToFileTimeUTC()) -and ($_.LastLogonTimeStamp -ne 0)) `
                    -and (!($_.serviceprincipalname -like "*MSClusterVirtualServer*"))){$True}else{$False}}}, `
                    @{Name="ParentOU";Expression={$_.distinguishedname.Substring($_.samaccountname.Length + 3)}}
                    $results = ($stalePCReport | Where-Object {$_.Stale -eq "true"}).count
                    $stalePCReport | Export-Csv -Append $default_log -NoTypeInformation
                    return $results
                }
                catch {
                    Write-Warning "[CATCH] Error, command (Get-staleADComputers) failed: $($_.Exception)"
                    Write-Warning $error[0].Exception.GetType().FullName
                }
            }
        }
        function Disable-staleADComputers {
            [CmdletBinding()]
            param(
                [int]$disableDay,
                [string]$searchOU,
                [string]$excludeOU1,
                [string]$excludeOU2,
                [string]$disable_log
            )
            PROCESS {
                # Local Variables
                $compareDate = ([DateTime]::Today.AddDays(-($disableDay))).ToFileTimeUTC()
                $disabledComputers = @()
                try {
                    $disableStalePC = Get-ADComputer -Filter {(pwdLastSet -lt $compareDate) -and (LastLogonTimeStamp -lt $compareDate) -and (Enabled -eq $True) -and (IsCriticalSystemObject -ne $True)}`
                    -Properties Name,PwdLastSet,WhenChanged,SamAccountName,LastLogonTimeStamp,Enabled,Description,IPv4Address,`
                    operatingsystem,operatingsystemversion,serviceprincipalname,DistinguishedName -SearchScope Subtree -SearchBase $SearchOU |
                    Where-Object {($_.DistinguishedName -notlike $excludeOU1) -and ($_.DistinguishedName -notlike $excludeOU2) -and (!($_.serviceprincipalname -like "*MSClusterVirtualServer*"))} |
                    ForEach-Object{
                        Set-ADComputer -Identity $_.Name -Description ($_.Description + " ::Disabled due to inactivity on $(Get-Date -Format d)::") -Enabled $false
                        $rc = New-Object PSObject
                        $rc | Add-Member -type NoteProperty -name Computer -Value $_.Name
                        $rc | Add-Member -type NoteProperty -name OS -Value $_.operatingsystem
                        $rc | Add-Member -type NoteProperty -name LastLogin -Value ([DateTime]::FromFileTime($_.LastLogonTimeStamp))
                        $rc | Add-Member -type NoteProperty -name PwdLastSet -Value ([DateTime]::FromFileTime($_.PwdLastSet))
                        $rc | Add-Member -type NoteProperty -name Status -Value "Disabled"
                        $rc | Add-Member -type NoteProperty -name Date -Value $(Get-Date -Format d)
                        $rc | Add-Member -type NoteProperty -name OU -Value ($_.distinguishedname.Substring($_.samaccountname.Length + 3))
                        $disabledComputers += $rc
                        remove-variable rc
                    }
                    $results = $disabledComputers.count
                    if ($results -gt 0) {
                        $disabledComputers | Export-Csv -Append $disable_log -NoTypeInformation
                        return $results
                    }
                    else {
                        return $false
                    }
                }
                catch {
                    Write-Warning "[CATCH] Error, command (Disable-staleADComputers) failed: $($_.Exception)"
                    Write-Warning $error[0].Exception.GetType().FullName
                }
            }
        }
        function Remove-disableADComputers {
            [CmdletBinding()]
            param(
                [int]$deleteDay,
                [string]$searchOU,
                [string]$excludeOU1,
                [string]$excludeOU2,
                [string]$delete_log
            )
            PROCESS {
                # Local Variables
                $compareDate = ([DateTime]::Today.AddDays(-($deleteDay))).ToFileTimeUTC()
                $deletedComputers = @()
                try {
                    Get-ADComputer -Filter {(pwdLastSet -lt $compareDate) -and (LastLogonTimeStamp -lt $compareDate) -and (Enabled -eq $false) -and (IsCriticalSystemObject -ne $True)} `
                    -Properties Name,pwdLastSet,operatingsystem,LastLogonTimeStamp,distinguishedname,servicePrincipalName,samaccountname -SearchScope Subtree -SearchBase $SearchOU |
                    Where-Object {($_.DistinguishedName -notlike $excludeOU1) -and ($_.DistinguishedName -notlike $excludeOU2) -and (!($_.serviceprincipalname -like "*MSClusterVirtualServer*"))} |
                    ForEach-Object{
                        try {
                            Remove-ADComputer -Identity $_.Name
                            $rc = New-Object PSObject
                            $rc | Add-Member -type NoteProperty -name Computer -Value $_.Name
                            $rc | Add-Member -type NoteProperty -name OS -Value $_.operatingsystem
                            $rc | Add-Member -type NoteProperty -name LastLogin -Value ([DateTime]::FromFileTime($_.LastLogonTimeStamp))
                            $rc | Add-Member -type NoteProperty -name PwdLastSet -Value ([DateTime]::FromFileTime($_.PwdLastSet))
                            $rc | Add-Member -type NoteProperty -name Status -Value "Deleted"
                            $rc | Add-Member -type NoteProperty -name Date -Value $(Get-Date -Format d)
                            $rc | Add-Member -type NoteProperty -name OU -Value ($_.distinguishedname.Substring($_.samaccountname.Length + 3))
                            $deletedComputers += $rc
                            remove-variable rc
                        }
                        catch [Microsoft.ActiveDirectory.Management.ADException] { 
                            Remove-ADObject -Identity $_.Name -Recursive
                            $rc = New-Object PSObject
                            $rc | Add-Member -type NoteProperty -name Computer -Value $_.Name
                            $rc | Add-Member -type NoteProperty -name OS -Value $_.operatingsystem
                            $rc | Add-Member -type NoteProperty -name LastLogin -Value ([DateTime]::FromFileTime($_.LastLogonTimeStamp))
                            $rc | Add-Member -type NoteProperty -name PwdLastSet -Value ([DateTime]::FromFileTime($_.PwdLastSet))
                            $rc | Add-Member -type NoteProperty -name Status -Value "Deleted"
                            $rc | Add-Member -type NoteProperty -name Date -Value $(Get-Date -Format d)
                            $rc | Add-Member -type NoteProperty -name OU -Value ($_.distinguishedname.Substring($_.samaccountname.Length + 3))
                            $deletedComputers += $rc
                            remove-variable rc
                        }
                        catch {
                            Write-Warning "[CATCH] Error, command (Remove-ADComputer; Remove-ADObject) failed: $($_.Exception)"
                            Write-Warning $error[0].Exception.GetType().FullName
                        }
                    }
                    $results = $deletedComputers.count
                    if ($results -gt 0) {
                        $deletedComputers | Export-Csv -Append $delete_log -NoTypeInformation
                        return $results
                    }
                    else {
                        return $false
                    }

                }
                catch {
                    Write-Warning "[CATCH] Error, command (Remove-disableADComputers) failed: $($_.Exception)"
                    Write-Warning $error[0].Exception.GetType().FullName
                }
            }
        }
        
        #####################################
        # Script
        #####################################

        # Menu logic
        do {
            Show-Menu $disableDay $deleteDay
            $yes = $null
            $no = $null
            $input = Read-Host "Please make a selection"
            if (($input -ne 1) -and ($input -ne 2) -and ($input -ne 3) -and ($input -ne 'q')) {
                do {
                    $input = Read-Host "Please make a valid selection"
                } until (($input -eq 1) -or ($input -eq 2) -or ($input -eq 3) -or ($input -eq 'q'))
            }
            switch ($input) {
                1 {
                    # Runs Get-staleADComputers report
                    Clear-Host
                    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Runs [Get-staleADComputers] to create COS AD stale computers report."
                    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Report does not run, returns to main menu."
                    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
                    $result = $host.ui.PromptForChoice("Run AD Stale Computers Report", "Do you want to run the AD stale computers report?", $options, 0)
                    switch ($result) {
                        0 {
                            "`nYou selected Yes, running report..."
                            $staleCount = Get-staleADComputers $disableDay $searchOU $excludeOU1 $excludeOU2 $default_log
                            if ($staleCount) {
                                "Report successfully ran and exported to $default_log"
                                "There are [$staleCount] stale computers reported in [$searchOU] OU and sub OU's`n"
                            }
                            else {
                                "Report successfully ran and exported to $default_log"
                                "No stale computers found"
                            }
                        }
                        1 {
                            "`nYou selected No, returning to the main menu.`n"
                        }
                    }
                } 
                2 {
                    # Runs Disable-staleADComputers function to disable inactive computers
                    Clear-Host
                    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Runs [Disable-staleADComputers] to disable computers inactive for $disableDay days."
                    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","No computers disabled, returns to main menu."
                    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
                    $result = $host.ui.PromptForChoice("Disable All Stale Computers", "Do you want to disable all stale computers in COS AD?", $options, 1)
                    switch ($result) {
                        0 {
                            "`nYou selected Yes, disabling computers..."
                            $disableCount = Disable-staleADComputers $disableDay $searchOU $excludeOU1 $excludeOU2 $disable_log
                            if ($disableCount) {
                                "Disabled computers log successfully exported to $disable_log"
                                "Disabled [$disableCount] computers in [$searchOU] OU and sub OU's`n"
                            }
                            else {
                                "No computers to be disabled"
                            }
                        }
                        1 {
                            "`nYou selected No, returning to the main menu.`n"
                        }
                    }
                } 
                3 {
                    # Runs Remove-disableADComputers function to delete disabled inactive computers
                    Clear-Host
                    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Runs [Remove-disableADComputers] to delete computers inactive for $deleteDay days."
                    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","No computers deleted, returns to main menu."
                    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
                    $result = $host.ui.PromptForChoice("Delete All Stale Disabled Computers", "Do you want to delete all stale disabled computers in COS AD?", $options, 1)
                    switch ($result) {
                        0 {
                            "`nYou selected Yes, deleting computers..."
                            $deleteCount = Remove-disableADComputers $deleteDay $searchOU $excludeOU1 $excludeOU2 $delete_log
                            if ($deleteCount) {
                                "Deleted computers log successfully exported to $delete_log"
                                "Deleted [$deleteCount] computers in [$searchOU] OU and sub OU's`n"
                            }
                            else {
                                "No computers to be deleted"
                            }

                        }
                        1 {
                            "`nYou selected No, returning to the main menu.`n"
                        }
                    }
                } 
                "q" {
                    # Quits the script
                    Clear-Host
                    return
                }
            }
            pause
        }
        until ($input -eq 'q')
    }   
}

# Run the function
Invoke-CleanCOSAD