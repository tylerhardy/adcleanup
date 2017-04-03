#requires -version 3.0
#requires -Modules ActiveDirectory

function Invoke-ADCleanup {
    <#
    .SYNOPSIS
        Finds/disables/deletes inactive computers from the chosen OU.
    .DESCRIPTION
        Script has three functions: 
        [Get-staleADComputer] - Finds all the potential stale computers on the 
        network inactive for 90 days and exports the computers to a .CSV file.
        [Disable-staleADComputer] - Disables all stale AD computers that have 
        been inactive within 90 days and export the log of the computers 
        disabled.
        [Delete-disabledSADComputers] - Deletes all disabled stale AD computers 
        that have been inactive within 120 days and export the log of the 
        computers deleted.

    .NOTES
        File Name       : Invoke-ADCleanup.ps1
        Date            : 04/03/2017
        Version         : 1.1
        Author          : Tyler Hardy (tylerhardy@gmail.com)
        Prerequisite    : PowerShell V3, RSAT
        Copyright 2017  - Tyler Hardy
    #>

    [CmdletBinding()]
    Param (
        # Age to disable inactive computers.
        [Parameter()]
            [int]
            $disableAge=90,
        # Age to delete previously disabled, inactive computers.
        [Parameter()]
            [int]
            $deleteAge=120,
        # Stale Computer Report file name [with or without full path].
        [Parameter()]
            [string]
            $StaleComputer_Report='',
        # Disabled Stale Computer Report file name [with or without full path].
        [Parameter()]
            [string]
            $Disabled_SC_Report='',
        # Deleted Disabled Stale Computer Report file name [with or without full path].
        [Parameter()]
            [string]
            $Deleted_DSC_Report='',
        # The OU to search.
        [Parameter()]
            [string]
            $searchOU='',
        # The OU to search.
        [Parameter()]
            [string[]]
            $exclusionList=''
    )
    BEGIN {
        #####################################
        # Global Variables - Change these variables to your environment
        #####################################
        $ReportFolder = "$env:userprofile\Documents"
        $TimeStamp = (Get-Date).ToString('yyyy-MM-dd')
        $FilePrefix = Join-Path -Path $ReportFolder -ChildPath $TimeStamp

        if ([string]::IsNullOrWhiteSpace($StaleComputer_Report)) {
            $StaleComputer_Report = "$($FilePrefix)_stale_computer_report.csv"
        }
        if ([string]::IsNullOrWhiteSpace($Disabled_SC_Report)) {
            $Disabled_SC_Report = "$($FilePrefix)_disabled_stale_computer_report.csv"
        }
        if ([string]::IsNullOrWhiteSpace($Deleted_DSC_Report)) {
            $Deleted_DSC_Report = "$($FilePrefix)_deleted_disabled_stale_computer_report.csv"
        }

        if ([string]::IsNullOrWhiteSpace($searchOU)) {
            $searchOU="OU=Corporate Computers,DC=corp,DC=agricorp,DC=com"
        }

        if ([string]::IsNullOrWhiteSpace($exclusionList)) {
            $exclusionList="OU=vm","OU=servers","OU=exclude3"
        }

        #####################################
        # Functions
        #####################################
        function Get-staleADComputer {
            [CmdletBinding()]
            Param (
                [int]$disableAge,
                [string]$searchOU,
                [string]$StaleComputer_Report,
                [string[]]$exclusionList
            )
            PROCESS {
                # Local Variables
                $compareDate = [DateTime]::Today.AddDays(-($disableAge))
                $results = $null
                try {
                    # Build the report
                    $stalePCReport = Get-ADComputer -Filter {(isCriticalSystemObject -eq $False)} -Properties Name,PwdLastSet,WhenChanged,SamAccountName,LastLogonTimeStamp,Enabled,IPv4Address,`
                    operatingsystem,operatingsystemversion,serviceprincipalname -SearchScope Subtree -SearchBase $SearchOU -ErrorAction Stop |
                    Where-Object {((ForEach-Object {$_.DistinguishedName -notmatch ($exclusionList -join "|")}) -eq 'true')} |
                    Select-Object Name,operatingsystem,operatingsystemversion,Enabled,@{Name="PwdLastSet";Expression={[datetime]::FromFileTime($_.PwdLastSet)}},`
                    @{Name="LastLogonTimeStamp";Expression={[datetime]::FromFileTime($_.LastLogonTimeStamp)}},WhenChanged,IPv4Address, `
                    @{Name="Stale";Expression={if((($_.pwdLastSet -lt $compareDate.ToFileTimeUTC()) -and ($_.pwdLastSet -ne 0)`
                    -and ($_.LastLogonTimeStamp -lt $compareDate.ToFileTimeUTC()) -and ($_.LastLogonTimeStamp -ne 0)) `
                    -and (!($_.serviceprincipalname -like "*MSClusterVirtualServer*"))){$True}else{$False}}}, `
                    @{Name="ParentOU";Expression={$_.distinguishedname.Substring($_.samaccountname.Length + 3)}}
                    $results = ($stalePCReport | Where-Object {$_.Stale -eq "true"}).count
                    $stalePCReport | Export-Csv -Append $StaleComputer_Report -NoTypeInformation
                    return $results
                }
                catch {
                    Write-Warning "[CATCH] Error, command (Get-staleADComputer) failed: $($_.Exception)"
                    Write-Warning $error[0].Exception.GetType().FullName
                }
            }
        }
        function Disable-staleADComputer {
            [CmdletBinding()]
            Param (
                [int]$disableAge,
                [string]$searchOU,
                [string]$Disabled_SC_Report,
                [string[]]$exclusionList
            )
            PROCESS {
                # Local Variables
                $compareDate = ([DateTime]::Today.AddDays(-($disableAge))).ToFileTimeUTC()
                $disabledComputers = @()
                try {
                    Get-ADComputer -Filter {(pwdLastSet -lt $compareDate) -and (LastLogonTimeStamp -lt $compareDate) -and (Enabled -eq $True) -and (IsCriticalSystemObject -ne $True)}`
                    -Properties Name,PwdLastSet,WhenChanged,SamAccountName,LastLogonTimeStamp,Enabled,Description,IPv4Address,`
                    operatingsystem,operatingsystemversion,serviceprincipalname,DistinguishedName -SearchScope Subtree -SearchBase $SearchOU |
                    Where-Object {((ForEach-Object {$_.DistinguishedName -notmatch ($exclusionList -join "|")}) -eq 'true') -and (!($_.serviceprincipalname -like "*MSClusterVirtualServer*"))} |
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
                        $disabledComputers | Export-Csv -Append $Disabled_SC_Report -NoTypeInformation
                        return $results
                    }
                }
                catch {
                    Write-Warning "[CATCH] Error, command (Disable-staleADComputer) failed: $($_.Exception)"
                    Write-Warning $error[0].Exception.GetType().FullName
                }
            }
        }
        function Remove-disabledSADComputer {
            [CmdletBinding()]
            Param (
                [int]$deleteAge,
                [string]$searchOU,
                [string]$Deleted_DSC_Report,
                [string[]]$exclusionList
            )
            PROCESS {
                # Local Variables
                $compareDate = ([DateTime]::Today.AddDays(-($deleteAge))).ToFileTimeUTC()
                $deletedComputers = @()
                try {
                    Get-ADComputer -Filter {(pwdLastSet -lt $compareDate) -and (LastLogonTimeStamp -lt $compareDate) -and (Enabled -eq $false) -and (IsCriticalSystemObject -ne $True)} `
                    -Properties Name,pwdLastSet,operatingsystem,LastLogonTimeStamp,distinguishedname,servicePrincipalName,samaccountname -SearchScope Subtree -SearchBase $SearchOU |
                    Where-Object {((ForEach-Object {$_.DistinguishedName -notmatch ($exclusionList -join "|")}) -eq 'true') -and (!($_.serviceprincipalname -like "*MSClusterVirtualServer*"))} |
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
                        $deletedComputers | Export-Csv -Append $Deleted_DSC_Report -NoTypeInformation
                        return $results
                    }

                }
                catch {
                    Write-Warning "[CATCH] Error, command (Remove-disabledSADComputer) failed: $($_.Exception)"
                    Write-Warning $error[0].Exception.GetType().FullName
                }
            }
        }
        function Show-Menu {
            Param (
                [int]$disableAge,
                [int]$deleteAge
            )
            $Today = Get-Date
            $compareDateDisable = [DateTime]::Today.AddDays(-($disableAge))
            $compareDateDelete = [DateTime]::Today.AddDays(-($deleteAge))
            Clear-Host
            Write-Host "`n================ AD Cleanup Menu v1.1 ================`n"
            Write-Host "Todays date: "$Today.ToShortDateString()            
            Write-Host "Disable cutoff date: "$compareDateDisable.ToShortDateString()           
            Write-Host "Delete cutoff date: "$compareDateDelete.ToShortDateString()             
            Write-Host "`n======================================================`n"            
            Write-Host "1: Press '1' To run report (Get-staleADComputer)."
            Write-Host "2: Press '2' To disable stale computers (Disable-staleADComputer)."
            Write-Host "3: Press '3' To delete disabled computers (Remove-disabledSADComputer)."
            Write-Host "Q: Press 'Q' to quit."
        }

    }
    PROCESS {

        #####################################
        # Script
        #####################################

        # Menu logic
        do {
            Show-Menu $disableAge $deleteAge
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
                    # Runs Get-staleADComputer report
                    Clear-Host
                    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Runs [Get-staleADComputer] to create COS AD stale computers report."
                    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Report does not run, returns to main menu."
                    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
                    $result = $host.ui.PromptForChoice("Run AD Stale Computers Report", "Do you want to run the AD stale computers report?", $options, 0)
                    switch ($result) {
                        0 {
                            Write-Output "`nYou selected Yes, running report..."
                            $staleCount = Get-staleADComputer $disableAge $searchOU $StaleComputer_Report $exclusionList
                            if ($staleCount -gt 0) {
                                Write-Output "Report successfully ran and exported to $StaleComputer_Report`nThere are [$staleCount] stale computers reported in [$searchOU] OU and sub OU's`n"
                            }
                            else {
                                Write-Output "Report successfully ran and exported to $StaleComputer_Report`nNo stale computers found"
                            }
                        }
                        1 {
                            Write-Output "`nYou selected No, returning to the main menu.`n"
                        }
                    }
                } 
                2 {
                    # Runs Disable-staleADComputer function to disable inactive computers
                    Clear-Host
                    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Runs [Disable-staleADComputer] to disable computers inactive for $disableAge days."
                    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","No computers disabled, returns to main menu."
                    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
                    $result = $host.ui.PromptForChoice("Disable All Stale Computers", "Do you want to disable all stale computers in COS AD?", $options, 1)
                    switch ($result) {
                        0 {
                            Write-Output "`nYou selected Yes, disabling computers..."
                            $disableCount = Disable-staleADComputer $disableAge $searchOU $Disabled_SC_Report $exclusionList
                            if ($disableCount -gt 0) {
                                Write-Output "Disabled computers log successfully exported to $Disabled_SC_Report`nDisabled [$disableCount] computers in [$searchOU] OU and sub OU's`n"
                            }
                            else {
                                Write-Output "No computers to be disabled"
                            }
                        }
                        1 {
                            Write-Output "`nYou selected No, returning to the main menu.`n"
                        }
                    }
                } 
                3 {
                    # Runs Remove-disabledSADComputer function to delete disabled inactive computers
                    Clear-Host
                    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Runs [Remove-disabledSADComputer] to delete computers inactive for $deleteAge days."
                    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","No computers deleted, returns to main menu."
                    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
                    $result = $host.ui.PromptForChoice("Delete All Stale Disabled Computers", "Do you want to delete all stale disabled computers in COS AD?", $options, 1)
                    switch ($result) {
                        0 {
                            Write-Output "`nYou selected Yes, deleting computers..."
                            $deleteCount = Remove-disabledSADComputer $deleteAge $searchOU $Deleted_DSC_Report $exclusionList
                            if ($deleteCount) {
                                Write-Output "Deleted computers log successfully exported to $Deleted_DSC_Report`nDeleted [$deleteCount] computers in [$searchOU] OU and sub OU's`n"
                            }
                            else {
                                Write-Output "No computers to be deleted"
                            }

                        }
                        1 {
                            Write-Output "`nYou selected No, returning to the main menu.`n"
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
Invoke-ADCleanup