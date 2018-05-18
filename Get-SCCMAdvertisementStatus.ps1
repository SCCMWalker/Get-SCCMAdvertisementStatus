#Triggers a machine policy update
<#
([wmiclass]'ROOT\ccm:SMS_Client').TriggerSchedule('{00000000-0000-0000-0000-000000000021}')
#>
$ServerLog = '\\global_server\public\Windows_10_Migration\WallOfWonder'
$LocalLog = 'C:\temp\SoftwareInstallStatus'
$Date = (Get-Date -Format yyyyMMddTHHmmss)
$CompareDate = Get-Date
$SoftwareCount = 0
#Checks for Computer CSV
$ListAllCSVS = Get-childItem "$ServerLog\Computers" -Filter '*.csv'
if ("$($Env:COMPUTERNAME).csv" -in $ListAllCSVS.Name) {
    #First Check for Currently Running Items

    if (!(Get-WmiObject -name "root\CCM\SoftMgmtAgent" -Class "ccm_executionrequestex")) {
        if (!(Test-Path "$LocalLog\")) {
            New-Item "$LocalLog\" -ItemType Directory
        }
        #Application WMI Query
        $Applications = get-wmiobject -ClassName "CCM_Application" -namespace "ROOT\ccm\ClientSDK" 

        #Package WMI Query and compares it to the Execution history, installation status is not stored directly in this WMI.
        $Packages = get-wmiobject -Class "CCM_SoftwareDistribution" -namespace "ROOT\ccm\Policy\Machine" 
        $ExecutionHistory = (Get-ChildItem -Path "HKLM:SOFTWARE\Microsoft\SMS\Mobile Client\Software Distribution\Execution History" -Recurse) | ForEach-Object { Get-ItemProperty $_.pspath}
        $PackageResults = $null
        $PackageResults = $Packages | ForEach-Object {
            If ($ExecutionHistory."_ProgramID" -contains $_.'PRG_ProgramID') {
                $Hit = $null
                $Hit = $ExecutionHistory |Where-Object -property "_ProgramID" -Contains $_.'PRG_ProgramID'
                $_ | Add-Member -MemberType NoteProperty -Name '_ProgramID' -Value $Hit.'_ProgramID' -force
                $_ | Add-Member -MemberType NoteProperty -Name '_RunStartTime' -Value $Hit.'_RunStartTime' -force
                $_ | Add-Member -MemberType NoteProperty -Name '_State' -Value $Hit.'_State' -force
                $_ | Add-Member -MemberType NoteProperty -Name 'SuccessOrFailureCode' -Value $Hit.SuccessOrFailureCode -force
                $_ | Add-Member -MemberType NoteProperty -Name 'SuccessOrFailureReason' -Value $Hit.SuccessOrFailureReason -force
                $_
            }
        } 

        #Pulls only the data we wish to log
        $PackageResults = $PackageResults | Where-Object -Property _State -ne "Success" | Select-Object -Property PKG_Name, PKG_PackageID, _State, SuccessOrFailureCode, SuccessOrFailureReason
        $SecondPackages = get-wmiobject -Class "CCM_SoftwareBase" -namespace "ROOT\ccm\ClientSDK" | Where-Object {($_.FullName -ne $null) -and ($_.InstallState -NE "Installed") -and ($_.ResolvedState -eq "Installed")}| Select-Object -Property FullName, InstallState, ResolvedState, ErrorCode


        #Task Sequences Queries
        $TaskSequences = $Packages | Where-Object -Property "__Class" -eq 'CCM_TaskSequence' 

        #Goes through each TS and compares to Applications/Packages installation status as the data isn't held directly in the WMI.
        $TaskSequenceResults = ForEach ($TaskSequence in $TaskSequences) {
            $TaskSequence.TS_References | ForEach-Object { 
            
                #Doing for loop to add new members to the TaskSequence Variable
                $Members = "SoftwareName,SoftwareVersion,PackageID,Type,RunStartTime,InstallState,ErrorCode,ConfigureState,ResolvedState,SuccessOrFailureReason" -split ','
                $Members | ForEach-Object {
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name $_ -Value '' -Force
                }

                #Regex Match, 2 = Program Name  3 = Application ID (called name below)
                $_ | Where-Object {$_ -match 'PackageID=(".*").*ProgramName=(".*")|ApplicationName=(".*")'} | Out-Null
            
            
                if ($Matches[2]) {

                    #Does Comparison
                    $Hit = $null
                    $Hit = $ExecutionHistory |Where-Object -property "_ProgramID" -Contains $Matches[2].Replace('"', '')
                    #Logs information
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "SoftwareName" -Value $Hit.'_ProgramID' -force
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "Type" -Value 'Package' -force
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "PackageID" -Value $Matches[1] -force
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "RunStartTime" -Value $Hit.'_RunStartTime' -force
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "InstallState" -Value $Hit.'_State' -force
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "ErrorCode" -Value $Hit.SuccessOrFailureCode -force
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "SuccessOrFailureReason" -Value $Hit.SuccessOrFailureReason -force
                    $TaskSequence | Select-Object -property *
                    #| Select-Object SoftwareName, Type, PackageID, RunStartTime, InstallState, ErrorCode, ConfigureState, ResolvedState, SuccessOrFailureReason
                
                }
                ElseIf ($Matches[3]) {
                    $Hit = $null
                    $Hit = $Applications | Where-Object {$_.ID -eq $Matches[3].Replace('"', '')} | Select-Object FullName, ErrorCode, ConfigureState, InstallState, LastInstallTime, ResolvedState, SoftwareVersion
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "Type" -Value 'Application' -force
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "SoftwareName" -Value $Hit.FullName -force
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "SoftwareVersion" -Value $Hit.SoftwareVersion -force
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "RunStartTime" -Value $Hit.LastInstallTime -force
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "InstallState" -Value $Hit.InstallState -force
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "ErrorCode" -Value $Hit.ErrorCode -force
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "ResolvedState" -Value $Hit.ResolvedState -force
                    $TaskSequence | Add-Member -MemberType NoteProperty -Name "ConfigureState" -Value $Hit.ConfigureState -force
                    $TaskSequence | Select-Object -property *
                    #$TaskSequence | Select-Object SoftwareName, Type, PackageID, RunStartTime, InstallState, ErrorCode, ConfigureState, ResolvedState, SuccessOrFailureReason
                }
                Else {
                    $TaskSequence | Select-Object -property *
                }
                Clear-Variable Matches
            }
        }
        #$TaskSequenceResults | Select-Object -Property PKG_Name, PKG_PackageID, SoftwareName
        #$TaskSequenceResults | Select-Object -Property PKG_Name, PKG_PackageID, SoftwareName, SoftwareVersion, PackageID, Type, RunStartTime, InstallState, ErrorCode, ConfigureState, ResolvedState, SuccessOrFailureReason
    
        #Pulls out data that shows an item is not successfully installed, this might need further tweaking. 
        #It appears that the Application state might be 'unknown' or 'available' in TS even though you would expect it to be required.
        $FinalTaskSequenceResults = $TaskSequenceResults | Where-Object {($_.Type -eq 'Package') -and ($_.InstallState -NE "Success")} | Select-Object -Property PKG_Name, PKG_PackageID, SoftwareName, SoftwareVersion, PackageID, Type, RunStartTime, InstallState, ErrorCode, ConfigureState, ResolvedState, SuccessOrFailureReason
        $FinalTaskSequenceResults += $TaskSequenceResults | Where-Object {($_.Type -eq 'Application') -and ($_.InstallState -NE "Installed") -and ($_.ResolvedState -eq "Installed")} | Select-Object -Property PKG_Name, PKG_PackageID, SoftwareName, SoftwareVersion, PackageID, Type, RunStartTime, InstallState, ErrorCode, ConfigureState, ResolvedState, SuccessOrFailureReason
    
        #Software Updates
        $Updates = get-wmiobject -Class "CCM_SoftwareUpdate" -namespace "ROOT\ccm\ClientSDK" | Where-Object -Property ErrorCode -NE "0"| Select-Object -Property  Name, Publisher, ArticleID, Evaluationstate, ErrorCode 

        #Limiting Applications (required for TaskSequence Previously)
        $Applications = $Applications | Where-Object {($_.FullName -ne $null) -and ($_.InstallState -NE "Installed") -and ($_.ResolvedState -eq "Installed")} | Select-Object -Property FullName, InstallState, ResolvedState, ErrorCode
    
        #Output
        #<#
        if (($Applications) -or ($PackageResults) -or ($SecondPackages) -or ($Updates) -or ($FinalTaskSequenceResults)) {
            <#
        
        #
        #Currently Commended out. This is for Local Logs
        #

        if ($Applications) {
            $Applications | Export-csv "$LocalLog\Applications.CSV" -NoTypeInformation -Force
        }
        if ($PackageResults) {
            $PackageResults  | Export-csv "$LocalLog\PackageResults.CSV" -NoTypeInformation -Force
        }
        if ($SecondPackages) {
            $SecondPackages | Export-csv "$LocalLog\SecondPackages.CSV" -NoTypeInformation -Force
        }
        if ($Updates) {
            $Updates | Export-csv "$LocalLog\Updates.CSV" -NoTypeInformation -Force
        }
        if ($FinalTaskSequenceResults) {
            $FinalTaskSequenceResults | Export-csv "$LocalLog\TaskSequence.CSV" -NoTypeInformation -Force
        }
        #>
            #Count The Number of pending/failed software.
            $SoftwareCount = (($Applications.Count) + (($PackageResults | Measure-Object).Count) + ($SecondPackages.Count) + ($FinalTaskSequenceResults.Count) + ($Updates.Count))
            #This show's Non-Compliance
            Write-Host 1
        }
        #>
        #Formats the output for the Server side logging.
        $Output = ([PSCustomObject]@{
                Computer          = $Env:COMPUTERNAME
                Imaged            = (([WMI]'').ConvertToDateTime((Get-WmiObject Win32_OperatingSystem).InstallDate)).tostring('yyyy.MM.dd-HH:mm:ss')
                TimesRan          = '0'
                TimeLastRan       = (Get-Date -Format yyyy-MM-ddTHH:mm:ss)
                ApplicationsLeft  = $Applications.Count
                PackagesLeft      = (($PackageResults | Measure-Object).Count + $SecondPackages.Count)
                TaskSequencesLeft = if ($FinalTaskSequenceResults.Count) {$FinalTaskSequenceResults.Count}Else {0}
                UpdatesLeft       = if ($Updates.Count) {$Updates.Count}Else {0}
                TotalLeft         = $SoftwareCount
                PeakTotal         = $SoftwareCount
                PeakReached       = ''
                ADGroupCount      = ''
            })
    }
    if ($Output.PeakTotal -ne 0) {
        $Output.PeakReached = $Output.TimeLastRan
    }

    #Searches DB file for any previous occurence of this computer running the script, captures the # of times ran and increments by one.
    #Searches DB for any computers that haven't ran script in 3 days and removes that.
    if (Get-Item "$ServerLog\Computers\$Env:COMPUTERNAME.CSV" -ErrorAction SilentlyContinue) {
        $CSV = Import-CSV "$ServerLog\Computers\$Env:COMPUTERNAME.CSV"
    }
    Else {
        New-Item "$ServerLog\Computers\$Env:COMPUTERNAME.CSV" -ItemType File | Out-Null
        $CSV = Import-CSV "$ServerLog\Computers\$Env:COMPUTERNAME.CSV"
    }
    Get-Acl "$ServerLog\Computers\Results.DB" | Set-Acl "$ServerLog\Computers\$Env:COMPUTERNAME.CSV"
    if ($CSV) {
        $Output.TimesRan = [int32]$CSV.TimesRan + 1
        if ($CSV.PeakTotal -ge $Output.PeakTotal) {
            $Output.PeakTotal = $CSV.PeakTotal
            $Output.PeakReached = $CSV.PeakReached
            $Output.ADGroupCount = $CSV.ADGroupCount
        }
    }

    #Removes Historical Results Data
    $ListAllResults = Get-childItem $ServerLog\ -Filter '*.csv'
    if ($ListAllResults) {
        $ListAllResults | ForEach-Object {   
            If ($_.CreationTime.AddDays('1') -lt $CompareDate) {
                $_ | Remove-Item -Force -ErrorAction SilentlyContinue | Out-Null
            }
        }
    }

    #Removes Historical Computer Data
    $ListAllCSVS = Get-childItem "$ServerLog\Computers" -Filter '*.csv'
    if ($ListAllCSVS) {
        $ListAllCSVS | ForEach-Object {   
            If ($_.LastWriteTime.AddDays('2') -lt $CompareDate) {
                $_ | Remove-Item -Force -ErrorAction SilentlyContinue | Out-Null
            }
        }
    }

    Clear-Variable ListAllCSVS, ListAllResults

    #Exports information to the DB
    #Creates a new CSV if last CSV is over 2 hours old, this is arbitrary.
    #The shorter the time the more CSV's that will spawn.

    Try {
        $Output| Export-Csv "$ServerLog\Computers\$Env:COMPUTERNAME.CSV" -NoTypeInformation -Force
        Get-Acl "$ServerLog\Computers\Results.DB" | Set-Acl "$ServerLog\Computers\$Env:COMPUTERNAME.CSV"
        If ((Get-Item $ServerLog\Computers\Results.db).LastWriteTime.AddMinutes('15') -lt $CompareDate) {
            $ListAllCSVS = Get-childItem "$ServerLog\Computers" -Filter '*.csv'
            $BigResults = $ListAllCSVS | ForEach-Object {Import-CSV $_.FullName}
            $BigResults | Export-CSV "$ServerLog\Computers\Results.DB" -NoTypeInformation
        }
        $ListAllResults = Get-childItem $ServerLog\ -Filter '*.csv'
        If (($ListAllResults | Sort-Object -Property CreationTime -Descending)[0].CreationTime.AddHours('2') -lt $CompareDate) {
            Copy-Item "$ServerLog\Computers\Results.DB" -Destination "$ServerLog\Results.$Date.csv"
            Get-Acl "$ServerLog\Computers\Results.DB" | Set-Acl "$ServerLog\Results.$Date.csv"
        }
    }
    #If the DB can't be written too or somehow the new CSV (shouldn't exist yet) is locked creates the following items
    #This shouldn't trigger unless someone has locked the DB.
    Catch {
        #New-Item "$ServerLog\CLOSETHEDBFILE.$Date.null" -ItemType file -Force | Out-Null
    }
}
<#
get-wmiobject -Class "AI_InstalledSoftwareCache" -namespace "ROOT\ccm\invagt"
get-wmiobject -Class "SMS_InstalledSoftware" -namespace "root\CIMV2\sms" | Out-File C:\temp\PackageList.txt
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SMS\Mobile Client\Software Distribution\Execution History
#>
#Remove-Variable *