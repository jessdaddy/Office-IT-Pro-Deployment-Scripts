Add-Type -ErrorAction SilentlyContinue -TypeDefinition @"
   public enum Channel
   {
      Current,
      Deferred,
      FirstReleaseCurrent,
      FirstReleaseDeferred
   }
"@


Function Change-OfficeChannel {
<#
.Synopsis
Uses OfficeC2RClient.exe tool to change version and branch of installed office
.DESCRIPTION
This function will upgrade or downgrade the version of Office to the version specified, or to the
most recent version of office found in the update folder in the Channel specified.  This function
also simultaneously changes the channel of Office installed on the client.
.NOTES   
Name: Change-OfficeChannel
DateUpdated: 2016-03-15
.PARAMETER Channel
The Channel associated with Office Updates ie Current, Deferred, etc.
.PARAMETER Version
The version of Office to install.  The install files must be located in the Channel update path
.EXAMPLE
Change-OfficeChannel -Channel Current
Description:
Switches the Channel of the Office install to Current, updates the current Office install to the most up recent update available in the
UpdateURL path.
.EXAMPLE
Change-OfficeChannel -Channel Current -Version 16.0.6001.1068
Description:
Switches the Channel of the Office install to Current, updates the current Office install to the version specified, as long as that version
is available in the UpdateURL path.
#>
param(
    [Parameter(Mandatory=$true)]
    [Channel]$Channel,
    
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string]$Version
)
begin {
    $UpdateURLKey = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'  #UpdateURL
    $Office2RClientKey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' #ClientFolder
    
}
process {
        #get local path
        $scriptPath = "."

        if ($PSScriptRoot) {
            $scriptPath = $PSScriptRoot
        } else {
            $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
        }

        #set the UpdateURL path so it can be updated in the registry
        $UpdateURLPath = $scriptPath + '\'+$Channel


        # update reg key here for updateurl
        if(Test-Path $UpdateURLPath){
            New-ItemProperty $Office2RClientKey -Name UpdateUrl -PropertyType String -Value $UpdateURLPath -Force | Out-Null
            Write-Host "Updating RegKey `"UpdateURL`" in path `""$Office2RClientKey"`" to value `""$UpdateURLPath"`""
        }

    
    #find update exe file
    $OfficeUpdatePath = Get-ItemProperty -Path $Office2RClientKey | Select-Object -Property ClientFolder
    $temp = Out-String -InputObject $OfficeUpdatePath
    $temp = $temp.Substring($temp.LastIndexOf('-')+2)
    $temp = $temp.Trim()
    $OfficeUpdatePath = $temp
    $OfficeUpdatePath+= '\OfficeC2RClient.exe'

    
    #get latest version available in branch
    if(!$Version){
    [array]$totalVersion = @()

    $LatestBranchVersionPath = $UpdateURLPath + '\Office\Data'
    if(Test-Path $LatestBranchVersionPath){
    $DirectoryList = Get-ChildItem $LatestBranchVersionPath
    Foreach($listItem in $DirectoryList){
        if($listItem.GetType().Name -eq 'DirectoryInfo'){
            $totalVersion+=$listItem.Name
        }
    }
    }

    $totalVersion = $totalVersion | Sort-Object


    #sets version number to the newest version in directory for channel if version is not set by user in argument  
    if($totalVersion.Count -gt 0){
    $Version = $totalVersion[0]
    }

    }


    $arguments = "/update user displaylevel=false updatepromptuser=false"
    if($Version){
        $arguments+= ' updatetoversion='+$Version
    }    
    

    
    #run update exe file
    Start-Process -FilePath $OfficeUpdatePath -ArgumentList $arguments
     
     Wait-ForOfficeCTRUpadate

}

}









Function Wait-ForOfficeCTRUpadate() {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [int] $TimeOutInMinutes = 120
    )

    begin {
        $HKLM = [UInt32] "0x80000002"
        $HKCR = [UInt32] "0x80000000"
    }

    process {
       Write-Host "Waiting for Update process to Complete..."

       [datetime]$operationStart = Get-Date
       [datetime]$totalOperationStart = Get-Date

       Start-Sleep -Seconds 10

       $mainRegPath = Get-OfficeCTRRegPath
       $scenarioPath = $mainRegPath + "\scenario"

       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -ErrorAction Stop

       [DateTime]$startTime = Get-Date

       [string]$executingScenario = ""
       $failure = $false
       $cancelled = $false
       $updateRunning=$false
       [string[]]$trackProgress = @()
       [string[]]$trackComplete = @()
       [int]$noScenarioCount = 0

       do {
           $allComplete = $true
           $executingScenario = $regProv.GetStringValue($HKLM, $mainRegPath, "ExecutingScenario").sValue
           
           $scenarioKeys = $regProv.EnumKey($HKLM, $scenarioPath)
           foreach ($scenarioKey in $scenarioKeys.sNames) {
              if (!($executingScenario)) { continue }
              if ($scenarioKey.ToLower() -eq $executingScenario.ToLower()) {
                $taskKeyPath = Join-Path $scenarioPath "$scenarioKey\TasksState"
                $taskValues = $regProv.EnumValues($HKLM, $taskKeyPath).sNames

                foreach ($taskValue in $taskValues) {
                    [string]$status = $regProv.GetStringValue($HKLM, $taskKeyPath, $taskValue).sValue
                    $operation = $taskValue.Split(':')[0]
                    $keyValue = $taskValue
                   
                    if ($status.ToUpper() -eq "TASKSTATE_FAILED") {
                        $failure = $true
                    }

                    if ($status.ToUpper() -eq "TASKSTATE_CANCELLED") {
                        $cancelled = $true
                    }

                    if (($status.ToUpper() -eq "TASKSTATE_COMPLETED") -or`
                        ($status.ToUpper() -eq "TASKSTATE_CANCELLED") -or`
                        ($status.ToUpper() -eq "TASKSTATE_FAILED")) {
                        if (($trackProgress -contains $keyValue) -and !($trackComplete -contains $keyValue)) {
                            $displayValue = $operation + "`t" + $status + "`t" + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                            #Write-Host $displayValue
                            $trackComplete += $keyValue 

                            $statusName = $status.Split('_')[1];

                            if (($operation.ToUpper().IndexOf("DOWNLOAD") -gt -1) -or `
                                ($operation.ToUpper().IndexOf("APPLY") -gt -1)) {

                                $operationTime = getOperationTime -OperationStart $operationStart

                                $displayText = $statusName + "`t" + $operationTime

                                Write-Host $displayText
                            }
                        }
                    } else {
                        $allComplete = $false
                        $updateRunning=$true


                        if (!($trackProgress -contains $keyValue)) {
                             $trackProgress += $keyValue 
                             $displayValue = $operation + "`t" + $status + "`t" + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

                             $operationStart = Get-Date

                             if ($operation.ToUpper().IndexOf("DOWNLOAD") -gt -1) {
                                Write-Host "Downloading Update: " -NoNewline
                             }

                             if ($operation.ToUpper().IndexOf("APPLY") -gt -1) {
                                Write-Host "Applying Update: " -NoNewline
                             }

                             if ($operation.ToUpper().IndexOf("FINALIZE") -gt -1) {
                                Write-Host "Finalizing Update: " -NoNewline
                             }

                             #Write-Host $displayValue
                        }
                    }
                }
              }
           }

           if ($allComplete) {
              break;
           }

           if ($startTime -lt (Get-Date).AddHours(-$TimeOutInMinutes)) {
              throw "Waiting for Update Timed-Out"
              break;
           }

           Start-Sleep -Seconds 5
       } while($true -eq $true) 

       $operationTime = getOperationTime -OperationStart $operationStart

       $displayValue = ""
       if ($cancelled) {
         $displayValue = "CANCELLED`t" + $operationTime
       } else {
         if ($failure) {
            $displayValue = "FAILED`t" + $operationTime
         } else {
            $displayValue = "COMPLETED`t" + $operationTime
         }
       }

       Write-Host $displayValue

       $totalOperationTime = getOperationTime -OperationStart $totalOperationStart

       if ($updateRunning) {
          if ($failure) {
            Write-Host "Update Failed"
          } else {
            Write-Host "Update Completed - Total Time: $totalOperationTime"
          }
       } else {
          Write-Host "Update Not Running"
       } 
    }
}





Function Get-OfficeCTRRegPath() {
    $path15 = 'SOFTWARE\Microsoft\Office\15.0\ClickToRun'
    $path16 = 'SOFTWARE\Microsoft\Office\ClickToRun'
    if (Test-Path "HKLM:\$path16") {
        return $path16
    }
    else {
        if (Test-Path "HKLM:\$path15") {
            return $path15
        }
    }
}




Function getOperationTime() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [DateTime] $OperationStart
    )

    $operationTime = ""

    $dateDiff = NEW-TIMESPAN –Start $OperationStart –End (GET-DATE)
    $strHours = formatTimeItem -TimeItem $dateDiff.Hours.ToString() 
    $strMinutes = formatTimeItem -TimeItem $dateDiff.Minutes.ToString() 
    $strSeconds = formatTimeItem -TimeItem $dateDiff.Seconds.ToString() 

    if ($dateDiff.Days -gt 0) {
        $operationTime += "Days: " + $dateDiff.Days.ToString() + ":"  + $strHours + ":" + $strMinutes + ":" + $strSeconds
    }
    if ($dateDiff.Hours -gt 0 -and $dateDiff.Days -eq 0) {
        if ($operationTime.Length -gt 0) { $operationTime += " " }
        $operationTime += "Hours: " + $strHours + ":" + $strMinutes + ":" + $strSeconds
    }
    if ($dateDiff.Minutes -gt 0 -and $dateDiff.Days -eq 0 -and $dateDiff.Hours -eq 0) {
        if ($operationTime.Length -gt 0) { $operationTime += " " }
        $operationTime += "Minutes: " + $strMinutes + ":" + $strSeconds
    }
    if ($dateDiff.Seconds -gt 0 -and $dateDiff.Days -eq 0 -and $dateDiff.Hours -eq 0 -and $dateDiff.Minutes -eq 0) {
        if ($operationTime.Length -gt 0) { $operationTime += " " }
        $operationTime += "Seconds: " + $strSeconds
    }

    return $operationTime
}




Function formatTimeItem() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $TimeItem = ""
    )

    [string]$returnItem = $TimeItem
    if ($TimeItem.Length -eq 1) {
       $returnItem = "0" + $TimeItem
    }
    return $returnItem
}