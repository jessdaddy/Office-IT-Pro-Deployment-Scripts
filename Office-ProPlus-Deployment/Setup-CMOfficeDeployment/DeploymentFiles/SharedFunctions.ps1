try {
$enumDef = "
using System;
       [FlagsAttribute]
       public enum OfficeChannel
       {
          FirstReleaseCurrent = 0,
          Current = 1,
          FirstReleaseDeferred = 2,
          Deferred = 3
       }
"
Add-Type -TypeDefinition $enumDef -ErrorAction SilentlyContinue
} catch { }

try {
$enumBitness = "
using System;
       [FlagsAttribute]
       public enum Bitness
       {
          Both = 0,
          v32 = 1,
          v64 = 2
       }
"
Add-Type -TypeDefinition $enumBitness -ErrorAction SilentlyContinue
} catch { }

Function Write-Log {
 
    PARAM
	(
         [String]$Message,
         [String]$Path = $Global:UpdateAnywhereLogPath,
         [String]$LogName = $Global:UpdateAnywhereLogFileName,
         [int]$severity,
         [string]$component
	)
 
    try {
        $Path = $Global:UpdateAnywhereLogPath
        $LogName = $Global:UpdateAnywhereLogFileName
        if([String]::IsNullOrWhiteSpace($Path)){
            # Get Windows Folder Path
            $windowsDirectory = [Environment]::GetFolderPath("Windows")

            # Build log folder
            $Path = "$windowsDirectory\CCM\logs"
        }

        if([String]::IsNullOrWhiteSpace($LogName)){
             # Set log file name
            $LogName = "Office365UpdateAnywhere.log"
        }
        # Build log path
        $LogFilePath = Join-Path $Path $LogName

        # Create log file
        If (!($(Test-Path $LogFilePath -PathType Leaf)))
        {
            $null = New-Item -Path $LogFilePath -ItemType File -ErrorAction SilentlyContinue
        }

	    $TimeZoneBias = Get-WmiObject -Query "Select Bias from Win32_TimeZone"
        $Date= Get-Date -Format "HH:mm:ss.fff"
        $Date2= Get-Date -Format "MM-dd-yyyy"
        $type=1
 
        if ($LogFilePath) {
           "<![LOG[$Message]LOG]!><time=$([char]34)$date$($TimeZoneBias.bias)$([char]34) date=$([char]34)$date2$([char]34) component=$([char]34)$component$([char]34) context=$([char]34)$([char]34) type=$([char]34)$severity$([char]34) thread=$([char]34)$([char]34) file=$([char]34)$([char]34)>"| Out-File -FilePath $LogFilePath -Append -NoClobber -Encoding default
        }
    } catch {

    }
}

Function Set-Reg {
	PARAM
	(
        [String]$hive,
        [String]$keyPath,
	    [String]$valueName,
	    [String]$value,
        [String]$Type
    )

    Try
    {
        $null = New-ItemProperty -Path "$($hive):\$($keyPath)" -Name "$($valueName)" -Value "$($value)" -PropertyType $Type -Force -ErrorAction Stop
    }
    Catch
    {
        Write-Log -Message $_.Exception.Message -severity 3 -component $LogFileName
    }
}

Function StartProcess {
	Param
	(
		[String]$execFilePath,
        [String]$execParams
	)

    Try
    {
        $execStatement = [System.Diagnostics.Process]::Start( $execFilePath, $execParams ) 
        $execStatement.WaitForExit()
    }
    Catch
    {
        Write-Log -Message $_.Exception.Message -severity 1 -component "Office 365 Update Anywhere"
    }
}

function Test-ItemPathUNC() {    [CmdletBinding()]	
    Param
	(	    [Parameter(Mandatory=$true)]
	    [String]$Path,	    [Parameter()]
	    [String]$FileName    )    Process {       $drvLetter = FindAvailable       $Network = New-Object -ComObject "Wscript.Network"       try {           if (!($drvLetter.EndsWith(":"))) {               $drvLetter += ":"           }           $Network.MapNetworkDrive($drvLetter, $Path)            #New-PSDrive -Name $drvLetter -PSProvider FileSystem -Root $Path -ErrorAction Stop | Out-Null           if ($FileName) {             $target = $drvLetter + "\" + $FileName           } else {             $target = $drvLetter + "\"            }           $result = Test-Path -Path $target            return $result       } catch {         return $false       } finally {         #Remove-PSDrive $drvLetter -ErrorAction SilentlyContinue         $Network.RemoveNetworkDrive($drvLetter)       }    }}

function Copy-ItemUNC() {    [CmdletBinding()]	
    Param
	(	    [Parameter(Mandatory=$true)]
	    [String]$SourcePath,	    [Parameter(Mandatory=$true)]
	    [String]$TargetPath,	    [Parameter(Mandatory=$true)]
	    [String]$FileName    )    Process {       $drvLetter = FindAvailable       $Network = New-Object -ComObject "Wscript.Network"       try {           if (!($drvLetter.EndsWith(":"))) {               $drvLetter += ":"           }           $target = $drvLetter + "\"           $Network.MapNetworkDrive($drvLetter, $TargetPath)                                 #New-PSDrive -Name $drvLetter -PSProvider FileSystem -Root $TargetPath | Out-Null           Copy-Item -Path $SourcePath -Destination $target -Force       } finally {         #Remove-PSDrive $drvLetter         $Network.RemoveNetworkDrive($drvLetter)       }    }}

function FindAvailable() {
   #$drives = Get-PSDrive | select Name
   $drives = Get-WmiObject -Class Win32_LogicalDisk | select DeviceID

   for($n=90;$n -gt 68;$n--) {
      $letter= [char]$n + ":"
      $exists = $drives | where { $_.DeviceID -eq $letter }
      if ($exists) {
        if ($exists.Count -eq 0) {
            return $letter
        }
      } else {
        return $letter
      }
   }
   return $null
}

Function Get-OfficeVersion {
<#
.Synopsis
Gets the Office Version installed on the computer

.DESCRIPTION
This function will query the local or a remote computer and return the information about Office Products installed on the computer

.NOTES   
Name: Get-OfficeVersion
Version: 1.0.4
DateCreated: 2015-07-01
DateUpdated: 2015-08-28

.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts

.PARAMETER ComputerName
The computer or list of computers from which to query 

.PARAMETER ShowAllInstalledProducts
Will expand the output to include all installed Office products

.EXAMPLE
Get-OfficeVersion

Description:
Will return the locally installed Office product

.EXAMPLE
Get-OfficeVersion -ComputerName client01,client02

Description:
Will return the installed Office product on the remote computers

.EXAMPLE
Get-OfficeVersion | select *

Description:
Will return the locally installed Office product with all of the available properties

#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    [switch]$ShowAllInstalledProducts,
    [System.Management.Automation.PSCredential]$Credentials
)

begin {
    $HKLM = [UInt32] "0x80000002"
    $HKCR = [UInt32] "0x80000000"

    $excelKeyPath = "Excel\DefaultIcon"
    $wordKeyPath = "Word\DefaultIcon"
   
    $installKeys = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                   'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'

    $officeKeys = 'SOFTWARE\Microsoft\Office',
                  'SOFTWARE\Wow6432Node\Microsoft\Office'

    $defaultDisplaySet = 'DisplayName','Version', 'ComputerName'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
}

process {

 $results = new-object PSObject[] 0;

 foreach ($computer in $ComputerName) {
    if ($Credentials) {
       $os=Get-WMIObject win32_operatingsystem -computername $computer -Credential $Credentials
    } else {
       $os=Get-WMIObject win32_operatingsystem -computername $computer
    }

    $osArchitecture = $os.OSArchitecture

    if ($Credentials) {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer -Credential $Credentials
    } else {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer
    }

    [System.Collections.ArrayList]$VersionList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$PathList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$PackageList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$ClickToRunPathList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$ConfigItemList = New-Object -TypeName  System.Collections.ArrayList
    $ClickToRunList = new-object PSObject[] 0;

    foreach ($regKey in $officeKeys) {
       $officeVersion = $regProv.EnumKey($HKLM, $regKey)
       foreach ($key in $officeVersion.sNames) {
          if ($key -match "\d{2}\.\d") {
            if (!$VersionList.Contains($key)) {
              $AddItem = $VersionList.Add($key)
            }

            $path = join-path $regKey $key

            $configPath = join-path $path "Common\Config"
            $configItems = $regProv.EnumKey($HKLM, $configPath)
            if ($configItems) {
               foreach ($configId in $configItems.sNames) {
                 if ($configId) {
                    $Add = $ConfigItemList.Add($configId.ToUpper())
                 }
               }
            }

            $cltr = New-Object -TypeName PSObject
            $cltr | Add-Member -MemberType NoteProperty -Name InstallPath -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name UpdatesEnabled -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name UpdateUrl -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name StreamingFinished -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name Platform -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name ClientCulture -Value ""
            
            $packagePath = join-path $path "Common\InstalledPackages"
            $clickToRunPath = join-path $path "ClickToRun\Configuration"
            $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue

            [string]$officeLangResourcePath = join-path  $path "Common\LanguageResources"
            $mainLangId = $regProv.GetDWORDValue($HKLM, $officeLangResourcePath, "SKULanguage").uValue
            if ($mainLangId) {
                $mainlangCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $mainLangId}
                if ($mainlangCulture) {
                    $cltr.ClientCulture = $mainlangCulture.Name
                }
            }

            [string]$officeLangPath = join-path  $path "Common\LanguageResources\InstalledUIs"
            $langValues = $regProv.EnumValues($HKLM, $officeLangPath);
            if ($langValues) {
               foreach ($langValue in $langValues) {
                  $langCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $langValue}
               } 
            }

            if ($virtualInstallPath) {

            } else {
              $clickToRunPath = join-path $regKey "ClickToRun\Configuration"
              $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue
            }

            if ($virtualInstallPath) {
               if (!$ClickToRunPathList.Contains($virtualInstallPath.ToUpper())) {
                  $AddItem = $ClickToRunPathList.Add($virtualInstallPath.ToUpper())
               }

               $cltr.InstallPath = $virtualInstallPath
               $cltr.StreamingFinished = $regProv.GetStringValue($HKLM, $clickToRunPath, "StreamingFinished").sValue
               $cltr.UpdatesEnabled = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdatesEnabled").sValue
               $cltr.UpdateUrl = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdateUrl").sValue
               $cltr.Platform = $regProv.GetStringValue($HKLM, $clickToRunPath, "Platform").sValue
               $cltr.ClientCulture = $regProv.GetStringValue($HKLM, $clickToRunPath, "ClientCulture").sValue
               $ClickToRunList += $cltr
            }

            $packageItems = $regProv.EnumKey($HKLM, $packagePath)
            $officeItems = $regProv.EnumKey($HKLM, $path)

            foreach ($itemKey in $officeItems.sNames) {
              $itemPath = join-path $path $itemKey
              $installRootPath = join-path $itemPath "InstallRoot"

              $filePath = $regProv.GetStringValue($HKLM, $installRootPath, "Path").sValue
              if (!$PathList.Contains($filePath)) {
                  $AddItem = $PathList.Add($filePath)
              }
            }

            foreach ($packageGuid in $packageItems.sNames) {
              $packageItemPath = join-path $packagePath $packageGuid
              $packageName = $regProv.GetStringValue($HKLM, $packageItemPath, "").sValue
            
              if (!$PackageList.Contains($packageName)) {
                if ($packageName) {
                   $AddItem = $PackageList.Add($packageName.Replace(' ', '').ToLower())
                }
              }
            }

          }
       }
    }

    

    foreach ($regKey in $installKeys) {
        $keyList = new-object System.Collections.ArrayList
        $keys = $regProv.EnumKey($HKLM, $regKey)

        foreach ($key in $keys.sNames) {
           $path = join-path $regKey $key
           $installPath = $regProv.GetStringValue($HKLM, $path, "InstallLocation").sValue
           if (!($installPath)) { continue }
           if ($installPath.Length -eq 0) { continue }

           $buildType = "64-Bit"
           if ($osArchitecture -eq "32-bit") {
              $buildType = "32-Bit"
           }

           if ($regKey.ToUpper().Contains("Wow6432Node".ToUpper())) {
              $buildType = "32-Bit"
           }

           if ($key -match "{.{8}-.{4}-.{4}-1000-0000000FF1CE}") {
              $buildType = "64-Bit" 
           }

           if ($key -match "{.{8}-.{4}-.{4}-0000-0000000FF1CE}") {
              $buildType = "32-Bit" 
           }

           if ($modifyPath) {
               if ($modifyPath.ToLower().Contains("platform=x86")) {
                  $buildType = "32-Bit"
               }

               if ($modifyPath.ToLower().Contains("platform=x64")) {
                  $buildType = "64-Bit"
               }
           }

           $primaryOfficeProduct = $false
           $officeProduct = $false
           foreach ($officeInstallPath in $PathList) {
             if ($officeInstallPath) {
                $installReg = "^" + $installPath.Replace('\', '\\')
                $installReg = $installReg.Replace('(', '\(')
                $installReg = $installReg.Replace(')', '\)')
                if ($officeInstallPath -match $installReg) { $officeProduct = $true }
             }
           }

           if (!$officeProduct) { continue };
           
           $name = $regProv.GetStringValue($HKLM, $path, "DisplayName").sValue          

           if ($ConfigItemList.Contains($key.ToUpper()) -and $name.ToUpper().Contains("MICROSOFT OFFICE")) {
              $primaryOfficeProduct = $true
           }

           $version = $regProv.GetStringValue($HKLM, $path, "DisplayVersion").sValue
           $modifyPath = $regProv.GetStringValue($HKLM, $path, "ModifyPath").sValue 

           $cltrUpdatedEnabled = $NULL
           $cltrUpdateUrl = $NULL
           $clientCulture = $NULL;

           [string]$clickToRun = $false
           if ($ClickToRunPathList.Contains($installPath.ToUpper())) {
               $clickToRun = $true
               if ($name.ToUpper().Contains("MICROSOFT OFFICE")) {
                  $primaryOfficeProduct = $true
               }

               foreach ($cltr in $ClickToRunList) {
                 if ($cltr.InstallPath) {
                   if ($cltr.InstallPath.ToUpper() -eq $installPath.ToUpper()) {
                       $cltrUpdatedEnabled = $cltr.UpdatesEnabled
                       $cltrUpdateUrl = $cltr.UpdateUrl
                       if ($cltr.Platform -eq 'x64') {
                           $buildType = "64-Bit" 
                       }
                       if ($cltr.Platform -eq 'x86') {
                           $buildType = "32-Bit" 
                       }
                       $clientCulture = $cltr.ClientCulture
                   }
                 }
               }
           }
           
           if (!$primaryOfficeProduct) {
              if (!$ShowAllInstalledProducts) {
                  continue
              }
           }

           $object = New-Object PSObject -Property @{DisplayName = $name; Version = $version; InstallPath = $installPath; ClickToRun = $clickToRun; 
                     Bitness=$buildType; ComputerName=$computer; ClickToRunUpdatesEnabled=$cltrUpdatedEnabled; ClickToRunUpdateUrl=$cltrUpdateUrl;
                     ClientCulture=$clientCulture }
           $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
           $results += $object

        }
    }

  }

  $results = Get-Unique -InputObject $results 

  return $results;
}

}

Function Get-InstalledLanguages() {
    [CmdletBinding()]
    Param(
    )
    process {
       $returnLangs = @()
       $mainRegPath = Get-OfficeCTRRegPath

       $activeConfig = Get-ItemProperty -Path "hklm:\$mainRegPath\ProductReleaseIDs"
       $activeId = $activeConfig.ActiveConfiguration
       $languages = Get-ChildItem -Path "hklm:\$mainRegPath\ProductReleaseIDs\$activeId\culture"

       foreach ($language in $languages) {
          $lang = Get-ItemProperty -Path  $language.pspath
          $keyName = $lang.PSChildName
          if ($keyName.Contains(".")) {
              $keyName = $keyName.Split(".")[0]
          }

          if ($keyName.ToLower() -ne "x-none") {
             $culture = New-Object system.globalization.cultureinfo($keyName)
             $returnLangs += $culture
          }
       }

       return $returnLangs
    }
}

Function Get-OfficeCDNUrl() {
    $CDNBaseUrl = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name CDNBaseUrl -ErrorAction SilentlyContinue).CDNBaseUrl
    if (!($CDNBaseUrl)) {
       $CDNBaseUrl = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration -Name CDNBaseUrl -ErrorAction SilentlyContinue).CDNBaseUrl
    }
    if (!($CDNBaseUrl)) {
        Push-Location
        $path15 = 'HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\Active\stream'
        $path16 = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\Active\stream'
        if (Test-Path -Path $path16) { Set-Location $path16 }
        if (Test-Path -Path $path15) { Set-Location $path15 }

        $items = Get-Item . | Select-Object -ExpandProperty property
        $properties = $items | ForEach-Object {
           New-Object psobject -Property @{"property"=$_; "Value" = (Get-ItemProperty -Path . -Name $_).$_}
        }

        $value = $properties | Select Value
        $firstItem = $value[0]
        [string] $cdnPath = $firstItem.Value

        $CDNBaseUrl = Select-String -InputObject $cdnPath -Pattern "http://officecdn.microsoft.com/.*/.{8}-.{4}-.{4}-.{4}-.{12}" -AllMatches | % { $_.Matches } | % { $_.Value }
        Pop-Location
    }
    return $CDNBaseUrl
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

function Test-URL {
   param( 
      [string]$url = $NULL
   )

   [bool]$validUrl = $false
   try {
     $req = [System.Net.HttpWebRequest]::Create($url);
     $res = $req.GetResponse()

     if($res.StatusCode -eq "OK") {
        $validUrl = $true
     }
     $res.Close(); 
   } catch {
      Write-Host "Invalid UpdateSource. File Not Found: $url" -ForegroundColor Red
      $validUrl = $false
      throw;
   }

   return $validUrl
}

function Change-UpdatePathToChannel {
   [CmdletBinding()]
   param( 
     [Parameter()]
     [string] $UpdatePath,

     [Parameter()]
     [bool] $ValidateUpdateSourceFiles = $true,

     [Parameter()]
     [string] $Channel = $null
   )

   $newUpdatePath = $UpdatePath
   $newUpdateLong = $UpdatePath

   if ($Channel) {
      $detectedChannel = $Channel
   } else {
      $detectedChannel = Detect-Channel
   }

   $branchName = $detectedChannel.branch

   $branchShortName = "DC"
   if ($branchName.ToLower() -eq "current") {
      $branchShortName = "CC"
   }
   if ($branchName.ToLower() -eq "firstreleasecurrent") {
      $branchShortName = "FRCC"
   }
   if ($branchName.ToLower() -eq "firstreleasedeferred") {
      $branchShortName = "FRDC"
   }
   if ($branchName.ToLower() -eq "deferred") {
      $branchShortName = "DC"
   }

   $channelNames = @("FRCC", "CC", "FRDC", "DC")
   $channelLongNames = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred", "Business", "FirstReleaseBusiness")

   $madeChange = $false
   foreach ($channelName in $channelNames) {
      if ($UpdatePath.ToUpper().EndsWith("\$channelName")) {
         $newUpdatePath = $newUpdatePath -replace "\\$channelName", "\$branchShortName"
         $newUpdateLong = $newUpdateLong -replace "\\$channelName", "\$branchName"
         $madeChange = $true
      } 
      if ($UpdatePath.ToUpper().Contains("\$channelName\")) {
         $newUpdatePath = $newUpdatePath -replace "\\$channelName\\", "\$branchShortName\"
         $newUpdateLong = $newUpdateLong -replace "\\$channelName\\", "\$branchName\"
         $madeChange = $true
      } 
      if ($UpdatePath.ToUpper().EndsWith("/$channelName")) {
         $newUpdatePath = $newUpdatePath -replace "\/$channelName", "/$branchShortName"
         $newUpdateLong = $newUpdateLong -replace "\\$channelName\\", "\$branchName\"
         $madeChange = $true
      }
      if ($UpdatePath.ToUpper().Contains("/$channelName/")) {
         $newUpdatePath = $newUpdatePath -replace "\/$channelName\/", "/$branchShortName/"
         $newUpdateLong = $newUpdateLong -replace "\/$channelName\/", "/$branchName/"
         $madeChange = $true
      }
   }

   foreach ($channelName in $channelLongNames) {
      $channelName = $channelName.ToString().ToUpper()
      if ($UpdatePath.ToUpper().EndsWith("\$channelName")) {
         $newUpdatePath = $newUpdatePath -replace "\\$channelName", "\$branchShortName"
         $newUpdateLong = $newUpdateLong -replace "\\$channelName", "\$branchName"
         $madeChange = $true
      } 
      if ($UpdatePath.ToUpper().Contains("\$channelName\")) {
         $newUpdatePath = $newUpdatePath -replace "\\$channelName\\", "\$branchShortName\"
         $newUpdateLong = $newUpdateLong -replace "\\$channelName\\", "\$branchName\"
         $madeChange = $true
      } 
      if ($UpdatePath.ToUpper().EndsWith("/$channelName")) {
         $newUpdatePath = $newUpdatePath -replace "\/$channelName", "/$branchShortName"
         $newUpdateLong = $newUpdateLong -replace "\\$channelName\\", "\$branchName\"
         $madeChange = $true
      }
      if ($UpdatePath.ToUpper().Contains("/$channelName/")) {
         $newUpdatePath = $newUpdatePath -replace "\/$channelName\/", "/$branchShortName/"
         $newUpdateLong = $newUpdateLong -replace "\/$channelName\/", "/$branchName/"
         $madeChange = $true
      }
   }

   if (!($madeChange)) {
      if ($newUpdatePath.Contains("/")) {
         if ($newUpdatePath.EndsWith("/")) {
           $newUpdatePath += "$branchShortName"
         } else {
           $newUpdatePath += "/$branchShortName"
         }
      }
      if ($newUpdatePath.Contains("\")) {
         if ($newUpdatePath.EndsWith("\")) {
           $newUpdatePath += "$branchShortName"
         } else {
           $newUpdatePath += "\$branchShortName"
         }
      }
   }

   if (!($madeChange)) {
      if ($newUpdateLong.Contains("/")) {
         if ($newUpdateLong.EndsWith("/")) {
           $newUpdateLong += "$branchName"
         } else {
           $newUpdateLong += "/$branchName"
         }
      }
      if ($newUpdateLong.Contains("\")) {
         if ($newUpdateLong.EndsWith("\")) {
           $newUpdateLong += "$branchName"
         } else {
           $newUpdateLong += "\$branchName"
         }
      }
   }

   try {
     $pathAlive = Test-UpdateSource -UpdateSource $newUpdatePath -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles
     if (!($pathAlive)) {
        $pathAlive = Test-UpdateSource -UpdateSource $newUpdateLong -ValidateUpdateSourceFiles $ValidateUpdateSourceFiles
        if ($pathAlive) {
           $newUpdatePath = $newUpdateLong
        }
     }
   } catch {
     $pathAlive = $false
   }
   
   if ($pathAlive) {
     return $newUpdatePath
   } else {
     return $UpdatePath
   }
}

Function Test-UpdateSource() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $UpdateSource = $NULL,

        [Parameter()]
        [bool] $ValidateUpdateSourceFiles = $true
    )

  	$uri = [System.Uri]$UpdateSource

    [bool]$sourceIsAlive = $false

    if($uri.Host){
	    $sourceIsAlive = Test-Connection -Count 1 -computername $uri.Host -Quiet
    }else{
        $sourceIsAlive = Test-Path $uri.LocalPath -ErrorAction SilentlyContinue
    }

    if ($ValidateUpdateSourceFiles) {
       if ($sourceIsAlive) {
           $sourceIsAlive = Validate-UpdateSource -UpdateSource $UpdateSource
       }
    }

    return $sourceIsAlive
}

Function Validate-UpdateSource() {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string] $UpdateSource = $NULL
    )

    [bool]$validUpdateSource = $true
    [string]$cabPath = ""

    if ($UpdateSource) {
        $mainRegPath = Get-OfficeCTRRegPath
        $configRegPath = $mainRegPath + "\Configuration"
        $currentplatform = (Get-ItemProperty HKLM:\$configRegPath -Name Platform -ErrorAction SilentlyContinue).Platform
        $updateToVersion = (Get-ItemProperty HKLM:\$configRegPath -Name UpdateToVersion -ErrorAction SilentlyContinue).UpdateToVersion
        $llcc = (Get-ItemProperty HKLM:\$configRegPath -Name ClientCulture -ErrorAction SilentlyContinue).ClientCulture

        $mainCab = "$UpdateSource\Office\Data\v32.cab"
        $bitness = "32"
        if ($currentplatform -eq "x64") {
            $mainCab = "$UpdateSource\Office\Data\v64.cab"
            $bitness = "64"
        }

        if (!($updateToVersion)) {
           $cabXml = Get-CabVersion -FilePath $mainCab
           $updateToVersion = $cabXml.Version.Available.Build
        }

        [xml]$xml = Get-ChannelXml -Bitness $bitness
        $languages = Get-InstalledLanguages

        $checkFiles = $xml.UpdateFiles.File | Where {   $_.language -eq "0" }
        foreach ($language in $languages) {
           $checkFiles += $xml.UpdateFiles.File | Where { $_.language -eq $language.LCID}
        }

        foreach ($checkFile in $checkFiles) {
           $fileName = $checkFile.name -replace "%version%", $updateToVersion
           $relativePath = $checkFile.relativePath -replace "%version%", $updateToVersion

           $fullPath = "$UpdateSource$relativePath$fileName"
           if ($fullPath.ToLower().StartsWith("http")) {
              $fullPath = $fullPath -replace "\\", "/"
           } else {
              $fullPath = $fullPath -replace "/", "\"
           }
           
           $updateFileExists = $false
           if ($fullPath.ToLower().StartsWith("http")) {
               $updateFileExists = Test-URL -url $fullPath
           } else {
               $updateFileExists = Test-Path -Path $fullPath
           }

           if (!($updateFileExists)) {
              $fileExists = $missingFiles.Contains($fullPath)
              if (!($fileExists)) {
                 $missingFiles.Add($fullPath)
                 Write-Host "Source File Missing: $fullPath"
                 Write-Log -Message "Source File Missing: $fullPath" -severity 1 -component "Office 365 Update Anywhere" 
              }     
              $validUpdateSource = $false
           }
        }

    }
    
    return $validUpdateSource
}

function Detect-Channel {
   param( 

   )

   Process {
      $currentBaseUrl = Get-OfficeCDNUrl
      $channelXml = Get-ChannelXml

      $currentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.URL -eq $currentBaseUrl -and $_.branch -notcontains 'Business' }
      return $currentChannel
   }

}

function Get-CabVersion {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [string] $FilePath = $NULL
   )

   process {
       $cabPath = $FilePath
       $fileName = Split-Path -Path $cabPath -Leaf

       if ($cabPath.ToLower().StartsWith("http")) {
           $webclient = New-Object System.Net.WebClient
           $XMLFilePath = "$env:TEMP/$fileName"
           $XMLDownloadURL= $FilePath
           $webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)
       } else {
         $XMLFilePath = $cabPath
       }

       $tmpName = "VersionDescriptor.xml"
       expand $XMLFilePath $env:TEMP -f:$tmpName | Out-Null
       $tmpName = $env:TEMP + "\VersionDescriptor.xml"
       [xml]$versionXml = Get-Content $tmpName

       return $versionXml
   }
}

function Get-ChannelXml {
   [CmdletBinding()]
   param( 
      [Parameter()]
      [string] $Bitness = "32"
   )

   process {
       $cabPath = "$PSScriptRoot\ofl.cab"

       if (!(Test-Path -Path $cabPath)) {
           $webclient = New-Object System.Net.WebClient
           $XMLFilePath = "$env:TEMP/ofl.cab"
           $XMLDownloadURL = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"
           $webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)
       } else {
           $XMLFilePath = $cabPath
       }

       $tmpName = "o365client_" + $Bitness + "bit.xml"
       expand $XMLFilePath $env:TEMP -f:$tmpName | Out-Null
       $tmpName = $env:TEMP + "\o365client_" + $Bitness + "bit.xml"
       [xml]$channelXml = Get-Content $tmpName

       return $channelXml
   }

}

Function GetScriptRoot() {
 process {
     [string]$scriptPath = "."

     if ($PSScriptRoot) {
       $scriptPath = $PSScriptRoot
     } else {
       $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
     }

     return $scriptPath
 }
}