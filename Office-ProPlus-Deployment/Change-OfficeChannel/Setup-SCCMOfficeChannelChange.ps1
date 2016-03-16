$enum = "
using System;
 
    [FlagsAttribute]
    public enum InstallType
    {
        ScriptInstall = 0,
        ConfigurationFileInstall = 1
    }
"
Add-Type -TypeDefinition $enum -ErrorAction SilentlyContinue

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

function Create-SCCMOfficeChannelPackages {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
	    [Parameter()]	
	    [Bool]$UpdateOnlyChangedBits = $false,

	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$SCCMPSModulePath = $NULL
    )
    Begin
    {
        $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }
    Process
    {
       . "$PSScriptRoot\Download-OfficeProPlusChannels.ps1"

       $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred")
       $ChannelXml = Get-ChannelXml

       foreach ($Channel in $ChannelList) {
           $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
           $latestVersion = Get-BranchLatestVersion -ChannelUrl $selectChannel.URL 
           $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel
           $versionExists = CheckIfVersionExists -Version $latestVersion -Channel $Channel
           $LargeDrv = Get-LargestDrive

           $Path = CreateOfficeChannelShare -Path "$LargeDrv\OfficeChannels"
           $packageName = "OfficeProPlus-$ChannelShortName-$latestVersion"

           $ChannelPath = "$Path\$packageName"
           $LocalPath = "$LargeDrv\OfficeChannels\$packageName"

           [System.IO.Directory]::CreateDirectory($LocalPath) | Out-Null
               
           Download-OfficeProPlusChannels -TargetDirectory $LocalPath -Channels $Channel -Version $latestVersion -UseChannelFolderShortName $true

           $OSSourcePath = "$PSScriptRoot\Change-OfficeChannel.ps1"
           $OCScriptPath = "$LocalPath\Change-OfficeChannel.ps1"
           if (!(Test-Path $OCScriptPath)) {
              Copy-Item -Path $OSSourcePath  -Destination $OCScriptPath -Force
           }

           if (!$versionExists) {
              LoadSCCMPrereqs -SiteCode $SiteCode -SCCMPSModulePath $SCCMPSModulePath

              $package = CreateSCCMPackage -Name $packageName -Path $ChannelPath -Channel $Channel -Version $latestVersion -UpdateOnlyChangedBits $UpdateOnlyChangedBits
              [string]$CommandLine = "powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -File .\Change-OfficeChannel -Channel $Channel"
              CreateSCCMProgram -Name $packageName -PackageName $packageName -CommandLine $CommandLine -RequiredPlatformNames $requiredPlatformNames
           } else {
             Write-Host "Package with Version already exists: $latestVersion"
           }

       }

    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Distribute-SCCMOfficeChannelPackages {
<#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (SCCM) to configure Office Click-To-Run Updates

.DESCRIPTION

.PARAMETER path
The UNC Path where the downloaded bits will be stored for installation to the target machines.

.PARAMETER Source
The UNC Path where the downloaded branch bits are stored. Required if source parameter is specified.

.PARAMETER Branch

The update branch to be used with the deployment. Current options are "Business, Current, FirstReleaseBusiness, FirstReleaseCurrent".

.PARAMETER $SiteCode
The 3 Letter Site ID.

.PARAMETER SCCMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if SCCM is installed in a non standard path.

.PARAMETER distributionPoint
Sets which distribution points will be used, and distributes the package.

.Example
Setup-SCCMOfficeProPlusPackage -Path \\SCCM-CM\OfficeDeployment -PackageName "Office ProPlus Deployment" -ProgramName "Office2016Setup.exe" -distributionPoint SCCM-CM.CONTOSO.COM -source \\SCCM-CM\updates -branch Current
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(0,1,2,3),

	    [Parameter()]
	    [string]$DistributionPoint,

	    [Parameter()]
	    [string]$DistributionPointGroupName,

	    [Parameter()]
	    [uint16]$DeploymentExpiryDurationInDays = 15,

	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$SCCMPSModulePath = $NULL

    )
    Begin
    {
        $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }
    Process
    {
        $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred")
        $ChannelXml = Get-ChannelXml

        foreach ($ChannelName in $ChannelList) {
           if ($Channels -contains $ChannelName) {
               $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $ChannelName.ToString() }
               $latestVersion = Get-BranchLatestVersion -ChannelUrl $selectChannel.URL 
               $ChannelShortName = ConvertChannelNameToShortName -ChannelName $ChannelName
               $versionExists = CheckIfVersionExists -Version $latestVersion -Channel $ChannelName

               $packageName = "OfficeProPlus-$ChannelShortName-$latestVersion"

               LoadSCCMPrereqs -SiteCode $SiteCode -SCCMPSModulePath $SCCMPSModulePath

               if ($versionExists) {
                    if ($DistributionPointGroupName) {
                        Write-Host "Starting Content Distribution for package: $packageName"
	                    Start-CMContentDistribution -PackageName $packageName -DistributionPointGroupName $DistributionPointGroupName
                    }

                    if ($DistributionPoint) {
                        Write-Host "Starting Content Distribution for package: $packageName"
                        Start-CMContentDistribution -PackageName $packageName -DistributionPointName $DistributionPoint
                    }
               }
           }
        }

        Write-Host 
        Write-Host "NOTE: In order to deploy the package you must run the function 'Deploy-SCCMOfficeChannelPackage'." -BackgroundColor Red
        Write-Host "      You should wait until the content has finished distributing to the distribution points." -BackgroundColor Red
        Write-Host "      otherwise the deployments will fail. The clients will continue to fail until the " -BackgroundColor Red
        Write-Host "      content distribution is complete." -BackgroundColor Red
    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Deploy-SCCMOfficeChannelPackage {
<#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (SCCM) to configure Office Click-To-Run Updates

.DESCRIPTION

.PARAMETER Collection
The target SCCM Collection

.PARAMETER PackageName
The Name of the SCCM package create by the Setup-SCCMOfficeProPlusPackage function

.PARAMETER ProgramName
The Name of the SCCM program create by the Setup-SCCMOfficeProPlusPackage function

.PARAMETER UpdateOnlyChangedBits
Determines whether or not the EnableBinaryDeltaReplication enabled or not

.PARAMETER SCCMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if SCCM is installed in a non standard path.

.Example
Deploy-SCCMOfficeProPlusPackage -Collection "CollectionName"
Deploys the Package created by the Setup-SCCMOfficeProPlusPackage function
#>
    [CmdletBinding()]	
    Param
	(
		[Parameter(Mandatory=$true)]
		[String]$Collection = "",

        [Parameter(Mandatory=$true)]
        [OfficeChannel] $Channel,

	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$SCCMPSModulePath = $NULL
	) 
    Begin
    {
        $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }
    Process
    {
        $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred")
        $ChannelXml = Get-ChannelXml

        foreach ($ChannelName in $ChannelList) {
          if ($Channel.ToString().ToLower() -eq $ChannelName.ToLower()) {
              $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $ChannelName.ToString() }
              $latestVersion = Get-BranchLatestVersion -ChannelUrl $selectChannel.URL 
              $ChannelShortName = ConvertChannelNameToShortName -ChannelName $ChannelName
              $versionExists = CheckIfVersionExists -Version $latestVersion -Channel $ChannelName

              LoadSCCMPrereqs -SiteCode $SiteCode -SCCMPSModulePath $SCCMPSModulePath

              $packageName = "OfficeProPlus-$ChannelShortName-$latestVersion"
              if ($versionExists) {

                  $package = Get-CMPackage -Name $packageName

                  $packageDeploy = Get-CMDeployment | where {$_.PackageId  -eq $package.PackageId }
                  if ($packageDeploy.Count -eq 0) {
                    try {
     	                Start-CMPackageDeployment -CollectionName "$Collection" -PackageName "$packageName" -ProgramName "$packageName" -StandardProgram  -DeployPurpose Required `
                                                -RerunBehavior AlwaysRerunProgram -ScheduleEvent AsSoonAsPossible -FastNetworkOption RunProgramFromDistributionPoint -SlowNetworkOption RunProgramFromDistributionPoint
                        Write-Host "Package Deployment created for: $packageName"
                    } catch {
                        [string]$ErrorMessage = $_.ErrorDetails 
                        if ($ErrorMessage.ToLower().Contains("Could not find property PackageID".ToLower())) {
                            Write-Host 
                            Write-Host "Package: $packageName"
                            Write-Host "The package has not finished deploying to the distribution points." -BackgroundColor Red
                            Write-Host "Please try this command against once the distribution points have been updated" -BackgroundColor Red
                        } else {
                            throw
                        }
                    }  
                  } else {
                    Write-Host "Package Deployment Already Exists for: $packageName"
                  }
              } else {
                 throw "Package does not exist: $packageName"
              }
          }
        }
    }
    End {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation 
    }
}


function CheckIfVersionExists() {
    [CmdletBinding()]	
    Param
	(
	   [Parameter(Mandatory=$True)]
	   [String]$Version,

		[Parameter()]
		[String]$Channel
    )
    Begin
    {
        $startLocation = Get-Location
    }
    Process {
       LoadSCCMPrereqs

       $VersionName = "$Channel - $Version"

       $existingPackage = Get-CMPackage | Where { $_.Version -eq $VersionName }
       if ($existingPackage) {
         return $true
       }

       return $false
    }
}

function LoadSCCMPrereqs() {
    [CmdletBinding()]	
    Param
	(
	    [Parameter()]
	    [String]$SiteCode = $null,

	    [Parameter()]
	    [String]$SCCMPSModulePath = $NULL
    )
    Begin
    {
        $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location
    }
    Process {

        $sccmModulePath = GetSCCMPSModulePath -SCCMPSModulePath $SCCMPSModulePath 
    
        if ($sccmModulePath) {
            Import-Module $sccmModulePath

            if (!$SiteCode) {
               $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
            }

            Set-Location "$SiteCode`:"	
        }
    }
}

function CreateSCCMPackage() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name = "Office ProPlus Deployment",
		
		[Parameter(Mandatory=$True)]
		[String]$Path,

		[Parameter()]
		[String]$Version,

		[Parameter()]
		[String]$Channel,

		[Parameter()]	
		[Bool]$UpdateOnlyChangedBits = $true
	) 

    Write-Host "`tPackage: $Name"

    $package = Get-CMPackage -Name $Name 

    if($package -eq $null -or !$package)
    {
        Write-Host "`t`tCreating Package: $Name"
        $package = New-CMPackage -Name $Name  -Path $path
    } else {
        Write-Host "`t`tAlready Exists"	
    }
		
    Write-Host "`t`tSetting Package Properties"

    $VersionName = "$Channel - $Version"

	Set-CMPackage -Name $Name -Priority Normal -EnableBinaryDeltaReplication $UpdateOnlyChangedBits `
                  -CopyToPackageShareOnDistributionPoint $True -Version $VersionName

    Write-Host ""

    $package = Get-CMPackage -Name $Name
    return $package
}

function CreateSCCMProgram() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$PackageName = "Office ProPlus Deployment",
		
		[Parameter(Mandatory=$True)]
		[String]$CommandLine, 

		[Parameter()]
		[String]$Name = "Office2016Setup.exe",
		
		[Parameter()]
		[String[]] $RequiredPlatformNames = @()

	) 

    $program = Get-CMProgram -PackageName $PackageName -ProgramName $Name

    Write-Host "`tProgram: $Name"

    if($program -eq $null -or !$program)
    {
        Write-Host "`t`tCreating Program..."	        
	    $program = New-CMProgram -PackageName $PackageName -StandardProgramName $Name -DriveMode RenameWithUnc -CommandLine $CommandLine -ProgramRunType OnlyWhenUserIsLoggedOn -RunMode RunWithAdministrativeRights -UserInteraction $false -RunType Normal
    } else {
        Write-Host "`t`tAlready Exists"
    }

    Write-Host ""
}

function CreateOfficeChannelShare() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name = "OfficeChannels$",
		
		[Parameter()]
		[String]$Path = "$env:SystemDrive\OfficeChannels"
	) 

    IF (!(TEST-PATH $Path)) { 
      $addFolder = New-Item $Path -type Directory 
    }
    
    $ACL = Get-ACL $Path

    $identity = New-Object System.Security.Principal.NTAccount  -argumentlist ("$env:UserDomain\$env:UserName") 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"FullControl","ContainerInherit, ObjectInherit","None","Allow")

    $addAcl = $ACL.AddAccessRule($accessRule) | Out-Null

    $identity = New-Object System.Security.Principal.NTAccount -argumentlist ("$env:UserDomain\Domain Admins") 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"FullControl","ContainerInherit, ObjectInherit","None","Allow")
    $addAcl = $ACL.AddAccessRule($accessRule) | Out-Null

    $identity = "Everyone"
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"Read","ContainerInherit, ObjectInherit","None","Allow")
    $addAcl = $ACL.AddAccessRule($accessRule) | Out-Null

    Set-ACL -Path $Path -ACLObject $ACL | Out-Null
    
    $share = Get-WmiObject -Class Win32_share | Where {$_.name -eq "$Name"}
    if (!$share) {
       Create-FileShare -Name $Name -Path $Path | Out-Null
    }

    $sharePath = "\\$env:COMPUTERNAME\$Name"
    return $sharePath
}


function CreateOfficeUpdateShare() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name = "OfficeDeployment$",
		
		[Parameter()]
		[String]$Path = "$env:SystemDrive\OfficeDeployment"
	) 

    IF (!(TEST-PATH $Path)) { 
      $addFolder = New-Item $Path -type Directory 
    }
    
    $ACL = Get-ACL $Path

    $identity = New-Object System.Security.Principal.NTAccount  -argumentlist ("$env:UserDomain\$env:UserName") 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"FullControl","ContainerInherit, ObjectInherit","None","Allow")

    $addAcl = $ACL.AddAccessRule($accessRule)

    $identity = New-Object System.Security.Principal.NTAccount -argumentlist ("$env:UserDomain\Domain Admins") 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"FullControl","ContainerInherit, ObjectInherit","None","Allow")
    $addAcl = $ACL.AddAccessRule($accessRule)

    $identity = "Everyone"
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"Read","ContainerInherit, ObjectInherit","None","Allow")
    $addAcl = $ACL.AddAccessRule($accessRule)

    Set-ACL -Path $Path -ACLObject $ACL
    
    $share = Get-WmiObject -Class Win32_share | Where {$_.name -eq "$Name"}
    if (!$share) {
       Create-FileShare -Name $Name -Path $Path
    }

    $sharePath = "\\$env:COMPUTERNAME\$Name"
    return $sharePath
}

function GetSupportedPlatforms([String[]] $requiredPlatformNames){
    $computerName = $env:COMPUTERNAME
    #$assignedSite = $([WmiClass]"\\$computerName\ROOT\ccm:SMS_Client").getassignedsite()
    $siteCode = Get-Site  
    $filteredPlatforms = Get-WmiObject -ComputerName $computerName -Class SMS_SupportedPlatforms -Namespace "root\sms\site_$siteCode" | Where-Object {$_.IsSupported -eq $true -and  $_.OSName -like 'Win NT' -and ($_.OSMinVersion -match "6\.[0-9]{1,2}\.[0-9]{1,4}\.[0-9]{1,4}" -or $_.OSMinVersion -match "10\.[0-9]{1,2}\.[0-9]{1,4}\.[0-9]{1,4}") -and ($_.OSPlatform -like 'I386' -or $_.OSPlatform -like 'x64')}

    $requiredPlatforms = $filteredPlatforms| Where-Object {$requiredPlatformNames.Contains($_.DisplayText) } #| Select DisplayText, OSMaxVersion, OSMinVersion, OSName, OSPlatform | Out-GridView

    $supportedPlatforms = @()

    foreach($p in $requiredPlatforms)
    {
        $osDetail = ([WmiClass]("\\$computerName\root\sms\site_$siteCode`:SMS_OS_Details")).CreateInstance()    
        $osDetail.MaxVersion = $p.OSMaxVersion
        $osDetail.MinVersion = $p.OSMinVersion
        $osDetail.Name = $p.OSName
        $osDetail.Platform = $p.OSPlatform

        $supportedPlatforms += $osDetail
    }

    $supportedPlatforms
}

function CreateDownloadXmlFile([string]$Path, [string]$ConfigFileName){
	#1 - Set the correct version number to update Source location
	$sourceFilePath = "$path\$configFileName"
    $localSourceFilePath = ".\$configFileName"

    Set-Location $PSScriptRoot

    if (Test-Path -Path $localSourceFilePath) {   
	  $doc = [Xml] (Get-Content $localSourceFilePath)

      $addNode = $doc.Configuration.Add
	  $addNode.OfficeClientEdition = $bitness

      $doc.Save($sourceFilePath)
    } else {
      $doc = New-Object System.XML.XMLDocument

      $configuration = $doc.CreateElement("Configuration");
      $a = $doc.AppendChild($configuration);

      $addNode = $doc.CreateElement("Add");
      $addNode.SetAttribute("OfficeClientEdition", $bitness)
      if ($Version) {
         if ($Version.Length -gt 0) {
             $addNode.SetAttribute("Version", $Version)
         }
      }
      $a = $doc.DocumentElement.AppendChild($addNode);

      $addProduct = $doc.CreateElement("Product");
      $addProduct.SetAttribute("ID", "O365ProPlusRetail")
      $a = $addNode.AppendChild($addProduct);

      $addLanguage = $doc.CreateElement("Language");
      $addLanguage.SetAttribute("ID", "en-us")
      $a = $addProduct.AppendChild($addLanguage);

	  $doc.Save($sourceFilePath)
    }
}

function CreateUpdateXmlFile([string]$Path, [string]$ConfigFileName, [string]$Bitness, [string]$Version){
    $newConfigFileName = $ConfigFileName -replace '\.xml'
    $newConfigFileName = $newConfigFileName + "$Bitness" + ".xml"

    Copy-Item -Path ".\$ConfigFileName" -Destination ".\$newConfigFileName"
    $ConfigFileName = $newConfigFileName

    $testGroupFilePath = "$path\$ConfigFileName"
    $localtestGroupFilePath = ".\$ConfigFileName"

	$testGroupConfigContent = [Xml] (Get-Content $localtestGroupFilePath)

	$addNode = $testGroupConfigContent.Configuration.Add
	$addNode.OfficeClientEdition = $bitness
    $addNode.SourcePath = $path	

	$updatesNode = $testGroupConfigContent.Configuration.Updates
	$updatesNode.UpdatePath = $path
	$updatesNode.TargetVersion = $version

	$testGroupConfigContent.Save($testGroupFilePath)
    return $ConfigFileName
}

function Create-FileShare() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name = "",
		
		[Parameter()]
		[String]$Path = ""
	)

    $description = "$name"

    $Method = "Create"
    $sd = ([WMIClass] "Win32_SecurityDescriptor").CreateInstance()

    #AccessMasks:
    #2032127 = Full Control
    #1245631 = Change
    #1179817 = Read

    $userName = "$env:USERDOMAIN\$env:USERNAME"

    #Share with the user
    $ACE = ([WMIClass] "Win32_ACE").CreateInstance()
    $Trustee = ([WMIClass] "Win32_Trustee").CreateInstance()
    $Trustee.Name = $userName
    $Trustee.Domain = $NULL
    #original example assigned this, but I found it worked better if I left it empty
    #$Trustee.SID = ([wmi]"win32_userAccount.Domain='york.edu',Name='$name'").sid   
    $ace.AccessMask = 2032127
    $ace.AceFlags = 3 #Should almost always be three. Really. don't change it.
    $ace.AceType = 0 # 0 = allow, 1 = deny
    $ACE.Trustee = $Trustee 
    $sd.DACL += $ACE.psObject.baseobject 

    #Share with Domain Admins
    $ACE = ([WMIClass] "Win32_ACE").CreateInstance()
    $Trustee = ([WMIClass] "Win32_Trustee").CreateInstance()
    $Trustee.Name = "Domain Admins"
    $Trustee.Domain = $Null
    #$Trustee.SID = ([wmi]"win32_userAccount.Domain='york.edu',Name='$name'").sid   
    $ace.AccessMask = 2032127
    $ace.AceFlags = 3
    $ace.AceType = 0
    $ACE.Trustee = $Trustee 
    $sd.DACL += $ACE.psObject.baseobject    
    
     #Share with the user
    $ACE = ([WMIClass] "Win32_ACE").CreateInstance()
    $Trustee = ([WMIClass] "Win32_Trustee").CreateInstance()
    $Trustee.Name = "Everyone"
    $Trustee.Domain = $Null
    #original example assigned this, but I found it worked better if I left it empty
    #$Trustee.SID = ([wmi]"win32_userAccount.Domain='york.edu',Name='$name'").sid   
    $ace.AccessMask = 1179817 
    $ace.AceFlags = 3 #Should almost always be three. Really. don't change it.
    $ace.AceType = 0 # 0 = allow, 1 = deny
    $ACE.Trustee = $Trustee 
    $sd.DACL += $ACE.psObject.baseobject    

    $mc = [WmiClass]"Win32_Share"
    $InParams = $mc.psbase.GetMethodParameters($Method)
    $InParams.Access = $sd
    $InParams.Description = $description
    $InParams.MaximumAllowed = $Null
    $InParams.Name = $name
    $InParams.Password = $Null
    $InParams.Path = $path
    $InParams.Type = [uint32]0

    $R = $mc.PSBase.InvokeMethod($Method, $InParams, $Null)
    switch ($($R.ReturnValue))
     {
          0 { break}
          2 {Write-Host "Share:$name Path:$path Result:Access Denied" -foregroundcolor red -backgroundcolor yellow;break}
          8 {Write-Host "Share:$name Path:$path Result:Unknown Failure" -foregroundcolor red -backgroundcolor yellow;break}
          9 {Write-Host "Share:$name Path:$path Result:Invalid Name" -foregroundcolor red -backgroundcolor yellow;break}
          10 {Write-Host "Share:$name Path:$path Result:Invalid Level" -foregroundcolor red -backgroundcolor yellow;break}
          21 {Write-Host "Share:$name Path:$path Result:Invalid Parameter" -foregroundcolor red -backgroundcolor yellow;break}
          22 {Write-Host "Share:$name Path:$path Result:Duplicate Share" -foregroundcolor red -backgroundcolor yellow;break}
          23 {Write-Host "Share:$name Path:$path Result:Reedirected Path" -foregroundcolor red -backgroundcolor yellow;break}
          24 {Write-Host "Share:$name Path:$path Result:Unknown Device or Directory" -foregroundcolor red -backgroundcolor yellow;break}
          25 {Write-Host "Share:$name Path:$path Result:Network Name Not Found" -foregroundcolor red -backgroundcolor yellow;break}
          default {Write-Host "Share:$name Path:$path Result:*** Unknown Error ***" -foregroundcolor red -backgroundcolor yellow;break}
     }
}

function GetSCCMPSModulePath() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$SCCMPSModulePath = $NULL
	)

    [bool]$pathExists = $false

    if ($SCCMPSModulePath) {
       if ($SCCMPSModulePath.ToLower().EndsWith(".psd1")) {
         $sccmModulePath = $SCCMPSModulePath
         $pathExists = Test-Path -Path $sccmModulePath
       }
    }

    if (!$pathExists) {
        $uiInstallDir = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Setup" -Name "UI Installation Directory").'UI Installation Directory'
        $sccmModulePath = Join-Path $uiInstallDir "bin\ConfigurationManager.psd1"

        $pathExists = Test-Path -Path $sccmModulePath
        if (!$pathExists) {
            $sccmModulePath = "$env:ProgramFiles\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
            $pathExists = Test-Path -Path $sccmModulePath
        }
    }

    if (!$pathExists) {
       $uiAdminPath = ${env:SMS_ADMIN_UI_PATH}
       if ($uiAdminPath.ToLower().EndsWith("\bin")) {
           $dirInfo = $uiAdminPath
       } else {
           $dirInfo = ([System.IO.DirectoryInfo]$uiAdminPath).Parent.FullName
       }
      
       $sccmModulePath = $dirInfo + "\ConfigurationManager.psd1"
       $pathExists = Test-Path -Path $sccmModulePath
    }

    if (!$pathExists) {
       $sccmModulePath = "${env:ProgramFiles(x86)}\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
       $pathExists = Test-Path -Path $sccmModulePath
    }

    if (!$pathExists) {
       $sccmModulePath = "${env:ProgramFiles(x86)}\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
       $pathExists = Test-Path -Path $sccmModulePath
    }

    if (!$pathExists) {
       throw "Cannot find the ConfigurationManager.psd1 file. Please use the -SCCMPSModulePath parameter to specify the location of the PowerShell Module"
    }

    return $sccmModulePath
}

# Specify one of SCCM servers and Site code is returned automatically 
function Get-Site([string[]]$computerName = $env:COMPUTERNAME) {
    Get-WmiObject -ComputerName $ComputerName -Namespace "root\SMS" -Class "SMS_ProviderLocation" | foreach-object{ 
        if ($_.ProviderForLocalSite -eq $true){$SiteCode=$_.sitecode} 
    } 
    if ($SiteCode -eq "") { 
        throw ("Sitecode of ConfigMgr Site at " + $ComputerName + " could not be determined.") 
    } else { 
        Return $SiteCode 
    } 
}

function DownloadBits() {
    [CmdletBinding()]	
    Param
	(
	    [Parameter()]
	    [OfficeBranch]$Branch = $null
	)

    $DownloadScript = "$PSScriptRoot\Download-OfficeProPlusBranch.ps1"
    if (Test-Path -Path $DownloadScript) {
       



    }
}

function Get-ChannelXml() {
   [CmdletBinding()]
   param( 
      
   )

   process {
       $cabPath = "$PSScriptRoot\ofl.cab"

       $webclient = New-Object System.Net.WebClient
       $XMLFilePath = "$env:TEMP/ofl.cab"
       $XMLDownloadURL = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"
       $webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)

       $tmpName = "o365client_64bit.xml"
       expand $XMLFilePath $env:TEMP -f:$tmpName | Out-Null
       $tmpName = $env:TEMP + "\o365client_64bit.xml"
       [xml]$channelXml = Get-Content $tmpName

       return $channelXml
   }

}

function Get-ChannelUrl() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [Channel]$Channel
   )

   Process {
      $channelXml = Get-ChannelXml

      $currentChannel = $channelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
      return $currentChannel
   }

}

function Get-BranchLatestVersion() {
   [CmdletBinding()]
   param( 
      [Parameter(Mandatory=$true)]
      [string]$ChannelUrl
   )

   process {
       $webclient = New-Object System.Net.WebClient
       $CABFilePath = "$env:TEMP/v64.cab"
       $XMLDownloadURL = "$ChannelUrl/Office/Data/v64.cab"
       $webclient.DownloadFile($XMLDownloadURL,$CABFilePath)

       $tmpName = "VersionDescriptor.xml"
       expand $CABFilePath $env:TEMP -f:$tmpName | Out-Null
       $tmpName = $env:TEMP + "\VersionDescriptor.xml"
       [xml]$versionXml = Get-Content $tmpName

       return $versionXml.Version.Available.Build
   }
}

function Get-LargestDrive() {
   [CmdletBinding()]
   param( 
   )
   process {
      $drives = Get-Partition | where {$_.DriveLetter}
      $driveInfoList = @()

      foreach ($drive in $drives) {
          $driveLetter = $drive.DriveLetter
          $deviceFilter = "DeviceID='" + $driveLetter + ":'" 
 
          $driveInfo = Get-WmiObject Win32_LogicalDisk -ComputerName "." -Filter $deviceFilter
          $driveInfoList += $driveInfo
      }

      $SortList = Sort-Object -InputObject $driveInfoList -Property FreeSpace

      $FreeSpaceDrive = $SortList[0]
      return $FreeSpaceDrive.DeviceID
   }
}

function ConvertChannelNameToShortName {
    Param(
       [Parameter()]
       [string] $ChannelName
    )
    Process {
       if ($ChannelName.ToLower() -eq "FirstReleaseCurrent".ToLower()) {
         return "FRCC"
       }
       if ($ChannelName.ToLower() -eq "Current".ToLower()) {
         return "CC"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseDeferred".ToLower()) {
         return "FRDC"
       }
       if ($ChannelName.ToLower() -eq "Deferred".ToLower()) {
         return "DC"
       }
       if ($ChannelName.ToLower() -eq "Business".ToLower()) {
         return "DC"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseBusiness".ToLower()) {
         return "FRDC"
       }
    }
}