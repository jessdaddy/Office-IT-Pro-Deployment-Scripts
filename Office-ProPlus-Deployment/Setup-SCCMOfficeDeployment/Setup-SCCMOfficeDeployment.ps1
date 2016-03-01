function Setup-SCCMOfficeProPlusPackage {
<#
.SYNOPSIS
Automates the configuration of System Center Configuration Manager (SCCM) to configure Office Click-To-Run Updates

.DESCRIPTION

.PARAMETER path
The UNC Path where the downloaded bits will be stored for installation to the target machines.

.PARAMETER $SiteCode
The 3 Letter Site ID.

.PARAMETER SCCMPSModulePath
Allows the user to specify that full path to the ConfigurationManager.psd1 PowerShell Module. This is especially useful if SCCM is installed in a non standard path.

.PARAMETER distributionPoint
Sets which distribution points will be used, and distributes the package.

.Example
Setup-SCCMOfficeUpdates -Path \\SCCM-CM\OfficeDeployment -PackageName "Office ProPlus Deployment" -ProgramName "Office2016Setup.exe" -distributionPoint SCCM-CM.CONTOSO.COM
#>

[CmdletBinding(SupportsShouldProcess=$true)]
Param
(
	[Parameter()]
	[String]$Path = $null,

	[Parameter()]
	[String]$Version,

	[Parameter()]
	[String]$SiteCode = $null,
	
	[Parameter()]
	[String]$PackageName = "Office ProPlus Deployment",
		
	[Parameter()]
	[String]$ProgramName = "Office2016Setup.exe",

	[Parameter()]	
	[Bool]$UpdateOnlyChangedBits = $false,

	[Parameter()]
	[String[]] $RequiredPlatformNames = @("All x86 Windows 7 Client", "All x86 Windows 8 Client", "All x86 Windows 8.1 Client", "All Windows 10 Professional/Enterprise and higher (32-bit) Client","All x64 Windows 7 Client", "All x64 Windows 8 Client", "All x64 Windows 8.1 Client", "All Windows 10 Professional/Enterprise and higher (64-bit) Client"),
	
	[Parameter()]
	[string]$distributionPoint,

	[Parameter()]
	[uint16]$DeploymentExpiryDurationInDays = 15,

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
    Write-Host
    Write-Host 'Configuring System Center Configuration Manager to Deploy Office ProPlus' -BackgroundColor DarkBlue
    Write-Host

    if (!$Path) {
         $Path = CreateOfficeUpdateShare
    }

    Set-Location $PSScriptRoot
	Set-Location $startLocation
    Set-Location $PSScriptRoot

    Write-Host "Loading SCCM Module"
    Write-Host ""

    #HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SMS\Setup

    $sccmModulePath = GetSCCMPSModulePath -SCCMPSModulePath $SCCMPSModulePath 
    
    if ($sccmModulePath) {
        Import-Module $sccmModulePath

        if (!$SiteCode) {
           $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
        }

	    Set-Location "$SiteCode`:"	

        $package = CreateSCCMPackage -Name $PackageName -Path $path -UpdateOnlyChangedBits $UpdateOnlyChangedBits

        CreateSCCMProgram -Name $programName -PackageName $PackageName -Path $path -RequiredPlatformNames $requiredPlatformNames

        Write-Host "Starting Content Distribution"	

        if ($distributionPoint) {
	        Start-CMContentDistribution -PackageName $PackageName -DistributionPointName $distributionPoint
        }

        Write-Host 
        Write-Host "NOTE: In order to deploy the package you must run the function 'Deploy-SCCMOfficeUpdates'." -BackgroundColor Red
        Write-Host "      You should wait until the content has finished distributing to the distribution points." -BackgroundColor Red
        Write-Host "      otherwise the deployments will fail. The clients will continue to fail until the " -BackgroundColor Red
        Write-Host "      content distribution is complete." -BackgroundColor Red

    } else {
        throw [System.IO.FileNotFoundException] "Could Not find file ConfigurationManager.psd1"
    }
}
End
{
    Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
    Set-Location $startLocation    
}
}

function Deploy-SCCMOfficeProPlusPackage {
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

		[Parameter()]
		[String]$PackageName = "Office ProPlus Deployment",

		[Parameter()]
		[String]$ProgramName = "Office2016Setup.exe",

		[Parameter()]	
		[Bool]$UpdateOnlyChangedBits = $true,

		[Parameter()]
		[String]$SCCMPSModulePath = $NULL
	) 
Begin
{

}
Process
{
    $sccmModulePath = GetSCCMPSModulePath -SCCMPSModulePath $SCCMPSModulePath 
    
    if ($sccmModulePath) {
        Import-Module $sccmModulePath

        if (!$SiteCode) {
            $SiteCode = (Get-ItemProperty -Path "hklm:\SOFTWARE\Microsoft\SMS\Identification" -Name "Site Code").'Site Code'
        }

	    Set-Location "$SiteCode`:"	

        $package = Get-CMPackage -Name $packageName

        $packageDeploy = Get-CMDeployment | where {$_.PackageId  -eq $package.PackageId }
        if ($packageDeploy.Count -eq 0) {
            Write-Host "Creating Package Deployment for: $packageName"

            $dtNow = [datetime]::Now
            $dtNow = $dtNow.AddDays(-1)
            $start = Get-Date -Year $dtNow.Year -Month $dtNow.Month -Day $dtNow.Day -Hour 12 -Minute 0

            $schedule = New-CMSchedule -Start $start -RecurInterval Days -RecurCount 7

     	    Start-CMPackageDeployment -CollectionName $Collection -PackageName $PackageName -ProgramName $ProgramName -StandardProgram  -DeployPurpose Required `
                                      -RerunBehavior AlwaysRerunProgram -ScheduleEvent AsSoonAsPossible `
                                      -Schedule $schedule

        } else {
            Write-Host "Package Deployment Already Exists for: $packageName"
        }
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

	Set-CMPackage -Name $packageName -Priority Normal -EnableBinaryDeltaReplication $UpdateOnlyChangedBits -CopyToPackageShareOnDistributionPoint $True

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
		[String]$Path, 

		[Parameter()]
		[String]$Name = "Office2016Setup.exe",
		
		[Parameter()]
		[String[]] $RequiredPlatformNames = @()

	) 

    $program = Get-CMProgram -PackageName $PackageName -ProgramName $Name

    $commandLine = "Office2016Setup.exe /configure Configuration_UpdateSource.xml"

    Write-Host "`tProgram: $Name"

    if($program -eq $null -or !$program)
    {
        Write-Host "`t`tCreating Program..."	        
	    $program = New-CMProgram -PackageName $PackageName -StandardProgramName $Name -DriveMode RenameWithUnc -CommandLine $commandLine -ProgramRunType OnlyWhenUserIsLoggedOn -RunMode RunWithUserRights -UserInteraction $false -RunType Normal
    } else {
        Write-Host "`t`tAlready Exists"
    }
	
    

    Write-Host ""
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
Function Get-Site([string[]]$computerName = $env:COMPUTERNAME) {
    Get-WmiObject -ComputerName $ComputerName -Namespace "root\SMS" -Class "SMS_ProviderLocation" | foreach-object{ 
        if ($_.ProviderForLocalSite -eq $true){$SiteCode=$_.sitecode} 
    } 
    if ($SiteCode -eq "") { 
        throw ("Sitecode of ConfigMgr Site at " + $ComputerName + " could not be determined.") 
    } else { 
        Return $SiteCode 
    } 
}
