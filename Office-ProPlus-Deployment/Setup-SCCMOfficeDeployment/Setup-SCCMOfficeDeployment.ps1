$enum = "
using System;
 
    [FlagsAttribute]
    public enum SCCMDeploymentType
    {
        DeployWithScript = 0,
        DeployWithConfigurationFile = 1
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

$enum2 = "
using System;
 
    [FlagsAttribute]
    public enum SCCMOfficeProgramType
    {
        DeployWithScript = 0,
        DeployWithConfigurationFile = 1,
        ChangeChannel = 2,
        RollBack = 3,
        Update = 4
    }
"
Add-Type -TypeDefinition $enum2 -ErrorAction SilentlyContinue

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


function Download-SCCMOfficeChannelFiles() {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(1,2,3),

        [Parameter()]
	    [String]$OfficeFilesPath = $NULL,

        [Parameter()]
        [ValidateSet("en-us","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
                    "ja-jp","kk-kh","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
                    "tr-tr","uk-ua")]
        [string[]] $Languages = ("en-us"),

        [Parameter()]
        [Bitness] $Bitness = 0
        
    )

    Process {
       . "$PSScriptRoot\Download-OfficeProPlusChannels.ps1"

       $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred")
       $ChannelXml = Get-ChannelXml -FolderPath $OfficeFilesPath -OverWrite $true

       foreach ($Channel in $ChannelList) {
         if ($Channels -contains $Channel) {

            $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
            $latestVersion = Get-BranchLatestVersion -ChannelUrl $selectChannel.URL -Channel $Channel
            $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel

            Download-OfficeProPlusChannels -TargetDirectory $OfficeFilesPath  -Channels $Channel -Version $latestVersion -UseChannelFolderShortName $true -Languages $Languages -Bitness $Bitness

            $cabFilePath = "$env:TEMP/ofl.cab"
            Copy-Item -Path $cabFilePath -Destination "$OfficeFilesPath\ofl.cab" -Force

            $latestVersion = Get-BranchLatestVersion -ChannelUrl $selectChannel.URL -Channel $Channel -FolderPath $OfficeFilesPath -OverWrite $true 
         }
       }
    }
}
 
function Create-SCCMOfficePackage {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(1,2,3),

        [Parameter()]
	    [String]$OfficeFilesPath = $NULL,

        [Parameter()]
        [bool]$IncludeSourceFilesInPackage = $true,

        [Parameter()]
	    [bool]$MoveOfflineFiles = $false,

		[Parameter()]
		[String]$CustomPackageShareName = $null,

	    [Parameter()]	
	    [Bool]$UpdateOnlyChangedBits = $true,

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
       try {
       . "$PSScriptRoot\Download-OfficeProPlusChannels.ps1"

       $cabFilePath = "$OfficeFilesPath\ofl.cab"
       if (Test-Path $cabFilePath) {
            Copy-Item -Path $cabFilePath -Destination "$PSScriptRoot\ofl.cab" -Force
       }

       $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred")
       $ChannelXml = Get-ChannelXml -FolderPath $OfficeFilesPath -OverWrite $false

       foreach ($Channel in $ChannelList) {
         if ($Channels -contains $Channel) {
           $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
           $latestVersion = Get-BranchLatestVersion -ChannelUrl $selectChannel.URL -Channel $Channel -FolderPath $OfficeFilesPath -OverWrite $false

           $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel
           $existingPackage = CheckIfPackageExists
           $LargeDrv = Get-LargestDrive

           $Path = CreateOfficeChannelShare -Path "$LargeDrv\OfficeDeployment"

           $packageName = "Office 365 ProPlus"
           $ChannelPath = "$Path\$Channel"
           $LocalPath = "$LargeDrv\OfficeDeployment"
           $LocalChannelPath = "$LargeDrv\OfficeDeployment\SourceFiles"

           [System.IO.Directory]::CreateDirectory($LocalChannelPath) | Out-Null
                          
           if ($IncludeSourceFilesInPackage) {
               if ($OfficeFilesPath) {
                   $officeFileChannelPath = "$OfficeFilesPath\$ChannelShortName"
                   $officeFileTargetPath = "$LocalChannelPath\$Channel"

                   if (!(Test-Path -Path $officeFileChannelPath)) {
                     throw "Channel Folder Missing: $officeFileChannelPath"
                   }

                   [System.IO.Directory]::CreateDirectory($officeFileTargetPath) | Out-Null

                   if ($MoveOfflineFiles) {
                       Move-Item -Path $officeFileChannelPath -Destination $officeFileTargetPath -Force
                   } else {
                       Copy-Item -Path $officeFileChannelPath -Destination $officeFileTargetPath -Recurse -Force
                   }

                   $cabFilePath = "$OfficeFilesPath\ofl.cab"
                   if (Test-Path $cabFilePath) {
                      Copy-Item -Path $cabFilePath -Destination "$LocalPath\ofl.cab" -Force
                   }
               } else {
                   Download-OfficeProPlusChannels -TargetDirectory $LocalChannelPath -Channels $Channel -Version $latestVersion -UseChannelFolderShortName $true 

                   $cabFilePath = "$env:TEMP/ofl.cab"
                   if (!(Test-Path $cabFilePath)) {
                     Copy-Item -Path $cabFilePath -Destination "$LocalPath\ofl.cab" -Force
                   }
               }
           } else {
             if (Test-Path -Path "$LocalChannelPath\Office") {
                Remove-Item -Path "$LocalChannelPath\Office" -Force -Recurse
             }
           }

           CreateMainCabFiles -LocalPath $LocalPath -ChannelShortName $ChannelShortName -LatestVersion $latestVersion

           $DeploymentFilePath = "$PSSCriptRoot\DeploymentFiles\*.*"
           if (Test-Path -Path $DeploymentFilePath) {
             Copy-Item -Path $DeploymentFilePath -Destination "$LocalPath" -Force -Recurse
           } else {
             throw "Deployment folder missing: $DeploymentFilePath"
           }

           LoadSCCMPrereqs -SiteCode $SiteCode -SCCMPSModulePath $SCCMPSModulePath

           if (!($existingPackage)) {
              $package = CreateSCCMPackage -Name $packageName -Path $Path -Channel $Channel -UpdateOnlyChangedBits $UpdateOnlyChangedBits -CustomPackageShareName $CustomPackageShareName
           } else {
             Write-Host "`tPackage Already Exists: $packageName"
           }

           Write-Host

         }
       }
       } catch {
         throw;
       }
    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Create-SCCMOfficeDeploymentProgram {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(1,2,3),

	    [Parameter()]
	    [SCCMDeploymentType]$DeploymentType = "DeployWithScript",

	    [Parameter()]
	    [String]$ScriptName = "SCCM-OfficeDeploymentScript.ps1",

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
         foreach ($channel in $Channels) {
             LoadSCCMPrereqs -SiteCode $SiteCode -SCCMPSModulePath $SCCMPSModulePath

             $existingPackage = CheckIfPackageExists

             [string]$CommandLine = ""
             [string]$ProgramName = ""

             if ($DeploymentType -eq "DeployWithScript") {
                 $ProgramName = "Deploy $channel Channel With Script"
                 $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive " + `
                                "-NoProfile -WindowStyle Hidden -File .\SCCM-OfficeDeploymentScript.ps1 -Channel $channel -SourceFileFolder SourceFiles"

             } elseif ($DeploymentType -eq "DeployWithConfigurationFile") {
                 $ProgramName = "Deploy $channel Channel With Configuration File"
                 $CommandLine = "Office2016Setup.exe /configure Configuration_UpdateSource.xml"

             }

             [string]$packageId = $null

             $packageId = $existingPackage.PackageId
             if ($packageId) {
                $comment = $DeploymentType.ToString() + "-" + $channel

                CreateSCCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -RequiredPlatformNames $requiredPlatformNames -Comment $comment
             }
         }
    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Create-SCCMOfficeChannelChangeProgram {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(1,2,3),

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
         LoadSCCMPrereqs -SiteCode $SiteCode -SCCMPSModulePath $SCCMPSModulePath

         $existingPackage = CheckIfPackageExists

         [string]$CommandLine = ""
         [string]$ProgramName = ""

         foreach ($channel in $Channels) {
             $ProgramName = "Change Channel to $channel"
             $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\Powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -File .\Change-OfficeChannel.ps1 -Channel $Channel"

             $SharePath = $existingPackage.PkgSourcePath

             $OSSourcePath = "$PSScriptRoot\DeploymentFiles\Change-OfficeChannel.ps1"
             $OCScriptPath = "$SharePath\Change-OfficeChannel.ps1"

             if (!(Test-Path $OSSourcePath)) {
                throw "Required file missing: $OSSourcePath"
             } else {
                 if (!(Test-ItemPathUNC -Path $SharePath -FileName "Change-OfficeChannel.ps1")) {
                    Copy-ItemUNC -SourcePath $OSSourcePath -TargetPath $SharePath -FileName "Change-OfficeChannel.ps1"
                 }

                 [string]$packageId = $existingPackage.PackageId
                 if ($packageId) {
                    $comment = "ChangeChannel-$channel"

                    CreateSCCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -RequiredPlatformNames $requiredPlatformNames -Comment $comment
                 }
             }
         }
    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Create-SCCMOfficeRollBackProgram {
    [CmdletBinding(SupportsShouldProcess=$true)]
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
    Process 
    {
         LoadSCCMPrereqs -SiteCode $SiteCode -SCCMPSModulePath $SCCMPSModulePath

         $existingPackage = CheckIfPackageExists

         [string]$CommandLine = ""
         [string]$ProgramName = ""

         $ProgramName = "Rollback"
         $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\Powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -File .\Change-OfficeChannel.ps1 -Rollback"

         $SharePath = $existingPackage.PkgSourcePath

         $OSSourcePath = "$PSScriptRoot\DeploymentFiles\Change-OfficeChannel.ps1"
         $OCScriptPath = "$SharePath\Change-OfficeChannel.ps1"

         if (!(Test-Path $OSSourcePath)) {
            throw "Required file missing: $OSSourcePath"
         } else {
             if (!(Test-ItemPathUNC -Path $SharePath -FileName "Change-OfficeChannel.ps1")) {
                Copy-ItemUNC -SourcePath $OSSourcePath -TargetPath $SharePath -FileName "Change-OfficeChannel.ps1"
             }

             [string]$packageId = $existingPackage.PackageId
             if ($packageId) {
                $comment = "RollBack"

                CreateSCCMProgram -Name $ProgramName -PackageID $packageId -CommandLine $CommandLine -RequiredPlatformNames $requiredPlatformNames -Comment $comment
             }
         }
    }
    End
    {
        Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
        Set-Location $startLocation    
    }
}

function Distribute-SCCMOfficePackage {
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
        [OfficeChannel[]] $Channels = @(1,2,3),

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
               $latestVersion = Get-BranchLatestVersion -ChannelUrl $selectChannel.URL -Channel $ChannelName
               $ChannelShortName = ConvertChannelNameToShortName -ChannelName $ChannelName
               $package = CheckIfPackageExists

               LoadSCCMPrereqs -SiteCode $SiteCode -SCCMPSModulePath $SCCMPSModulePath

               if ($package) {
                    [string]$packageName = $package.Name

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

function Deploy-SCCMOfficeProgram {
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

        [Parameter(Mandatory=$true)]
        [SCCMOfficeProgramType] $ProgramType,
        
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
                $latestVersion = Get-BranchLatestVersion -ChannelUrl $selectChannel.URL -Channel $ChannelName
                $ChannelShortName = ConvertChannelNameToShortName -ChannelName $ChannelName
                $package = CheckIfPackageExists

                LoadSCCMPrereqs -SiteCode $SiteCode -SCCMPSModulePath $SCCMPSModulePath

                $pType = ""

                Switch ($ProgramType) {
                    "DeployWithScript" { $pType = "DeployWithScript-$Channel" }
                    "DeployWithConfigurationFile" { $pType = "DeployWithConfigurationFile-$Channel" }
                    "ChangeChannel" { $pType = "ChangeChannel-$Channel" }
                    "RollBack" { $pType = "RollBack" }
                    "Update" { $pType = "Update" }
                }

                $Program = Get-CMProgram | Where {$_.Comment -eq $pType }
                $programName = $Program.ProgramName

                $packageName = "Office 365 ProPlus"
                if ($package) {
                   if ($Program) {
                        $packageDeploy = Get-CMDeployment | where {$_.PackageId -eq $package.PackageId -and $_.ProgramName -eq $programName }
                        if ($packageDeploy.Count -eq 0) {
                            try {
                                $packageId = $package.PackageId

                                if ($Program) {
                                    $ProgramName = $Program.ProgramName

     	                            Start-CMPackageDeployment -CollectionName "$Collection" -PackageId $packageId -ProgramName "$ProgramName" `
                                                                -StandardProgram  -DeployPurpose Available -RerunBehavior AlwaysRerunProgram `
                                                                -ScheduleEvent AsSoonAsPossible -FastNetworkOption RunProgramFromDistributionPoint `
                                                                -SlowNetworkOption RunProgramFromDistributionPoint `
                                                                -AllowSharedContent $false

                                    Update-CMDistributionPoint -PackageId $package.PackageId

                                    Write-Host "Deployment created for: $packageName ($ProgramName)"
                                } else {
                                    Write-Host "Could Not find Program in Package for Type: $ProgramType - Channel: $ChannelName" -ForegroundColor White -BackgroundColor Red
                                }
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
                          Write-Host "Deployment already exists for: $packageName ($ProgramName)"
                        }
                   } else {
                        Write-Host "Could Not find Program in Package for Type: $ProgramType - Channel: $ChannelName" -ForegroundColor White -BackgroundColor Red
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



function Test-ItemPathUNC() {    [CmdletBinding()]	
    Param
	(	    [Parameter(Mandatory=$true)]
	    [String]$Path,	    [Parameter(Mandatory=$true)]
	    [String]$FileName    )    Process {       $drvLetter = FindAvailable       try {           New-PSDrive -Name $drvLetter -PSProvider FileSystem -Root $Path | Out-Null           $target = $drvLetter + ":\" + $FileName           $result = Test-Path -Path $target            return $result       } finally {         Remove-PSDrive $drvLetter       }    }}

function Copy-ItemUNC() {    [CmdletBinding()]	
    Param
	(	    [Parameter(Mandatory=$true)]
	    [String]$SourcePath,	    [Parameter(Mandatory=$true)]
	    [String]$TargetPath,	    [Parameter(Mandatory=$true)]
	    [String]$FileName    )    Process {       $drvLetter = FindAvailable       try {          $target = $drvLetter + ":\"          New-PSDrive -Name $drvLetter -PSProvider FileSystem -Root $TargetPath | Out-Null          Copy-Item -Path $SourcePath -Destination $target -Force       } finally {         Remove-PSDrive $drvLetter       }    }}

function FindAvailable() {
   $drives = Get-PSDrive | select Name

   for($n=90;$n -gt 68;$n--) {
      $letter= [char]$n
      $exists = $drives | where { $_ -eq $letter }
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

function CreateMainCabFiles() {
    [CmdletBinding()]	
    Param
	(
		[Parameter(Mandatory=$true)]
		[String]$LocalPath = "",

        [Parameter(Mandatory=$true)]
        [String] $ChannelShortName,

        [Parameter(Mandatory=$true)]
        [String] $LatestVersion
	) 
    Process {
        $versionFile32 = "$LocalPath\$ChannelShortName\Office\Data\v32_$LatestVersion.cab"
        $v32File = "$LocalPath\$ChannelShortName\Office\Data\v32.cab"
        $versionFile64 = "$LocalPath\$ChannelShortName\Office\Data\v64_$LatestVersion.cab"
        $v64File = "$LocalPath\$ChannelShortName\Office\Data\v64.cab"

        if (Test-Path -Path $versionFile32) {
            Copy-Item -Path $versionFile32 -Destination $v32File -Force
        }

        if (Test-Path -Path $versionFile64) {
            Copy-Item -Path $versionFile64 -Destination $v64File -Force
        }
    }
}

function CheckIfPackageExists() {
    [CmdletBinding()]	
    Param
	(

    )
    Begin
    {
        $startLocation = Get-Location
    }
    Process {
       LoadSCCMPrereqs

       $packageName = "Office 365 ProPlus"

       $existingPackage = Get-CMPackage | Where { $_.Name -eq $packageName }
       if ($existingPackage) {
         return $existingPackage
       }

       return $null
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

       $packageName = "Office 365 ProPlus"

       $existingPackage = Get-CMPackage | Where { $_.Name -eq $packageName -and $_.Version -eq $Version }
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
		[String]$CustomPackageShareName = $null,

		[Parameter()]	
		[Bool]$UpdateOnlyChangedBits = $true
	) 

    $package = Get-CMPackage | Where { $_.Name -eq $Name }
    if($package -eq $null -or !$package)
    {
        Write-Host "`tCreating Package: $Name"
        $package = New-CMPackage -Name $Name -Path $path -Version $Version
    } else {
        Write-Host "`t`tPackage Already Exists: $Name"        
    }
		
    Write-Host "`t`tSetting Package Properties"

    $VersionName = "$Channel - $Version"

    if ($CustomPackageShareName) {
	    Set-CMPackage -Id $package.PackageId -Priority Normal -EnableBinaryDeltaReplication $UpdateOnlyChangedBits `
                      -CopyToPackageShareOnDistributionPoint $True -Version $Version -CustomPackageShareName $CustomPackageShareName
    } else {
	    Set-CMPackage -Id $package.PackageId -Priority Normal -EnableBinaryDeltaReplication $UpdateOnlyChangedBits `
                      -CopyToPackageShareOnDistributionPoint $True -Version $Version
    }

    $package = Get-CMPackage | Where { $_.Name -eq $Name -and $_.Version -eq $Version }
    return $package
}

function RemovePreviousSCCMPackages() {
    [CmdletBinding()]	
    Param
	(
		[Parameter()]
		[String]$Name = "Office ProPlus Deployment",
		
		[Parameter()]
		[String]$Version
	) 
    
    if ($Version) {
        $packages = Get-CMPackage | Where { $_.Name -eq $Name -and $_.Version -ne $Version }
        foreach ($package in $packages) {
           $packageName = $package.Name
           $pkversion = $package.Version

           Write-Host "Removing previous version: $packageName - $pkversion"
           Remove-CMPackage -Id $package.PackageId -Force | Out-Null
        }
    }
}


function CreateSCCMProgram() {
    [CmdletBinding()]	
    Param
	(
		[Parameter(Mandatory=$True)]
		[String]$PackageID,
		
		[Parameter(Mandatory=$True)]
		[String]$CommandLine, 

		[Parameter(Mandatory=$True)]
		[String]$Name,
		
		[Parameter(Mandatory=$True)]
		[String]$Comment = $null,

		[Parameter()]
		[String[]] $RequiredPlatformNames = @()

	) 

    $program = Get-CMProgram | Where { $_.PackageID -eq $PackageID -and $_.ProgramName -eq $Name }

    if($program -eq $null -or !$program)
    {
        Write-Host "`t`tCreating Program: $Name ..."	        
	    $program = New-CMProgram -PackageId $PackageID -StandardProgramName $Name -DriveMode RenameWithUnc `
                                 -CommandLine $CommandLine -ProgramRunType OnlyWhenUserIsLoggedOn `
                                 -RunMode RunWithAdministrativeRights -UserInteraction $true -RunType Normal 
    } else {
        Write-Host "`t`tProgram Already Exists: $Name"
    }

    if ($program) {
        Set-CMProgram -InputObject $program -Comment $Comment -StandardProgramName $Name -StandardProgram
    }
}

function CreateOfficeChannelShare() {
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
    Param
	(
	    [Parameter()]
	    [string]$FolderPath = $null,

	    [Parameter()]
	    [bool]$OverWrite = $false
	)

   process {
       $cabPath = "$PSScriptRoot\ofl.cab"
       [bool]$downloadFile = $true

       if (!($OverWrite)) {
          if ($FolderPath) {
              $XMLFilePath = "$FolderPath\ofl.cab"
              if (Test-Path -Path $XMLFilePath) {
                 $downloadFile = $false
              } else {
                throw "File missing $FolderPath\ofl.cab"
              }
          }
       }

       if ($downloadFile) {
           $webclient = New-Object System.Net.WebClient
           $XMLFilePath = "$env:TEMP/ofl.cab"
           $XMLDownloadURL = "http://officecdn.microsoft.com/pr/wsus/ofl.cab"
           $webclient.DownloadFile($XMLDownloadURL,$XMLFilePath)

           if ($FolderPath) {
             [System.IO.Directory]::CreateDirectory($FolderPath) | Out-Null
             $targetFile = "$FolderPath\ofl.cab"
             Copy-Item -Path $XMLFilePath -Destination $targetFile -Force
           }
       }

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
      [string]$ChannelUrl,

      [Parameter(Mandatory=$true)]
      [string]$Channel,

	  [Parameter()]
	  [string]$FolderPath = $null,

	  [Parameter()]
	  [bool]$OverWrite = $false
   )

   process {

       [bool]$downloadFile = $true

       $channelShortName = ConvertChannelNameToShortName -ChannelName $Channel

       if (!($OverWrite)) {
          if ($FolderPath) {
              $CABFilePath = "$FolderPath\$channelShortName\v64.cab"
              if (Test-Path -Path $CABFilePath) {
                 $downloadFile = $false
              } else {
                throw "File missing $FolderPath\$channelShortName\v64.cab"
              }
          }
       }

       if ($downloadFile) {
           $webclient = New-Object System.Net.WebClient
           $CABFilePath = "$env:TEMP/v64.cab"
           $XMLDownloadURL = "$ChannelUrl/Office/Data/v64.cab"
           $webclient.DownloadFile($XMLDownloadURL,$CABFilePath)

           if ($FolderPath) {
             [System.IO.Directory]::CreateDirectory($FolderPath) | Out-Null

             $channelShortName = ConvertChannelNameToShortName -ChannelName $Channel 

             $targetFile = "$FolderPath\$channelShortName\v64.cab"
             Copy-Item -Path $CABFilePath -Destination $targetFile -Force
           }
       }

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

function Setup-SCCMOfficeDeploymentPackageOLD {
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
	[Parameter(Mandatory=$True)]
	[String]$Collection,

	[Parameter()]
	[OfficeBranch]$Branch = $null,

	[Parameter()]
	[InstallType]$InstallType = "ScriptInstall",

	[Parameter()]
	[String]$ScriptName = "SCCM-OfficeDeploymentScript.ps1",

	[Parameter()]
	[String]$Path = $null,

	[Parameter()]
	[String]$SiteCode = $null,
	
	[Parameter()]
	[String]$PackageName = $null,

	[Parameter()]	
	[Bool]$UpdateOnlyChangedBits = $false,

	[Parameter()]
	[String[]] $RequiredPlatformNames = @("All x86 Windows 7 Client", "All x86 Windows 8 Client", "All x86 Windows 8.1 Client", "All Windows 10 Professional/Enterprise and higher (32-bit) Client","All x64 Windows 7 Client", "All x64 Windows 8 Client", "All x64 Windows 8.1 Client", "All Windows 10 Professional/Enterprise and higher (64-bit) Client"),
	
	[Parameter()]
	[string]$DistributionPoint,

	[Parameter()]
	[string]$DistributionPointGroupName,

	[Parameter()]
	[uint16]$DeploymentExpiryDurationInDays = 15,

	[Parameter()]
	[String]$SCCMPSModulePath = $NULL,

	[Parameter()]
	[String]$Source = $null


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

    if ($PackageName) {
       $SavedPackageName = $PackageName
    }

    if ($ProgramName) {
       $SavedProgramName = $ProgramName
    }

    if (!$Path) {
         $Path = CreateOfficeUpdateShare
    }

    if ($Branch) {
        $OfficeFolder = "$Path\Office"

        if (Test-Path $OfficeFolder) {
           Remove-Item $OfficeFolder -Recurse -Force
        }

        $TempPath = $Source + "\" + $Branch + "\*"
        Copy-Item $TempPath $Path -Recurse
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

        $SourceDirectory = "$PSScriptRoot\DeploymentFiles"

        if (Test-Path -Path $SourceDirectory) {
           Copy-Item "$SourceDirectory\*.*" $Path
        }
        
	    Set-Location "$SiteCode`:"	

        $package = CreateSCCMPackage -Name $SavedPackageName -Path $path -UpdateOnlyChangedBits $UpdateOnlyChangedBits
        [string]$CommandLine = ""

        if ($InstallType -eq "ScriptInstall") {
            $SavedProgramName = "ScriptInstall"
            $CommandLine = "%windir%\Sysnative\windowsPowershell\V1.0\powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -File .\SCCM-OfficeDeploymentScript.ps1"
        } else {
            $SavedProgramName = "SetupInstall"
            $CommandLine = "Office2016Setup.exe /configure Configuration_UpdateSource.xml"
        }

        CreateSCCMProgram -Name $SavedProgramName -PackageName $SavedPackageName -CommandLine $CommandLine -RequiredPlatformNames $requiredPlatformNames

        Write-Host "Starting Content Distribution"	

        if ($DistributionPointGroupName) {
	        Start-CMContentDistribution -PackageName $SavedPackageName -DistributionPointGroupName $DistributionPointGroupName
        }

        if ($DistributionPoint) {
            Start-CMContentDistribution -PackageName $SavedPackageName -DistributionPointName $DistributionPoint
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
