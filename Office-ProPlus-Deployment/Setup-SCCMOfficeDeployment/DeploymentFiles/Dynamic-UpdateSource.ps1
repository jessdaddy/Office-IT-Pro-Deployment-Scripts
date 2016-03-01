$enum3 = "
using System;
namespace Microsoft.Office
{
    [FlagsAttribute]
    public enum Branches
    {
        Current=0,
        Business=1,
        Validation=2
    }
}
"
try {
Add-Type -TypeDefinition $enum3 -ErrorAction SilentlyContinue
} catch {}

Function Dynamic-UpdateSource {
<#
.Synopsis
Dynamically updates the ODT Configuration Xml Update Source based on the location of the computer
.DESCRIPTION
If Office Click-to-Run is installed the administrator will be prompted to confirm
uninstallation. A configuration file will be generated and used to remove all Office CTR 
products.
.PARAMETER TargetFilePath
Specifies file path and name for the resulting XML file, for example "\\comp1\folder\config.xml".  Is also the source of the XML that will be updated.
.PARAMETER LookupFilePath
Specifies the source of the csv that contains ADSites with their corresponding SourcePath, for example "\\comp1\folder\sources.csv"
.EXAMPLE
Dynamic-UpdateSource -TargetFilePath "\\comp1\folder\config.xml" -LookupFilePath "\\comp1\folder\sources.csv"
Description:
Will Dynamically set the Update Source based a list Provided
#>
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath = $NULL,
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $LookupFilePath = $NULL
    )

     Process{

     #get computer ADSite
     $computerADSite = "ADSite"
     $SourceValue = ""
     
     [bool] $isInPipe = $true
     if (($PSCmdlet.MyInvocation.PipelineLength -eq 1) -or ($PSCmdlet.MyInvocation.PipelineLength -eq $PSCmdlet.MyInvocation.PipelinePosition)) {
        $isInPipe = $false
     }

     $computerADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name

     #looks for sourcepathlookup.csv file in local directory if parameter was not entered.
     if(!$LookupFilePath){
        $LookupFilePath = GetScriptPath
        $LookupFilePath+= "\SourcePathLookup.csv"
     }

     #get csv file for "SourcePath update"
     
     $importedSource = Import-Csv -Path $LookupFilePath -Delimiter ","

     foreach($imp in $importedSource){
        if($imp.ADSite -eq $computerADSite){#try to match source from the ADSite gathered from csv
            $SourceValue = $imp.source
        }
     }     

     if ($SourceValue) {
        Set-ODTAdd -TargetFilePath $TargetFilePath -SourcePath $SourceValue
     } else {
        if ($isInPipe) {
            $results = new-object PSObject[] 0;
            $Result = New-Object �TypeName PSObject 
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "TargetFilePath" -Value $TargetFilePath
            $Result
        } 
     }



    }
}

Function Set-ODTAdd{
<#
.SYNOPSIS
Modifies an existing configuration xml file's add section
.PARAMETER SourcePath
Optional.
The SourcePath value can be set to a network, local, or HTTP path that contains a 
Click-to-Run source. Environment variables can be used for network or local paths.
SourcePath indicates the location to save the Click-to-Run installation source 
when you run the Office Deployment Tool in download mode.
SourcePath indicates the installation source path from which to install Office 
when you run the Office Deployment Tool in configure mode. If you don�t specify 
SourcePath in configure mode, Setup will look in the current folder for the Office 
source files. If the Office source files aren�t found in the current folder, Setup 
will look on Office 365 for them.
SourcePath specifies the path of the Click-to-Run Office source from which the 
App-V package will be made when you run the Office Deployment Tool in packager mode.
If you do not specify SourcePath, Setup will attempt to create an \Office\Data\... 
folder structure in the working directory from which you are running setup.exe.
.PARAMETER Version
Optional. If a Version value is not set, the Click-to-Run product installation streams 
the latest available version from the source. The default is to use the most recently 
advertised build (as defined in v32.CAB or v64.CAB at the Click-to-Run Office installation source).
Version can be set to an Office 2013 build number by using this format: X.X.X.X
.PARAMETER Bitness
Required. Specifies the edition of Click-to-Run for Office 365 product to use: 32- or 64-bit.
.PARAMETER TargetFilePath
Full file path for the file to be modified and be output to.
.PARAMETER Branch
Optional. Specifies the update branch for the product that you want to download or install.
.Example
Set-ODTAdd -SourcePath "C:\Preload\Office" -TargetFilePath "$env:Public/Documents/config.xml"
Sets config SourcePath property of the add element to C:\Preload\Office
.Example
Set-ODTAdd -SourcePath "C:\Preload\Office" -Version "15.1.2.3" -TargetFilePath "$env:Public/Documents/config.xml"
Sets config SourcePath property of the add element to C:\Preload\Office and version to 15.1.2.3
.Notes
Here is what the portion of configuration file looks like when modified by this function:
<Configuration>
  ...
  <Add SourcePath="\\server\share\" Version="15.1.2.3" OfficeClientEdition="32"> 
      ...
  </Add>
  ...
</Configuration>
#>
    Param(

        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $SourcePath = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Version,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $Bitness,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Microsoft.Office.Branches] $Branch = "Current"

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath



        #Load file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

        if ($TargetFilePath) {
           $ConfigFile.Load($TargetFilePath) | Out-Null
        } else {
            if ($ConfigurationXml) 
            {
              $ConfigFile.LoadXml($ConfigurationXml) | Out-Null
              $global:saveLastConfigFile = $NULL
              $global:saveLastFilePath = $NULL
            }
        }

        $global:saveLastConfigFile = $ConfigFile.OuterXml

        #Check for proper root element
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }

        #Get Add element if it exists
        if($ConfigFile.Configuration.Add -eq $null){
            [System.XML.XMLElement]$AddElement=$ConfigFile.CreateElement("Add")
            $ConfigFile.Configuration.appendChild($AddElement) | Out-Null
        }

        #Set values as desired
        if($Branch -ne $null){
            $ConfigFile.Configuration.Add.SetAttribute("Branch", $Branch);
        }

        if([string]::IsNullOrWhiteSpace($SourcePath) -eq $false){
            $ConfigFile.Configuration.Add.SetAttribute("SourcePath", $SourcePath) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('SourcePath')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("SourcePath")
            }
        }

        if([string]::IsNullOrWhiteSpace($Version) -eq $false){
            $ConfigFile.Configuration.Add.SetAttribute("Version", $Version) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('Version')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("Version")
            }
        }

        if([string]::IsNullOrWhiteSpace($Bitness) -eq $false){
            $ConfigFile.Configuration.Add.SetAttribute("OfficeClientEdition", $Bitness) | Out-Null
        } else {
            if ($PSBoundParameters.ContainsKey('OfficeClientEdition')) {
                $ConfigFile.Configuration.Add.RemoveAttribute("OfficeClientEdition")
            }
        }

        $ConfigFile.Save($TargetFilePath) | Out-Null
        $global:saveLastFilePath = $TargetFilePath

        if (($PSCmdlet.MyInvocation.PipelineLength -eq 1) -or `
            ($PSCmdlet.MyInvocation.PipelineLength -eq $PSCmdlet.MyInvocation.PipelinePosition)) {
            Write-Host

            Format-XML ([xml](cat $TargetFilePath)) -indent 4

            Write-Host
            Write-Host "The Office XML Configuration file has been saved to: $TargetFilePath"
        } else {
            $results = new-object PSObject[] 0;
            $Result = New-Object �TypeName PSObject 
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "TargetFilePath" -Value $TargetFilePath
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "SourcePath" -Value $SourcePath
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Version" -Value $Version
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "Bitness" -Value $Bitness
            $Result
        }
    }

}

Function Get-ODTAdd{
<#
.SYNOPSIS
Gets the value of the Add section in the configuration file
.PARAMETER TargetFilePath
Required. Full file path for the file.
.Example
Get-ODTAdd -TargetFilePath "$env:Public\Documents\config.xml"
Returns the value of the Add section if it exists in the specified
file. 
#>
    Param(

        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $ConfigurationXML = $NULL,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string] $TargetFilePath

    )

    Process{
        $TargetFilePath = GetFilePath -TargetFilePath $TargetFilePath

        #Load the file
        [System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument

        if ($TargetFilePath) {
           $ConfigFile.Load($TargetFilePath) | Out-Null
        } else {
            if ($ConfigurationXml) 
            {
              $ConfigFile.LoadXml($ConfigurationXml) | Out-Null
              $global:saveLastConfigFile = $NULL
              $global:saveLastFilePath = $NULL
            }
        }

        #Check that the file is properly formatted
        if($ConfigFile.Configuration -eq $null){
            throw $NoConfigurationElement
        }
        
        $ConfigFile.Configuration.GetElementsByTagName("Add") | Select OfficeClientEdition, SourcePath, Version, Branch
    }

}

Function GetFilePath() {
    Param(
       [Parameter(ValueFromPipelineByPropertyName=$true)]
       [string] $TargetFilePath
    )

    if (!($TargetFilePath)) {
        $TargetFilePath = $global:saveLastFilePath
    }  

    if (!($TargetFilePath)) {
       Write-Host "Enter the path to the XML Configuration File: " -NoNewline
       $TargetFilePath = Read-Host
    } else {
       #Write-Host "Target XML Configuration File: $TargetFilePath"
    }

    return $TargetFilePath
}

Function Format-XML ([xml]$xml, $indent=2) { 
    $StringWriter = New-Object System.IO.StringWriter 
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter 
    $xmlWriter.Formatting = "indented" 
    $xmlWriter.Indentation = $Indent 
    $xml.WriteContentTo($XmlWriter) 
    $XmlWriter.Flush() 
    $StringWriter.Flush() 
    Write-Output $StringWriter.ToString() 
}

Function GetScriptPath() {
 process {
     [string]$scriptPath = "."

     if ($PSScriptRoot) {
       $scriptPath = $PSScriptRoot
     } else {
       #$scriptPath = (Split-Path $MyInvocation.MyCommand.Path) + "\"
       $scriptPath = (Get-Location).Path
     }

     return $scriptPath
 }
} 