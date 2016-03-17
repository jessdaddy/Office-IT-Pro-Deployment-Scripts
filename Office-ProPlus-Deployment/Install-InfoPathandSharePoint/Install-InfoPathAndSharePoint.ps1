Function Install-InfoPathAndSharePoint {
<#
.Synopsis
Installs infopath and sharepoint
.DESCRIPTION
Installs infopath and sharepoint since they don't come automatically with Office 16
.NOTES   
Name: Install-InfoPathAndSharePoint
DateCreated: 2015-03-17
.PARAMETER InstallInfopath
bool value for whether infopath gets installed or not
.PARAMETER InstallSharePoint
bool value for whether sharepoint gets installed or not
.EXAMPLE
Install-InfoPathAndSharePoint -InstallInfopath $false -InstallSharePoint $true
Description:
Will install sharepoint, but not infopath
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(mandatory=$true)]
    [bool]$InstallInfopath,
    
    [Parameter(mandatory=$true)]
    [bool]$InstallSharePoint
)
process {

    $scriptPath = Get-ScriptPath
    $infoPathFile = $scriptPath + "\infopath_4753-1001_x86_en-us.exe"
    $SharePointFile = $scriptPath + "\sharepointdesigner_32bit.exe"
    
 
    if($InstallInfopath){
        Start-Process -FilePath $infoPathFile -Wait
    }

    if($InstallSharePoint){
        Start-Process -FilePath $SharePointFile -Wait
    }
}
}


Function Get-ScriptPath() {
  [CmdletBinding()]
  param(

  )

  process {
    #get local path
    $scriptPath = "."

    if ($PSScriptRoot) {
        $scriptPath = $PSScriptRoot
    } else {
        $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
    }
    return $scriptPath
  }
}