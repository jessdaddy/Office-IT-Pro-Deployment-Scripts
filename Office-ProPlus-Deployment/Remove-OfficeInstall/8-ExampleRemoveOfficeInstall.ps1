#  Office ProPlus Click-To-Run Deployment Script example
#
#  This script demonstrates how utilize the scripts in OfficeDev/Office-IT-Pro-Deployment-Scripts repository together to create
#  Office ProPlus Click-To-Run deployment script that will be adaptive to the configuration of the computer it is run from.

Process {
 $scriptPath = "."

 if ($PSScriptRoot) {
   $scriptPath = $PSScriptRoot
 } else {
   $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
 }

#Importing all required functions
. $scriptPath\Generate-ODTConfigurationXML.ps1
. $scriptPath\Edit-OfficeConfigurationFile.ps1
. $scriptPath\Install-OfficeClickToRun.ps1
. $scriptPath\Remove-OfficeInstall.ps1

$targetFilePath = $scriptPath + "\configuration.xml"
$version = $null
$targetLogPath = $scriptPath + "\logFile.txt"

$TotalTime = [System.Diagnostics.Stopwatch]::StartNew()

#Small block to make sure there is an XML copy that doesn't get overwritten
$PathXMLToKeep = Split-Path -Parent $targetFilePath
$PathXMLToKeep += "\originalXML\configuration.xml"
if(!(Test-Path $PathXMLToKeep)){
    Generate-ODTConfigurationXml -Languages AllInUseLanguages -TargetFilePath $PathXMLToKeep 
}

#This example will create an Office Deployment Tool (ODT) configuration file and include all of the Languages currently in use on the computer
#from which the script is run. It will then remove the existing Office installation and then it will then remove the Version attribute from the XML to ensure the installation gets the latest version
#when updating an existing install and then it will initiate a install of Office 2016 Click-To-Run.

Generate-ODTConfigurationXml -Languages AllInUseLanguages -TargetFilePath $targetFilePath | Set-ODTAdd -Version $version -TargetFilePath $targetFilePath
Remove-OfficeInstall -LogFile  $targetLogPath
Install-OfficeClickToRun -OfficeVersion Office2016 -TargetFilePath $targetFilePath -LogFile $targetLogPath


            $time = Get-Date -Format g
            Write-Host ""
            Out-File -FilePath $targetLogPath -InputObject "" -Append
            Write-Host ""$time": Total Time for Entire Script: $($TotalTime.Elapsed.ToString())"
            Out-File -FilePath $targetLogPath -InputObject $time"Total Time for Entire Script: "$($TotalTime.Elapsed.ToString()) -Append

# Configuration.xml file for Click-to-Run for Office 365 products reference. https://technet.microsoft.com/en-us/library/JJ219426.aspx

}