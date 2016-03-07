Param(
    [string]$targetFilePath = ".\configuration.xml"
)

Process {
$scriptPath = "."
if($PSScriptRoot){
    $scriptPath = $PSScriptRoot
}
else{
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

#Importing all required functions
. $scriptPath\Generate-ODTConfigurationXml.ps1
. $scriptPath\Nuke-Office-Local.ps1
. $scriptPath\Install-OfficeClickToRun.ps1
. $scriptPath\Edit-OfficeConfigurationFile.ps1

Generate-ODTConfigurationXml -Languages AllInUseLanguages -TargetFilePath $targetFilePath

Nuke-Office

Install-OfficeClickToRun -TargetFilePath $targetFilePath

}
