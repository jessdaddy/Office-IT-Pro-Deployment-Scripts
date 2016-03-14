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

        $UpdateURLPath = $scriptPath + '\'+$Channel

        if(Test-Path $UpdateURLPath){# update reg key here for updateurl
            New-ItemProperty $Office2RClientKey -Name UpdateUrl -PropertyType String -Value $UpdateURLPath -Force
        }

    #find update exe file
    $OfficeUpdatePath = Get-ItemProperty -Path $Office2RClientKey | Select-Object -Property ClientFolder
    Write-Host $OfficeUpdatePath
    $temp = Out-String -InputObject $OfficeUpdatePath
    $temp = $temp.Substring($temp.LastIndexOf('-')+2)
    $temp = $temp.Trim()
    Write-Host $temp

    $temp+= '\OfficeC2RClient.exe'


    $arguments = "/update user displaylevel=false updatepromptuser=false"
    if($Version){
        $arguments+= ' updatetoversion='+$Version
    }


    
    Write-Host $temp

    
    #run update exe file
    & $temp  $arguments


}

}