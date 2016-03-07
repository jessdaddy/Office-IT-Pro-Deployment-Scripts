function Get-OfficeVersion{

[CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        [string[]]$ComputerName,
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

        $defaultDisplaySet = 'ComputerName','DisplayName','Version', 'ClicktoRun'

        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
        $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    }


    process {

     $results = new-object PSObject[] 0;

     $ComputerName += $env:COMPUTERNAME



     foreach ($computer in $ComputerName) {
        if ($Credentials) {
           $os=Get-WMIObject win32_operatingsystem -Credential $Credentials
        } else {
           $os=Get-WMIObject win32_operatingsystem
        }

        $osArchitecture = $os.OSArchitecture

        if ($Credentials) {
           $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -Credential $Credentials
        } else {
           $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default
        }

        $VersionList = New-Object -TypeName System.Collections.ArrayList
        $PathList = New-Object -TypeName System.Collections.ArrayList
        $PackageList = New-Object -TypeName System.Collections.ArrayList
        $ClickToRunPathList = New-Object -TypeName System.Collections.ArrayList
        $ConfigItemList = New-Object -TypeName System.Collections.ArrayList
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
                foreach ($configId in $configItems.sNames) {
                   $Add = $ConfigItemList.Add($configId.ToUpper())
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
                    $AddItem = $PackageList.Add($packageName.Replace(' ', '').ToLower())
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

           

               $object = New-Object PSObject -Property @{ComputerName = $computer; DisplayName = $name; Version = $version; ClickToRun = $clickToRun }
               $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
               $results += $object
  
               
            }

            
         }

         return $results

         
      }

      
     
    }

}

function Nuke-Office{

    

    $c2rVBS = "OffScrubc2r.vbs"
    $03VBS = "OffScrub03.vbs"
    $07VBS = "OffScrub07.vbs"
    $10VBS = "OffScrub10.vbs"
    $15MSIVBS = "OffScrub_O15msi.vbs"
    $16MSIVBS = "OffScrub_O16msi.vbs"



    $versionTest = Get-OfficeVersion
    $c2r = $versionTest.ClicktoRun

    if($c2r -eq $true){
            $ActionFile = $c2rVBS
        }else{
            #Set script file based on office version, if no office detected continue to next computer skipping this one.
            switch -wildcard ($versionTest.Version)
            {
                "11.*"
                {
                    $ActionFile = $03VBS
                }
                "12.*"
                {
                    $ActionFile = $07VBS
                }
                "14.*"
                {
                    $ActionFile = $10VBS
                }
                "15.*"
                {
                    $ActionFile = $15MSIVBS
                }
                "16.*"
                {
                    $ActionFile = $16MSIVBS
                }
                default 
                {
                    "Did not detect Office on target computer ($computer)."
                    continue
                }
            }
        }

        $time = Get-Date -Format g
        Write-Host ""
        Write-Host ""$time": Removing Office products..."

        wscript $ActionFile

        Do{
            Start-Sleep -Seconds 5
            $cscriptProcess = Get-Process cscript -ErrorAction Ignore
        }
        Until($cscriptProcess -eq $null)
       
        if($cscriptProcess -eq $null){
            Stop-Process -Name cmd -ErrorAction Ignore
            
            $time = Get-Date -Format g
            Write-Host ""
            Write-Host ""$time": Searching for remaining registry keys..."

            $Hives = Get-ChildItem Microsoft.PowerShell.Core\Registry::

            $OfficeRegistries = foreach($Hive in $Hives){
                Get-ChildItem "$($Hive.PSPath)" -Recurse -ErrorAction SilentlyContinue | ? PSPath -like *\Software\Microsoft\Office
            }
            foreach($Item in $OfficeRegistries){
                $time = Get-Date -Format g
                Write-Host ""
                Write-Host ""$time": Removing registry key $Item"
                Remove-Item $Item.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }

            $time = Get-Date -Format g
            Write-Host ""
            Write-Host ""$time": All Office products have been removed."
        }

        #.\Nuke-OfficeRegistry.ps1

}
