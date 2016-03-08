if(Test-Path $env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe){
    & $env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe
}
else{
    & .\OneDriveSetup.exe
}