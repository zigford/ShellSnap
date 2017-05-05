$ScriptPath = "$env:WinDir\Resources\USC\Scripts"
$ScriptFile = "$ScriptPath\$($MyInvocation.MyCommand.Name)"

function Get-OSBuild {
    (cmd.exe /c ver).split('.')[-1].TrimEnd(']')
}

try {
    $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
    $logPath = $tsenv.Value("LogPath")
} catch {
    Write-Host "This script is not running in a task sequence"
    $logPath = $env:windir + "\temp"
}

$logFile = "$logPath\$($myInvocation.MyCommand).log"
Start-Transcript $logFile
Write-Host "Logging to $logFile"

Try {
    $AssociationPath = New-Item -Path $env:WinDir\Resources\USC -ItemType Directory -Name DefaultApps -Force
    $OSBuild = Get-OSBuild
    $DefaultAppFile = "$env:WinDir\Resources\USC\DefaultApps\DefaultApps-$OSBuild.xml"
    Copy-Item -Path "$(Split-Path -Path $PSCommandPath -Parent)\DefaultApps-$OSBuild.xml" -Destination $AssociationPath.FullName
    Write-Host "Succesfully copied DefaultApps-$OSBuild.xml"
} Catch {
    Write-Host "Unable to copy layout file Default-$OSBuild.xml"
}
Write-Host "Applying DefaultApps.xml to $env:SystemDrive"
Try {
    & Dism.exe /Online /Import-DefaultAppAssociations:C:\Windows\Resources\USC\DefaultApps\DefaultApps-$OSBuild.xml
    Write-Host "Succesfully applied Default app associations"
} Catch {
    Write-Host "Failed to apply default app associations"
}

#Stop logging
Stop-Transcript