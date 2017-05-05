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

Write-Host "Copying Layout file into resources"
Try {
    $LayoutPath = New-Item -Path $env:WinDir\Resources\USC -ItemType Directory -Name Start-MenuLayouts -Force
    Copy-Item -Path "$(Split-Path -Path $PSCommandPath -Parent)\Layout-*.xml" -Destination $LayoutPath.FullName
    Write-Host "Provisioning lnk files"
    'Internet Explorer.lnk' | ForEach-Object {
        If (Test-Path -Path $_) {
            Copy-Item -Path $_ "$env:SystemDrive\ProgramData\Microsoft\Windows\Start Menu\Programs\" -Force
            If ($?) { Write-Host "Succesfully provisioned $_" }
        }
    }
} Catch {
    Write-Host "Unable to copy layout files"
}

Try {
    Import-Module -Name StartLayout
    $OSBuild = Get-OSBuild
    $LayoutFile = "$env:WinDir\Resources\USC\Start-MenuLayouts\Layout-$OSBuild.xml"
    If (Test-Path -Path $LayoutFile) {
        Write-Host "Applying start layout Layout-$OSBuild.xml to $env:SystemDrive"
        Import-StartLayout -LayoutPath $LayoutFile -MountPath $env:SystemDrive\
        #Copy-Item -Path $LayoutFile -Destination $env:SystemDrive\Users\Administrator\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml
        Write-Host "Succesfully applied Start-Layout"
    } Else {
        Write-Host "Could not location $LayoutFile"
    }
} Catch {
    Write-Host "Failed to apply start layout"
}

#Stop logging
Stop-Transcript