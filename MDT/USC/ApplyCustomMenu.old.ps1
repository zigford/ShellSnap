$ScriptPath = "$env:WinDir\Resources\USC\Scripts"
$ScriptFile = "$ScriptPath\$($MyInvocation.MyCommand.Name)"

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

If ($tsenv) {
    Write-Host "Running in Task Sequence. Lets check if the scheduled task has been added."
    If (-Not (Test-Path -Path $ScriptPath)) {
        New-Item -Path $ScriptPath -ItemType Directory -Force
    }
    
    If (-Not (Test-Path -Path $ScriptFile)) {
        Copy-Item -Path $MyInvocation.MyCommand.Source -Destination $ScriptPath
    }
    Write-Host "Lets check if the Scheduled task has been added"
    & "$env:WinDir\System32\schtasks.exe" /Query /TN StartLayout
    If ($? -eq $False) {
        Write-Host "Scheduled task not found. Creating StartLayout Task"
        & "$env:WinDir\System32\schtasks.exe" /Create /F /RU "SYSTEM" /SC ONSTART /TN StartLayout /TR "'C:\Windows\System32\WindowsPowershell\v1.0\powershell.exe' -File $ScriptFile" /RL Highest /DELAY 0005:00
        If ($?) {Write-Host "Task StartLayout succesfully created"} Else {Write-Host "Failed to create task"}
    }
    Write-Host "Copying Layout file into resources"
    Try {
        $LayoutPath = New-Item -Path $env:WinDir\Resources -ItemType Directory -Name Start-MenuLayouts
        Copy-Item -Path "$(Split-Path -Path $PSCommandPath -Parent)\Default.xml" -Destination $LayoutPath.FullName
    } Catch {
        Write-Host "Unable to copy layout file Default.xml"
    }
} Else {

    #lets check if we are now in the booted operating system.
    $ParentProcess = (Get-Process -Pid (Get-WmiObject -Class Win32_Process -Filter "processid='$PID'").ParentProcessId).ProcessName
    If ($ParentProcess -eq 'svchost' -and $env:USERDOMAIN -eq 'USC') {
        $DUTileDB = "$env:SystemDrive\Users\Default\AppData\Local\TileDataLayer"
        If (Test-Path -Path $DUTileDB) {
            Write-Host "Found default user TileData, renaming to TileDataLayer2"
            Rename-Item $DUTileDB -NewName TileDataLayer2 -Force    
        }

        Import-Module -Name StartLayout
        Write-Host "Applying start layout to $env:SystemDrive"
        Try {
            Import-StartLayout -LayoutPath $env:WinDir\Resources\Start-MenuLayouts\Default.xml -MountPath $env:SystemDrive\
            Write-Host "Succesfully applied Start-Layout"
        } Catch {
            Write-Host "Failed to apply start layout"
        }
        & "$env:WinDir\System32\schtasks.exe" /Delete /TN StartLayout /F
        If ($?) {Write-Host "Successfully removed Scheduled Task"} Else {Write-Host "Failed to delete Scheduled Task"}
    }
}    

#Stop logging
Stop-Transcript