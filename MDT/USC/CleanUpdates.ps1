Param($OSDPath = $env:SystemDrive)
#Init Task Sequence Stuff
try {
    $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
    $logPath = $tsenv.Value("LogPath")
} catch {
    Write-Host "This script is not running in a task sequence"
    $logPath = $env:windir + "\temp"
    Start-Transcript $logFile
}

$logFile = "$logPath\$($myInvocation.MyCommand).log"

Write-Host "Logging to $logFile"

$CurrentPath = Split-Path -Path $MyInvocation.InvocationName -Parent
$RegFile = "$CurrentPath\VolumeCaches.Reg"


function Get-FreeSpace {
Param($Drive="C:")
    $FreeSpace = (Get-WmiObject -Query "Select * from Win32_LogicalDisk Where DeviceID = ""$Drive""").FreeSpace

    Switch ($FreeSpace) {
        {$_ -gt 1024} {$Space = "$([math]::Floor($_ /1024)) Kb" }
        {$_ -gt 1048576} {$Space = "$([math]::Floor($_ /1024/1024)) Mb" }
        {$_ -gt 1073741824} {$Space =  "$([math]::Floor($_ /1024/1024/1024)) Gb" }
        default { $Space = "$_ Bytes"}
    }
    return $Space
}

Write-Host "CleanUpdates.ps1"
Write-Host "FreeSpace is $(Get-FreeSpace -Drive $OSDPath)"
Write-Host "Applying Reg Settings"
Start-Process -FilePath "$OSDPath\Windows\System32\reg.exe" -ArgumentList "IMPORT $RegFile" -Wait
Write-Host "Running Cleanup Command"
Start-Process -FilePath "$OSDPath\Windows\System32\cleanmgr.exe" -ArgumentList "/sagerun:11" -Wait
Write-Host "FreeSpace is now $(Get-FreeSpace -Drive $OSDPath)"

If (!$tsenv) {Stop-Transcript }