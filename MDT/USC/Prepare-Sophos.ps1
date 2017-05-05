Param($OSDPath = $env:SystemDrive)
#Init Task Sequence Stuff
try {
    $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
    $logPath = $tsenv.Value("LogPath")
    $logFile = "$logPath\$($myInvocation.MyCommand).log"
} catch {
    Write-Host "This script is not running in a task sequence"
    $logPath = $env:windir + "\temp"
    $logFile = "$logPath\$($myInvocation.MyCommand).log"
    Start-Transcript $logFile
}



Write-Host "Logging to $logFile"

$CurrentPath = Split-Path -Path $MyInvocation.InvocationName -Parent

Stop-Service -DisplayName "Sophos Message Router"
Stop-Service -DisplayName "Sophos Agent"
Stop-Service -DisplayName "Sophos AutoUpdate Service"
Remove-ItemProperty -Path "HKLM\Software\Sophos\Messaging System\Router\Private" -Name 'pkc' -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKLM\Software\Sophos\Messaging System\Router\Private" -Name 'pkp' -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKLM\Software\Sophos\Remote Management System\ManagementAgent\Private" -Name pkc -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKLM\Software\Sophos\Remote Management System\ManagementAgent\Private" -Name pkp -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKLM\Software\Wow6432Node\Sophos\Messaging System\Router\Private" -Name pkc -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKLM\Software\Wow6432Node\Sophos\Messaging System\Router\Private" -Name pkp -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKLM\Software\Wow6432Node\Sophos\Remote Management System\ManagementAgent\Private" -Name pkc -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKLM\Software\Wow6432Node\Sophos\Remote Management System\ManagementAgent\Private" -Name pkp -Force -ErrorAction SilentlyContinue
Remove-Item -Path C:\ProgramData\Sophos\AutoUpdate\data\machine_ID.txt -Force -ErrorAction SilentlyContinue
Remove-Item -Path C:\ProgramData\Sophos\AutoUpdate\machine_ID.txt -Force -ErrorAction SilentlyContinue

If (!$tsenv) {Stop-Transcript }