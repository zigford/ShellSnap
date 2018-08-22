[CmdLetBinding()]
Param()

function Get-OSBuild {
    cmd.exe /c ver 2>$null | ForEach-Object {
        $v = ([regex]'(\d+(\d+|\.)+)+').Matches($_).Value
        if ($v) {
            [Version]::Parse($v).Build
        }
    }
}

try {
    $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
    $logPath = $tsenv.Value("_SMSTSLogPath")
} catch {
    Write-Verbose "This script is not running in a task sequence"
    $logPath = $env:windir + "\AppLog"
}

$logFile = "$logPath\$($myInvocation.MyCommand).log"
Start-Transcript $logFile
Write-Verbose "Logging to $logFile"
$Build = Get-OSBuild
Write-Verbose "Applying Updates for build $Build"
$Updates = Get-ChildItem -Path "$PSScriptRoot\$Build" -File
Switch ($Updates.Count) {
    {$_ -gt 1} {Write-Verbose "There are $($Updates.Count) updates to apply"}
    1 {Write-Verbose "There is 1 update to apply"}
    0 {Write-Verbose "There are no updates to apply"; $Abort = $True}
}
If (-Not $Abort) {
    $Updates | ForEach-Object { 
        $KB = ([regex]'^.*(?<kb>kb\d+-x(64|86)).*').Match($_).Groups|`
            Select-Object -Last 1|`
            Select-Object -ExpandProperty Value
        Write-Verbose "Applying $KB with $($_.FullName) for build $Build"
        $Package = $_
        Switch ($_.Extension) {
            '.msu' {
                Write-Verbose "wusa.exe ""$($Package.FullName)"" /norestart /log:""$logPath\$KB.log"""
                Start-Process -FilePath wusa.exe -ArgumentList """$($Package.FullName)"" /norestart /quiet /log:""$logPath\$KB.log""" -Wait
            }
            '.cab' {
                Write-Verbose "dism.exe /Online /Add-Package /PackagePath:""$($Package.FullName)"" /NoRestart"
                Start-Process -FilePath dism.exe -ArgumentList "/Online /Add-Package /PackagePath:""$($Package.FullName)"" /NoRestart" -Wait
            }
        }
    }
}
#Stop logging
Stop-Transcript