[CmdLetBinding()]
Param()
$ScriptPath = "$env:WinDir\Resources\USC\Scripts"
$ScriptFile = "$ScriptPath\$($MyInvocation.MyCommand.Name)"

<#function Get-OSBuild {
    cmd.exe /c ver 2>$null | ?{$_ -ne ""}|%{$_.Split('.')[-1].TrimEnd(']').Trim()}
}
#>

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

Write-Verbose "Copying Layout file into resources"
Try {
    $LayoutPath = New-Item -Path $env:WinDir\Resources\USC -ItemType Directory -Name Start-MenuLayouts -Force
    Copy-Item -Path "$(Split-Path -Path $PSCommandPath -Parent)\Layout-*.xml" -Destination $LayoutPath.FullName -Force
    Write-Verbose "Provisioning lnk files"
    'Internet Explorer.lnk' | ForEach-Object {
        If (Test-Path -Path $_) {
            Copy-Item -Path $_ "$env:SystemDrive\ProgramData\Microsoft\Windows\Start Menu\Programs\" -Force
            If ($?) { Write-Verbose "Succesfully provisioned $_" }
        }
    }
} Catch {
    Write-Verbose "Unable to copy layout files"
}

Try {
    Import-Module -Name StartLayout
    $OSBuild = Get-OSBuild
    $LayoutFile = "$env:WinDir\Resources\USC\Start-MenuLayouts\Layout-$OSBuild.xml"
    If (Test-Path -Path $LayoutFile) {
#19-08-2018 Newer builds of office use different office shortcut names
#here we will check which type the system has
        logMsg "Testing office shortcut type"  
        $OldIcons = $True
        $NewIcons = $True
        $TestList = 'Word','Outlook','Excel','PowerPoint'
        $IconLoc  = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs"
        ForEach ($Icon in $TestList) {
            $NewPath = "$IconLoc\$Icon.lnk"
            $OldPath = "$IconLoc\$Icon 2016.lnk"
            If (Test-Path -Path $OldPath) {
                $NewIcons = $False
            } ElseIf (Test-Path $NewPath) {
                $OldIcons = $False
            }
        }
        If ($OldIcons -and ($NewIcons -eq $False)) {
            #Nothing to do here as icons are still the old style
            logMsg "Older 2016 icons"
        } ElseIf ($NewIcons -and ($OldIcons -eq $False)) {
            logMsg "New icons without 2016. Update XML on the fly"
            Try {
                Get-Content -Path $LayoutFile | ForEach-Object {
                    If ($_ -match 'Programs\\(?!OneNote)(\w+\s)+2016') {
                        $_ -replace ' 2016',''
                    } else {
                        $_
                    }
                } | Out-File $LayoutFile
            } catch {
                logMsg "Failed to update xml with new icon names"
            }
        }
        Write-Verbose "Applying start layout Layout-$OSBuild.xml to $env:SystemDrive"
        Import-StartLayout -LayoutPath $LayoutFile -MountPath $env:SystemDrive\
        #Copy-Item -Path $LayoutFile -Destination $env:SystemDrive\Users\Administrator\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml
        Write-Verbose "Succesfully applied Start-Layout"
    } Else {
        Write-Verbose "Could not location $LayoutFile"
    }
} Catch {
    Write-Verbose "Failed to apply start layout"
}

#Stop logging
Stop-Transcript
