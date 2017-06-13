
<#PSScriptInfo

.VERSION 0.1

.GUID 604c8c16-4987-4de8-b860-3a2e2f3a8394

.AUTHOR jpharris

.COMPANYNAME University of the Sunshine Coast

.COPYRIGHT 

.TAGS 

.LICENSEURI 

.PROJECTURI https://github.com/zigford/ShellSnap/tree/master/Testing 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES
 Currently doesn't cleanup after itself and scheduled tasks must be manually deleted and AutoLogin must be manually cleared.


#>

<# 

.DESCRIPTION 
 Sets up a Windows Performance Recorder Trace and configures scheduled tasks to remove the current profile before trace commences 

#> 
[CmdLetBinding()]
Param()

function New-EventTask {
    [CmdLetBinding()]
    Param($EventID,$Source,$EventLog,$ActionPath,$ArgumentList,$TaskName,[switch]$RunAsSystem)

    $Service = New-Object -ComObject ("Schedule.Service")
    $Service.Connect($Env:COMPUTERNAME)
    $RootFolder = $Service.GetFolder("\")
    $Def = $Service.NewTask(0)
    $regInfo = $Def.RegistrationInfo
    $regInfo.Description = $TaskName
    $regInfo.Author = (whoami)
    $settings = $Def.Settings
    $settings.Enabled = $True
    $settings.StartWhenAvailable = $True
    $settings.Hidden = $false
    $Triggers = $Def.Triggers
    $Trigger = $Triggers.Create(0)
    $Trigger.Id = $EventID
    If ($Source) {
        $SubString = "<QueryList><Query Id='0' Path='$EventLog'><Select Path='$EventLog'>*[System[Provider[@Name='$Source'] and EventID=$EventID]]</Select></Query></QueryList>"
    } Else {
        $SubString = "<QueryList><Query Id='0' Path='$EventLog'><Select Path='$EventLog'>*[System[EventID=$EventID]]</Select></Query></QueryList>"
    }
    Write-Verbose "Subscription = $SubString"
    $Trigger.Subscription = $SubString
    $Trigger.Enabled = $True
    $Action = $Def.Actions.Create(0)
    $Action.Path = $ActionPath
    $Action.Arguments = "$ArgumentList"
    Write-Verbose "Adding task $TaskName with args: $ActionPath $ArgumentList"
    If ($RunAsSystem) {
        $Result = $RootFolder.RegisterTaskDefinition($TaskName,$Def,6,'System',$null,1)
    } Else {
        $Result = $RootFolder.RegisterTaskDefinition($TaskName,$Def,6,(Get-CurrUser),$null,3)
    }
    If ($Result) {
        Write-Verbose "Succesfully created Task $TaskName"
    } else {
        Write-Error "Failed to create task $TaskName"
    }
}

function Get-CurrUser {
    [CmdLetBinding()]
    Param()
    Write-Verbose "Getting Current User: "
    $ProcOwner = (Get-WMIObject -Query "Select * from Win32_Process Where ProcessID = $PID").GetOwner()
    $DomainUserName = "$($ProcOwner.Domain)\$($ProcOwner.User)"
    Write-Verbose $DomainUserName
    $DomainUserName
}

#Setup Autologin
function Add-AutoLogin {
    [CmdletBinding()]
    Param(
        $UserName=(Get-CurrUser).Split('\')[1],
        $Domain=(Get-CurrUser).split('\')[0]
    )
    If (-Not $Password) {
        $Password = Read-Host -prompt "Enter the password for the user $UserName" -AsSecureString 
    }
    Write-Verbose "Setting up AutoLogin values for $UserName"
    $WinLogonPath = 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon'
    New-ItemProperty -Path $WinLogonPath -Name DefaultUserName -Value $UserName -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $WinLogonPath -Name DefaultPassword -Value ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))) -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $WinLogonPath -Name DefaultDomainName -Value $Domain -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $WinLogonPath -Name AutoAdminLogon -Value '1' -PropertyType String -Force | Out-Null
}

function Get-WPR {
[CmdLetBinding()]
Param()

    function Search-WPRLocations {
        [CmdletBinding()]
        Param()
        $PossibleLocations = 'C:\Program Files (x86)\Windows Kits\10\Windows Performance Toolkit\wpr.exe' 
        ForEach ($Location in $PossibleLocations) {
            Write-Verbose "Checking $Location for WPR"
            $Success = Test-Path -Path $Location
            If ($Success) {
                Write-Verbose "Found WPR at $Location"
                $WPR = $Location
            }
        }
        $WPR
    }

    While(-Not $Tried -and (-Not (Search-WPRLocations))) {
        Write-Verbose "Unable to find WPR. Downloading"
        $TempDownload = New-Item -Path $env:temp -Name (Get-Random) -ItemType Directory
        Push-Location
        Set-Location $TempDownload
        Invoke-WebRequest -Uri 'https://go.microsoft.com/fwlink/p/?LinkId=845542' -OutFile adksetup.exe
        Start-Process -FilePath .\adksetup.exe -ArgumentList '/quiet /features OptionId.WindowsPerformanceToolkit' -Wait
        Pop-Location
        Remove-Item $TempDownload -Recurse -Force
        $Tried = $True
    }

    $WPR = Search-WPRLocations
    If ($WPR) {
        return $WPR
    } Else {
        Write-Error "Unable to find Windows Performance Recorder"
    }

}

$CancelShutdown ="
shutdown.exe /a
shutdown.exe /l
schtasks /Delete /TN ""Redirect Shutdown to Logoff"" /F
del ""%~f0""
"
$CncelShutDownF = New-Item -Path $env:temp -Name CncelShutd.bat -ItemType File `
    -Value $CancelShutdown -Force
New-EventTask -EventID 1074 -Source User32 -EventLog System -ActionPath $CncelShutDownF `
    -TaskName 'Redirect Shutdown to Logoff'

#Stage Logoff Script and Scheduled Task.
$LogOffScript ="
#Delete Profile for user
Get-WMIObject -Class Win32_UserProfile | 
? LocalPath -match ""$((Get-CurrUser).Split('\')[1])"" | %{`$_.Delete()}
restart-computer
schtasks /Delete /TN ""Wipe Profile and Restart"" /F
Remove-Item -Path `$PSCommandPath
"

$ClrProfileF = New-Item -Path "$env:SystemRoot\Temp" -Name ClrProfileF.ps1 `
    -ItemType File -Value $LogOffScript -Force
New-EventTask -EventID 4 -EventLog 'Microsoft-Windows-User Profile Service/Operational' `
    -ActionPath 'powershell.exe' -ArgumentList "-Exe Bypass -File $ClrProfileF" `
    -TaskName 'Wipe Profile and Restart' -RunAsSystem

Add-AutoLogin

$WPR = Get-WPR
$WPRParams = "-start GeneralProfile.light " `
    + "-onoffscenario boot "`
    + "-numiterations 1 "`
    + "-onoffresultspath $env:SystemRoot\Temp"

Write-Host "Your trace will be saved in $env:SystemRoot\Temp. Press a key to continue`n"

$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
If (Test-Path -Path $WPR) {
    Write-Verbose "Starting WPR with args: $WPR $WPRParams"
    $ProcRes = Start-Process -FilePath $WPR -ArgumentList "$WPRParams" -PassThru -NoNewWindow -Wait
    Write-Host "Exit Code: $($ProcRes.ExitCode)"
} Else {
    Write-Error "Unable to locate Windows Performance Recorder"
}