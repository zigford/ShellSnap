[CmdLetBinding()]
Param([switch]$Online,$NetFxSource,$OSDisk)
$ErrorActionPreference = 'Stop'

# Determine where to do the logging 
Try {
    $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment 
    $logPath = $tsenv.Value("LogPath")
    $IsTS = $True
    Write-Verbose "Running inside a task sequence"
} catch {
    $logPath = 'C:\Windows\AppLog'
    $IsTS = $False
    If (-Not (Test-Path -Path $logPath)) { New-Item -Path $logPath -ItemType Directory -Force}
    Write-Verbose "Not running in a tasks sequence"
}

$logFile = "$logPath\$($myInvocation.MyCommand).log"
 
# Create Logfile
Write-Output "Create Logfile" > $logFile
 
Function Logit($TextBlock1){
	$TimeDate = Get-Date
	$OutPut = "$ScriptName - $Section - $TextBlock1 - $TimeDate"
	Write-Output $OutPut >> $logFile
}
 
# Start Main Code Here
$ScriptName = $MyInvocation.MyCommand
 
# Get data
$Section = "Initialization"
If ($IsTS) {
    $OSDisk = $tsenv.Value("OSDisk")
    $ScratchDir = $tsenv.Value("OSDisk") + "\Windows\temp"
    $NetFxSource = $tsenv.Value("SourcePath") + "\sources\sxs"
} Else {
    If (-Not $NetFxSource -and -Not $OSDisk) {
        Write-Error "When not running in a task sequence you must specify `$NetFxSource and `$OSDisk"
    } else {
	$NetFxSource = $NetFxSource.TrimEnd('\')
	$OSDisk = Join-Path -Path $OSDisk -Child '\'
	$ScratchDir = Join-Path -Path $OSDisk -Child 'Windows\temp'
    }
}
$RunningFromFolder = $MyInvocation.MyCommand.Path | Split-Path -Parent 
. Logit "Running from $RunningFromFolder"
. Logit "Property OSDisk is now $OSDisk"
. Logit "Property ScratchDir is now $ScratchDir"
. Logit "Property NetFxSource is now $NetFxSource"
 
$Section = "Installation"
. Logit "Adding .NET Framework 3.5...."
if ($Online) {
    Write-Verbose "Running command dism.exe /Online /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:""$NetFxSource"" /ScratchDir:""$ScratchDir"""
    dism.exe /Online /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:"$NetFxSource" /ScratchDir:"$ScratchDir"
} else {
    Write-Verbose "Running command dism.exe /Image:$OSDisk /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:""$NetFxSource"" /ScratchDir:""$ScratchDir"""
    dism.exe /Image:$OSDisk /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:"$NetFxSource" /ScratchDir:"$ScratchDir"
}