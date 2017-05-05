'Find statestore

Dim WshShell, strCommand, oExec, Context
Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("Wscript.Shell")
Set env = CreateObject("Microsoft.SMS.TSEnvironment") 
Set Shell = WshShell.Environment("User")
ComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
SystemDrive = wshShell.ExpandEnvironmentStrings( "%SystemDrive%" )
WinDir = wshShell.ExpandEnvironmentStrings( "%WINDIR%" )
BackupFileName = ComputerName & ".wim"
BackupFilePath = WinDir & "\Temp\StateStore\"
BackupFullPath = BackupFilePath & BackupFileName
DestFolder = SystemDrive & "_SMSTaskSequence\StateStore\"
DestFile = DestFolder & BackupFileName
If FileSystem.FileExists(BackupFullPath) then
  If NOT FileSystem.FolderExists(DestFolder) then
    FileSystem.CreateFolder DestFolder
  End If
  wscript.echo "Backup file exists in " & BackupFullPath
  wscript.echo "Moving file"
  FileSystem.MoveFile BackupFullPath, DestFile
  wscript.echo "Setting RestoreOSDFile variable"
  env("RestoreOSDFile") = DestFile
ElseIf FileSystem.FileExists(DestFile) then
  wscript.echo "Backup file exists in " & DestFile
  wscript.echo "Setting RestoreOSDFile variable"
  env("RestoreOSDFile") = DestFile
Else
  wscript.echo "cannot find backup file"
End If