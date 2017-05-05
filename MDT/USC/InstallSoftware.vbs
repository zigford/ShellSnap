'InstallSoftware.vbs Version 3.0.0

'Changes from v2.10.3 - 23/10/2014
' New logging feature supports the same format as read by CMTrace.
'	Logging now shows thread, date/time log file source for merging 
'	with other longs.

'Changes from v2.10.2 - 10/10/2014
' Added new context "Administrator" to handle MDT logging in the same way
'   as System context logging. Ie, C:\Windows\AppLog
'
'Changes from v2.10.1 - 02/05/2014
' Added variable PROGNATIVE for Program Files
'
'Changes from v2.10.0 - 09/04/2014
' Bugfix for SCCM2012 R2 and architecture detection.
'   As SCCM 2012 R2 client is native x64, sysnative does not exist from
'   a native 64bit process, so I have separated out detection based on 
'   the presence of %PRROCESSOR_ARCHITECHTURE6432% environment var.
'
'Changes from v2.9.7 - 10/03/2014
' Added Exitcode matching for individual commands and removed feature
'   from version 2.9.5
'	New Feature: Prefix a command with the following hashtable: 
'	#1-0#9-3010#4-0# cmd.exe /c regular command
'   Explanation: Begin a command with # to invoke exitcode mapping
'	Enter a number which represents a possible exitcode from the command
'	delimit with a dash followed by the exitcode you want returned by the
'   script for normal processing.
'	In the above example, the command may return ExitCode 1, but the script
'	will translate the exitcode to 0. The command may also return exitcode
'	9, but the script will return exitcode 3010 for processing per normal.
'
'Changes from v2.9.6 - 05/03/2014
' Added Exitcode support for individual commands.
'	Each Line executed will return its exitcode and interrupt/return the
'	exitcode back to the parent process. This enabled SCCM to report back
'	failures correctly instead of blindly moving on. Supported exitcodes:
'	0: Success (Continue to next command)
'	1707: Success (Continue to next command)
'	1618: Another installation is in progress (Stop running commands)
'	1641: Hard restart. (Stop running commands, a restart is impending)
'	3010: Soft restart. (Continue running commands, but finally exit with 
'	3010
'	All other exit codes will be result in a stop running commands and exit
'
'  Also duplicated variables to other common names. 
'	ie, %SCRIPTPATH% = %SCRIPTDIR% and %LOGPATH% = %LOGDIR%
'Changes from v2.9.5 - 11/02/2014
' Added function, Clist command beginning with a # will return exit code of
'   postnumber. IE, line #3010 will return exit code indicating a reboot is
'   required.
'Changes from v2.9.4 - 06/09/2012
' Bugfix, modified Architecture detection to test both env vars: 
'	%PROCESSOR_ARCHITEW6432% and %PROCESSOR_ARCHITECTURE%
'	This resolves architecture detection issues when launched from a 
'	native system32 64bit process
'Changes from v2.9.3 - 13/07/2012
' Added variable SYSNATIVE for System32 regardless of architecture.
' 	PLEASE NOTE, this only works from dos builtin commands like dir, copy
'	del, rmdir. It Does not work for EXE's like xcopy, robocopy
'Changes from v2.9.2
' Added variable for LOGDIR. This is so that when you want to specify an
'	application paramater for logging, you can send the logfile to the same
'	place as the clist logfile
'	eg, msiexec nvivo.msi /l* %LOGDIR%\msilog.log /qb
'	In user context the logfile will go in the temp dir. In system context
'	the logfile will go into AppLog
'Changes from v2.9.1
' Bugfix, DEFAULTPROG and DEFAULTSYS Environment variables were previously
'	nesting environment variable, modified to reflect direct paths which 
'	seemed to only be an issue when running script manually.
'Changes from v2.9
' Bugfix, Environment variables are stored and retreived from Process 
'	context rather than User context. This bug caused variables to be
'	saved into a roaming profile if script was run as a user.
'Changes from v2.8
' Added variable to working dir path. This is so that you can explicitly 
'	call executables which may be registered as an App Path in the registry
'	Initially implemented for Outlook deployment on XP
'	Path to script can be accessed via %SCRIPTPATH%
'Changes from v2.7
' Added support to run clist files from a UNC path. This allows you double
'   click InstallSoftware.vbs or drag and drop a clist file onto 
'   InstallSoftware.vbs from a UNC path for testing purposes.
'Changes from v2.6
' Added support for Windows XP, which does not have a Username variable when
'   running in the system context. Also Logfile support has a different dir
'   when run from user context. Updated for all platforms
' Added feature which allows you to just run the script and it will search
'   for Clist files in the current directory. If multiple Clist files are
'   found all of the commands in all of the clists will be run. Order of
'   running is determined alpha-numerically.
' Added more error reporting and handling for log files. Script will now exit
'   more gracefully if the log file cannot be opened for append.
'Changes from v2.5
' Added special variables for 64 and 32 bit registry entries. See
'   instructions for details
'Changes from v2.4
' Added custom error handling of command execution.
'  Previously if a command in the clist was not found
'   the execution would end with no feedback in the log.
'   Now an error level is generated (Exit code 1) with a message in the Log. 
'Changes from v2.3
' Under System context a new log folder is dedicated to log files
'  Folder is %Windir%\AppLog
'  Folder will be auto-created on first use.
'  Recommended filename convention for commandlist is:
'  SoftwareTitle-Action.clist (eg, Sophos-PrepareOS.clist, DotNet4-Install.clist)
'Changes from v2.2
' Changed logging to include dates of each line written to text file.
' Increased pause between commands to 5 second
' Fixed issue with detecting user/system context
'Changes from v2.1
' Added built-in dos variables for x86 and x64 independence. See instructions
'  for more details
'Changes from v1.0
' Added logging of Dos command errors as well as Dos command standard output
' Reworded instructions

'Instructions
'Simple InstallSoftware.VBS file
'Call with "cscript //Nologo InstallSoftware.vbs commandlist
'Where commandlist is a plaintext file with a series of commands to be executed
'commands in the commandlist file must be executable files.
'Dos shell commands such as dir, copy, del are not executable files
'To use these types of commands, prefix them with "cmd /c del file.txt"

'What this script does:
'1 Detects whether running under User or System context and writes a logfile
'  System logfile is: C:\Windows\Temp
'  User logfile is: %temp%
'2 Creates Env variables %DEFAULTSYS% and %DEFAULTPROG%, %REGISTRY%, %SCRIPTPATH%
'  %DEFAULTSYS% = SYSWOW64 on x64 and SYSTEMP 32 on x86
'  %DEFAULTPROG% = Program Files (x86) on x64 and Program Files on x86
'  %REGISTRY% = "HKLM\Software" on x86 and "HKLM\Software\Wow6432Node" on x64
'    examle: reg add "%Registry%\Adobe\Reader" /v "Version" /d "11"
'  %SCRIPTPATH% = \\wsp-sccm01\smspkgd$\USC00001 or
'  %SCRIPTPATH% = C:\Windows\Syswow64\CCM\Cache\USC00001
'  %ARCH% = %PROCESSOR_ARCHITECHTURE%
'3 Executes dos commands without spawning a dos window
'4 Logs Dos output to log file

'Example Commands processor agnostic
'1 uninstall_flash_player_%ARCH%.exe -silent
'Example Commands update a folder
'1 robocopy schemes "%DEFAULTPROG%\WimbaCreate\resources\en.lproj\schemes" /E /XC /XN /XO 

'Configure Environment
Dim WshShell, strCommand, oExec, Context, boolRestart, ProcessId, Line
Dim CMDFile()
Line = 0
Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("Wscript.Shell")

ProcessId = CurrProcessId 'now it's ready to consecutive uses along all runtime, for logging purposes
'Get list of commands to be run from Command List file
CMDList = GetCMDList()
'Get location of logfile depending on login conext and set Context variable
LogFile = GetLogFile()
'Initiate logfile and display context.
LogInfo "Script is running in " & Context & " Context"
'Set Architecture Special Folders ENV Variables and Reg Keys
Call SetSYS64()
'Execute Commands
LogInfo "Executing CMDList from " & CMDFile(0)
For Each CMDListSet in CMDList
  For Each CMD in CMDListSet
	Call Install(CMD)
  Next
Next
'Wrap it up
LogInfo "All Commands Executed"
If boolRestart = 1 then
	LogInfo "Pending restart detected exiting with ExitCode: 3010"
	wscript.quit 3010
Else 
	wscript.quit 0
End If

'Functions

function Install(strCommand)
  'Executes each command and redirects stderror and stdout to the logfile
 If Left(strCommand,1) = "#" then
	errPair = Split(strCommand, "#", -1, 1)
	strCommand = LTrim(errPair(UBound(errPair)))
 End If
 On Error Resume Next
  LogInfo "Executing (" & strCommand & ")"
  Set oExec = WshShell.Exec(strCommand)
  If Err Then
    LogInfo "Error # " & CStr(Err.Number) & " " & Err.Description
    LogInfo "Exiting with error code 1"
    wscript.quit 1
  End If
  Do While Not oExec.StdOut.AtEndOfStream
    sLine = oExec.StdOut.ReadLine
    LogInfo sLine
  Loop
  Do While Not oExec.StdErr.AtEndOfStream
    sLine = oExec.StdErr.ReadLine
    LogInfo sLine
  Loop
  
  errCode = CInt(oExec.ExitCode)
  
  If Not IsNull(errPair) Then
	For i=1 to (Ubound(errPair)-1)
		errSet = Split(errPair(i),"-")
		REM LogInfo "Checking errorcode " & CStr(errCode) & " for match in errPair " & errSet(0)
		If errCode = CInt(errSet(0)) Then
			LogInfo "RAW exit code: " & errCode & " translates to exit code " & errSet(1)
			errCode = CInt(errSet(1))
		End If
	Next
  End If
  LogInfo "Processing exit code " & errCode
  
  Select Case errCode
	Case 0
		LogInfo "ExitCode: 0 Success, Continue to next command"
	Case 1618
		LogInfo "ExitCode: 1618 Another installation is running"
		wscript.quit 1618
	Case 1641
		LogInfo "ExitCode: 1641 Forced reboot"
		wscript.quit 1641
	Case 1707
		LogInfo "ExitCode: 1707 Success, Continue to next command"
	Case 3010
		LogInfo "ExitCode: 3010 Soft Reboot, Continue to next command and restart later"
		boolRestart = 1
	Case Else
		LogInfo "Unmatched ExitCode: " & errCode & " - Hard Exit"
		wscript.quit errCode
  End Select
' Pause for 5 seconds to allow the installer to finish properly
Wscript.Sleep 5000

end function

function GetLogFile()
  Set Shell = WshShell.Environment("Process")
UserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
ComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
If ComputerName = Left(UserName,Len(Username)-1) or Username = "SYSTEM" or Username = "%USERNAME%" then
  LogDir = wshShell.ExpandEnvironmentStrings( "%Windir%" ) & "\AppLog\"
  If Not FileSystem.FolderExists(LogDir) Then
    Set NewFolder = FileSystem.CreateFolder(LogDir)
  End If
  Context = "System"
ElseIf Username = "Administrator" then
  LogDir = wshShell.ExpandEnvironmentStrings( "%Windir%" ) & "\AppLog\"
  If Not FileSystem.FolderExists(LogDir) Then
    Set NewFolder = FileSystem.CreateFolder(LogDir)
  End If
  Context = "Administrator"
Else
  LogDir = wshShell.ExpandEnvironmentStrings( "%TEMP%" ) & "\"
  Context = "User"
End If
	Shell( "LOGDIR" ) = LogDir
	Shell( "LOGPATH" ) = LogDir
	clistFileName = CMDFile(0)
	arrPath = Split(clistFileName, "\")
	intIndex = Ubound(arrPath)
	LogThing = arrPath(intIndex)
  GetLogFile = LogDir & LogThing & ".log"
End Function 

function SetSYS64()
  Set Shell = WshShell.Environment("Process")
  Arch64 = wshShell.ExpandEnvironmentStrings( "%PROCESSOR_ARCHITEW6432%" )
  Arch32 = wshShell.ExpandEnvironmentStrings( "%PROCESSOR_ARCHITECTURE%" )
  If Arch64 = "AMD64" then
    Shell( "DEFAULTSYS" ) = wshShell.ExpandEnvironmentStrings( "%SystemRoot%" ) & "\SysWOW64"
    Shell( "SYSNATIVE" ) = wshShell.ExpandEnvironmentStrings( "%SystemRoot%" ) & "\sysnative"
    Shell( "DEFAULTPROG" ) = wshShell.ExpandEnvironmentStrings( "%ProgramFiles(x86)%" )
	Shell( "PROGNATIVE" ) = wshShell.ExpandEnvironmentStrings( "%ProgramW6432%" )
    Shell( "REGISTRY" ) = "HKLM\SOFTWARE\Wow6432Node"
    Shell( "ARCH" ) = "AMD64"
  elseif Arch32 = "AMD64" then
	Shell( "DEFAULTSYS" ) = wshShell.ExpandEnvironmentStrings( "%SystemRoot%" ) & "\SysWOW64"
    Shell( "SYSNATIVE" ) = wshShell.ExpandEnvironmentStrings( "%SystemRoot%" ) & "\system32"
    Shell( "DEFAULTPROG" ) = wshShell.ExpandEnvironmentStrings( "%ProgramFiles(x86)%" )
	Shell( "PROGNATIVE" ) = wshShell.ExpandEnvironmentStrings( "%ProgramFiles%" )
    Shell( "REGISTRY" ) = "HKLM\SOFTWARE\Wow6432Node"
    Shell( "ARCH" ) = "AMD64"
  else
    Shell( "DEFAULTSYS" ) = wshShell.ExpandEnvironmentStrings( "%SystemRoot%" ) & "\System32"
    Shell( "SYSNATIVE" ) = wshShell.ExpandEnvironmentStrings( "%SystemRoot%" ) & "\System32"
    Shell( "DEFAULTPROG" ) = wshShell.ExpandEnvironmentStrings( "%ProgramFiles%" )
	Shell( "PROGNATIVE" ) = wshShell.ExpandEnvironmentStrings( "%ProgramFiles%" )
    Shell( "REGISTRY" ) = "HKLM\SOFTWARE"
    Shell( "ARCH" ) = "X86"
  End if
  'Shell( "SCRIPTPATH" ) = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
  scriptPath = FileSystem.GetFile(Wscript.ScriptFullName).ParentFolder 
  LogInfo "Working Dir is: " + scriptPath
  LogInfo "System Architechture is: " + wshShell.ExpandEnvironmentStrings( "%ARCH%" )
  Shell( "SCRIPTPATH" ) = scriptPath
  Shell( "SCRIPTDIR" ) = scriptPath
End function


function GetCMDList()
  intCount = 0
  ReDim Preserve CMDFile(intCount)
  Set args = WScript.Arguments
  If args.Count = 0 then
    'Find CMD file by yourself
	
	objStartFolder = "." 
 
	Set objFolder = FileSystem.GetFolder(objStartFolder) 
	Set colFiles = objFolder.Files 
 
	For Each objFile in colFiles 
	If FileSystem.GetExtensionName(objFile) = "clist" Then 
		ReDim Preserve CMDFile(intCount)
		CMDFile(intCount) = objFile.Name
		intCount = intCount + 1
    End If 
    Next
  Else
    CMDFile(0) = args.Item(0)
  End If
  If CMDFile(0) = "" then
    wscript.echo "Error, no CMD file: Exiting with error code 1"
	'LogInfo "Error, no CMD file"
    'LogInfo "Exiting with error code 1"
    wscript.quit 1
  End If 
  Const ForReading = 1
  intCount = 0
  For Each ClistFile in CMDFile
	Set objCMDFile = Filesystem.OpenTextFile(ClistFile, ForReading)
    ReDim Preserve Clist(intCount)
	Clist(intCount) = Split(objCMDFile.ReadAll(), VbCrLf)
	intCount = intCount + 1
  Next
  GetCMDList = Clist
  
End Function

Function CurrProcessId
    Dim oShell, sCmd, oWMI, oChldPrcs, oCols, lOut
    lOut = 0
    Set oShell  = CreateObject("WScript.Shell")
    Set oWMI    = GetObject(_
        "winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    sCmd = "/K " & Left(CreateObject("Scriptlet.TypeLib").Guid, 38)
    oShell.Run "%comspec% " & sCmd, 0
    WScript.Sleep 100 'For healthier skin, get some sleep
    Set oChldPrcs = oWMI.ExecQuery(_
        "Select * From Win32_Process Where CommandLine Like '%" & sCmd & "'", ,32)
    For Each oCols In oChldPrcs
        lOut = oCols.ParentProcessId 'get parent
        oCols.Terminate 'process terminated
        Exit For
    Next
    Set oChldPrcs = Nothing
    Set oWMI = Nothing
    Set oShell = Nothing
    CurrProcessId = lOut
End Function

Function dateStamp()
    Dim t 
    t = Now
    dateStamp = Right("0" & Month(t),2)  & "-" & _
    Right("0" & Day(t),2) & "-" & _  
    Year(t)
End Function

Function timeStamp()
    Dim t 
    t = Now
    timeStamp = Right("0" & Hour(t),2) & ":" & _
    Right("0" & Minute(t),2) & ":" & _
    Right("0" & Second(t),2)
End Function


Sub LogInfo(msg)
	'wscript.echo Right("0" & Month(Now), 2)
	Line = Line + 1
	On Error Resume Next
    Const ForAppending = 8
	wscript.echo msg
    Set WriteOut = FileSystem.OpenTextFile(LogFile, ForAppending, True)
	If Err Then
		wscript.echo "Error # " & CStr(Err.Number) & " " & Err.Description
		wscript.echo "Error opening log file for appending"
		wscript.echo "Logfile path: " & LogFile
    End If
    
	WriteOut.WriteLine "<![LOG[" & msg & "]LOG]!><time=""" & timeStamp & ".297-600"" date=""" & dateStamp & """ component=""CLISTEngine"" context="""& Context & """ type=""1"" thread=""" & ProcessID & """ file=""" & LogFile & ":" & Line & """>"
    WriteOut.Close
    Set WriteOut = Nothing

End Sub
