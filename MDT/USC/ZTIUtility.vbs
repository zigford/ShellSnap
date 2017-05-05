 
' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      ZTIUtility.vbs
' // 
' // Version:   5.1.1642.01
' // 
' // Purpose:   Common Libraries for Microsoft Deployment Toolkit
' // 
' // ***************************************************************************


Option Explicit

' Public constants

Public Const ForReading = 1
Public Const ForWriting = 2
Public Const ForAppending = 8

Public Const Success = 0
Public Const Failure = 1

Public Const LogTypeInfo = 1
Public Const LogTypeWarning = 2
Public Const LogTypeError = 3
Public Const LogTypeVerbose = 4
Public Const LogTypeDeprecated = 5

Public Const TextCompare = 1

Public Const adOpenStatic = 3
Public Const adLockReadOnly = 1
Public Const adLockOptimistic = 3

Public Const Version = "5.1.1642.01"

' Global variables

Dim oShell, oEnv, oNetwork, oFSO, objWMI, oDrive
Dim oUtility, oLogging, oEnvironment
Dim oStrings
Dim oFileHandling

' Initialization code

On Error Resume Next
Set oUtility = New Utility
oUtility.PrepareEnvironment


function PrnErrValue ( iError )
	PrnErrValue = oLogging.FormatError ( iError )
end function


Sub ProcessResults(iRc)
	ProcessResults = oLogging.ProcessResults ( iRc )
End Sub

function RunNewInstance 
	Dim oScriptClass
	RunNewInstance = oUtility.RunNewInstanceEx ( oUtility.ScriptName, "oScriptClass", "oScriptClass.Main" )
end function 


Function TestAndLog ( iRc , sMessage)

	TestAndLog = oLogging.TRACEEX(iRc, "", sMessage, FALSE)

End function

Function TestAndFail ( iRc, iError, sMessage)

	TestAndFail = oLogging.TRACEEX(iRc, iError, sMessage, TRUE)

End function

'//---------------------------------------------------------------------------
'//  Function:	ConvertBooleanToString()
'//  Purpose:	Perform a Cstr operation manually to prevent localization 
'//             from converting True/False to non-english values.
'//---------------------------------------------------------------------------
Function ConvertBooleanToString(bValue)

	ConvertBooleanToString = oUtility.ConvertBooleanToString(bValue)

End Function



' Public classes

' //////////////////////////////////////////////////////

class Logging

	' //
	' // Logging is a public class used for trace logging within the BDD/MDT scripting Environment
	' //
	' // Assumes: oUtility, oEnvironment, oFSO - Must be initialized and active.
	' //


	' ***  Public variables ***

	Public LogFile
	Public MasterLogFile
	Public Component
	Public Debug
	Public NetworkLogging


	' ***  Private variables ***


	' ***  Constructor and destructor ***

	Private Sub Class_Initialize


		' Set file names and paths

		Component = oUtility.ScriptName
		LogFile = Component & ".log"
		MasterLogFile = "BDD.log"
		NetworkLogging = ""


		' Set debug to false, allow PrepareEnvironment to override

		Debug = False


	End Sub


	' ***  Public methods ***

	' //
	' // Standard MDT logging routine, used for most logging operations
	' //
	' //  Standard Error Types
	' //     Public Const LogTypeInfo        = 1   ' Informational Message
	' //     Public Const LogTypeWarning     = 2   ' Warning Message
	' //     Public Const LogTypeError       = 3   ' Error Message
	' //     Public Const LogTypeVerbose     = 4   ' Verbose Messages only logged when Debug has been set.
	' //     Public Const LogTypeDeprecated  = 5   ' Informational Message that is elevated to Error when Debug is set.
	' //

	Public Function CreateEntry(sLogMsg, iType)
		Dim sTime, sDate, sTempMsg, oLog, bConsole


		' Each of the operations below has the potential to cause a runtime error.
		' However, we must not stop operation if there is a failure, so allways continue.

		On Error Resume Next


		' Special Handling for Debug vs. Non-Debug messages

		If not Debug then

			If iType = LogTypeVerbose Then
				Exit Function  ' Verbose Messages are only displayed when Debug = True
			Elseif iType = LogTypeDeprecated Then
				iType = LogTypeInfo ' Deprecated messages are normally Info messages
			End if

			' Suppress messages containing password

			If Instr(1, sLogMsg, "password", 1) > 0 then
				sLogMsg = "<Message containing password has been suppressed>"
			End if

		Else  ' Debug = True

			If iType = LogTypeVerbose then
				iType = LogTypeInfo
			Elseif iType = LogTypeDeprecated Then
				iType = LogTypeError
			End if

		End if


		' Populate the variables to log

		sTime = Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2) & ".000+000"
		sDate = Right("0"& Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & "-" & Year(Now)
		sTempMsg = "<![LOG[" & sLogMsg & "]LOG]!><time=""" & sTime & """ date=""" & sDate & """ component=""" & Component & """ context="""" type=""" & iType & """ thread="""" file=""" & oUtility.ScriptName & """>"


		' Make sure the LogPath directory exists

		oUtility.VerifyPathExistsEx LogPath, False


		' If debug, echo the message

		bConsole = InStr(1,Wscript.FullName,"CSCRIPT.EXE", vbTextCompare ) <> 0
		If bConsole = True then
			Wscript.echo sLogMsg
		End if


		' Create the log entry

		Set oLog = oFSO.OpenTextFile(LogPath & "\" & LogFile, ForAppending, True)
		oLog.WriteLine sTempMsg
		oLog.Close


		' Create the master log entry

		Set oLog = oFSO.OpenTextFile(LogPath & "\" & MasterLogFile, ForAppending, True)
		oLog.WriteLine sTempMsg
		oLog.Close


		' Write to a network Share Entry

		If NetworkLogging <> "" then
			Set oLog = oFSO.OpenTextFile(NetworkLogging & "\" & MasterLogFile, ForAppending, True)
			oLog.WriteLine sTempMsg
			oLog.Close
		End if

		On error goto 0

	End Function

	' //
	' // Standard MDT Event Messaging routine, used To send event status messages to a networked log file.
	' //
	' //  Standard Error Types
	' //     Public Const LogTypeInfo     = 1   ' Informational Message
	' //     Public Const LogTypeWarning  = 2   ' Warning Message
	' //     Public Const LogTypeError    = 3   ' Error Message
	' //     Public Const LogTypeVerbose  = 4   ' Verbose Messages only logged when Debug has been set.
	' //

	Function CreateEvent(iEventID, iType, sMessage, arrParms)

		Dim re
		Dim sEventFile
		Dim fptr
		Dim sLine
		Dim iMOMType
		Dim sDomain
		Dim sComputer
		Dim sPackage
		Dim sAdvert
		Dim i
		Dim sCurrentStep
		Dim sTotalSteps


		On Error Resume Next

		Set re = New RegExp


		' Log the message specified

		CreateEntry sMessage, iType


		' Make sure an event share was specified

		If not oEnvironment.Item("EventShare") <> "" then
			Exit Function
		End if


		' Make sure the path is accessible

		oUtility.ValidateConnection oEnvironment.Item("EventShare")


		' Create a unique file name on the share.

		sEventFile = oEnvironment.Item("EventShare") & "\" & oUtility.ComputerName & "_" & oFSO.GetTempName
		Set fptr = oFSO.CreateTextFile(sEventFile, True)
		If Err then
			CreateEntry "Unable to create event file " & sEventFile & ": " & Err.Description & " (" & Err.Number & ")", LogTypeError
			Exit function
		End if


		' Build the first line to write

		Select Case iType
		Case LogTypeInfo
			iMOMType = 4
		Case LogTypeWarning
			iMOMType = 2
		Case LogTypeError
			iMOMType = 1
		End Select

		sDomain = oEnvironment.Item("JoinDomain")
		sComputer = oUtility.ComputerName

		If oEnvironment.Item("_SMSTSAdvertID") <> "" then
			sAdvert = oEnvironment.Item("_SMSTSAdvertId")
			sPackage = oEnvironment.Item("_SMSTSPackageID")
		ElseIf oEnvironment.Item("OSDADVERTID") = "" then
			sAdvert = "OSD00000"
			sPackage = "OSD00000"
		Else
			sAdvert = oEnvironment.Item("OSDADVERTID")
			sPackage = oEnvironment.Item("OSDPACKAGEID")
		End if

		sCurrentStep = oEnvironment.Item("_SMSTSNextInstructionPointer")
		sTotalSteps = oEnvironment.Item("_SMSTSInstructionTableSize")

		' Write out the first line (common content)

		sLine = CStr(iEventID) & "," & CStr(iMOMType) & "," & sDomain & "," & sComputer & "," & sPackage & "," & sAdvert & "," & sCurrentStep & "," & sTotalSteps & "," & oEnvironment.Item("DeploymentMethod")
		fptr.WriteLine sLine
		If Err then
			CreateEntry "Unable to write event file " & sEventFile & ": " & Err.Description & " (" & Err.Number & ")", LogTypeError
			Exit function
		End if


		' Write out the second line (insertion strings)

		fptr.WriteLine Join(arrParms, ",")
		If Err then
			CreateEntry "Unable to write event file " & sEventFile & ": " & Err.Description & " (" & Err.Number & ")", LogTypeError
			Exit function
		End if


		' Write out the third line (message)

		fptr.WriteLine sMessage
		If Err then
			CreateEntry "Unable to write event file " & sEventFile & ": " & Err.Description & " (" & Err.Number & ")", LogTypeError
			Exit function
		End if


		' Close the file

		fptr.Close
		Set fptr = Nothing

		CreateEntry "Event " & CStr(iEventID) & " sent: " & sMessage, LogTypeInfo


	End Function



	Function GetDiscoveryArray

		Dim sLine
		Dim dicMac
		Dim arrMac
		Dim i


		' Add the local entries

		sLine = oEnvironment.Item("PHASE") & "," & oEnvironment.Item("AssetTag") & "," & oEnvironment.Item("UUID")
		Set dicMac = oEnvironment.ListItem("MacAddress")
		arrMac = dicMac.Keys

		For i = 0 to 4
			If i > UBound(arrMac) then
				sLine = sLine & ","
			Else
				sLine = sLine & "," & arrMac(i)
			End if
		Next


		' Add the user data entries (which might not be available yet)

		sLine = oEnvironment.Substitute(sLine & "," & oEnvironment.Item("SLShare") & "," & oEnvironment.Item("UDShare") & "," & oEnvironment.Item("UDDir"))


		' Split it into an array and return it

		GetDiscoveryArray = Split(sLine,",")

	End Function


	Public Function CopyLog()

		Dim iRetVal, fptr1, fptr2, sLine, sNewLogFolderName, sLogFile
		Dim sComputer
		Dim sLog


		On Error Resume Next


		' Figure out where to copy the local logfile to

		If oEnvironment.Item("SLShare") = "" then
			oLogging.CreateEntry "Unable to copy log to the network as no SLShare value was specified.", LogTypeInfo
			Exit Function
		End if


		' Make sure the path is accessible
		oUtility.ValidateConnection oEnvironment.Item("SLShare")
		oUtility.VerifyPathExists oEnvironment.Item("SLShare")
		If not oFSO.FolderExists(oEnvironment.Item("SLShare")) then
			oLogging.CreateEntry "An invalid SLShare value of " & oEnvironment.Item("SLShare") & " was specified.", LogTypeWarning
			Exit Function
		End if


		' Figure out the computer name

		sComputer = oUtility.ComputerName


		' Construct the new folder name

		sNewLogFolderName = oEnvironment.Item("SLShare") & "\" & sComputer
		oUtility.VerifyPathExists sNewLogFolderName


		' Copy various logs if they exist

		For each sLog in Array("Scanstate.log", "USMTCapture.log", "Loadstate.log", "USMTRestore.log", "USMTEstimate.log", "ZTIBackup_imagex.log", "LTIApply_imagex.log", _
			"debug.log", "LTIApply_wdsmcast.log", "ZTIAuthorizeDHCP_netsh.log", "ZTIShrinkDrive_diskpart.log", "ZTIBdeFix_dispart.log", "ZTIdefrag.log", _
			"ZTIConfigureDHCP_netsh.log", "ZTIDiskpart_diskpart.log", "ZTIMarkActive_diskpart.log", "ServerMgr.log", "BDD_pkgmgr.log", _
			"SMSTS.LOG", "Debug\Netsetup.log", "wpeinit.log", "PnpEnum.xml", "ZTICopyUSMT.log", "Debug\DCPROMO.LOG", "Debug\DCPROMOUI.LOG", "OSDSetupWizard.log")

			If oFSO.FileExists(oLogging.LogPath & "\" & sLog) then
				oLogging.CreateEntry "Copying " & oLogging.LogPath & "\" & sLog & " to " & sNewLogFolderName & "\" & sLog, LogTypeInfo
				oFSO.CopyFile oLogging.LogPath & "\" & sLog, sNewLogFolderName & "\", True
			ElseIf oFSO.FileExists(oEnvironment.Item("SMSTSLogPath_Cache") & "\" & sLog) then
				oLogging.CreateEntry "Copying " & oEnvironment.Item("SMSTSLogPath_Cache") & "\" & sLog & " to " & sNewLogFolderName & "\" & sLog, LogTypeInfo
				oFSO.CopyFile oEnvironment.Item("SMSTSLogPath_Cache") & "\" & sLog, sNewLogFolderName & "\", True
			ElseIf oFSO.FileExists(oEnvironment.Item("_SMSTSLogPath") & "\" & sLog) then
				oLogging.CreateEntry "Copying " & oEnvironment.Item("_SMSTSLogPath") & "\" & sLog & " to " & sNewLogFolderName & "\" & sLog, LogTypeInfo
				oFSO.CopyFile oEnvironment.Item("_SMSTSLogPath") & "\" & sLog, sNewLogFolderName & "\", True
			ElseIf oFSO.FileExists(oEnv("TEMP") & "\" & sLog) then
				oLogging.CreateEntry "Copying " & oEnv("TEMP") & "\" & sLog & " to " & sNewLogFolderName & "\" & sLog, LogTypeInfo
				oFSO.CopyFile oEnvironment.Item("_SMSTSLogPath") & "\" & sLog, sNewLogFolderName & "\", True
			ElseIf oFSO.FileExists(oEnv("SystemRoot") & "\" & sLog) then
				oLogging.CreateEntry "Copying " & oENV("SystemRoot") & "\" & sLog & " to " & sNewLogFolderName & "\", LogTypeInfo
				oFSO.CopyFile oENV("SystemRoot") & "\" & sLog,  sNewLogFolderName & "\", True
			ElseIf oFSO.FileExists(oEnv("SystemRoot") & "\System32\" & sLog) then
				oLogging.CreateEntry "Copying " & oENV("SystemRoot") & "\System32\" & sLog & " to " & sNewLogFolderName & "\", LogTypeInfo
				oFSO.CopyFile oENV("SystemRoot") & "\System32\" & sLog,  sNewLogFolderName & "\", True
			End if

		Next


		' Copy the Panther Logs

		If oFSO.FileExists(oENV("SystemRoot") & "\Panther\setupact.log") Then

			oUtility.VerifyPathExists sNewLogFolderName & "\Panther"
			OLogging.CreateEntry "Copying " & oENV("SystemRoot") & "\Panther\setupact.log to " & sNewLogFolderName & "\Panther", LogTypeInfo
			oFSO.CopyFile oENV("SystemRoot") & "\Panther\setupact.log",  sNewLogFolderName & "\Panther\", True
		End If

		If oFSO.FileExists(oENV("SystemRoot") & "\Panther\setuperr.log") Then

			oUtility.VerifyPathExists sNewLogFolderName & "\Panther"
			OLogging.CreateEntry "Copying " & oENV("SystemRoot") & "\Panther\setuperr.log to " & sNewLogFolderName & "\Panther", LogTypeInfo
			oFSO.CopyFile oENV("SystemRoot") & "\Panther\setuperr.log",  sNewLogFolderName & "\Panther\", True
		End If

		If oFSO.FileExists(oENV("SystemRoot") & "\Panther\cbs_unattend.log") Then

			oUtility.VerifyPathExists sNewLogFolderName & "\Panther"
			OLogging.CreateEntry "Copying " & oENV("SystemRoot") & "\Panther\cbs_unattend.log to " & sNewLogFolderName & "\Panther", LogTypeInfo
			oFSO.CopyFile oENV("SystemRoot") & "\Panther\cbs_unattend.log",  sNewLogFolderName & "\Panther\", True
		End If

		If oFSO.FileExists(oENV("SystemRoot") & "\Panther\UnattendGC\setupact.log") Then

			oUtility.VerifyPathExists sNewLogFolderName & "\Panther\UnattendGC"
			OLogging.CreateEntry "Copying " & oENV("SystemRoot") & "\Panther\UnattendGC\setupact.log to " & sNewLogFolderName & "\Panther\UnattendGC", LogTypeInfo
			oFSO.CopyFile oENV("SystemRoot") & "\Panther\UnattendGC\setupact.log",  sNewLogFolderName & "\Panther\UnattendGC\", True
		End If

		If oFSO.FileExists(oENV("SystemRoot") & "\Panther\UnattendGC\setuperr.log") Then

			oUtility.VerifyPathExists sNewLogFolderName & "\Panther\UnattendGC"
			OLogging.CreateEntry "Copying " & oENV("SystemRoot") & "\Panther\UnattendGC\setuperr.log to " & sNewLogFolderName & "\Panther\UnattendGC", LogTypeInfo
			oFSO.CopyFile oENV("SystemRoot") & "\Panther\UnattendGC\setuperr.log",  sNewLogFolderName & "\Panther\UnattendGC\", True

		End If			

		If oFSO.FileExists(oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\setupact.log") Then

			oUtility.VerifyPathExists sNewLogFolderName & "\Panther"
			OLogging.CreateEntry "Copying " & oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\setupact.log to " & sNewLogFolderName & "\Panther", LogTypeInfo
			oFSO.CopyFile oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\setupact.log",  sNewLogFolderName & "\Panther\", True
		End If

		If oFSO.FileExists(oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\setuperr.log") Then

			oUtility.VerifyPathExists sNewLogFolderName & "\Panther"
			OLogging.CreateEntry "Copying " & oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\setuperr.log to " & sNewLogFolderName & "\Panther", LogTypeInfo
			oFSO.CopyFile oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\setuperr.log",  sNewLogFolderName & "\Panther\", True
		End If
		
		If oFSO.FileExists(oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\cbs_unattend.log") Then

			oUtility.VerifyPathExists sNewLogFolderName & "\Panther"
			OLogging.CreateEntry "Copying " & oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\cbs_unattend.log to " & sNewLogFolderName & "\Panther", LogTypeInfo
			oFSO.CopyFile oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\cbs_unattend.log",  sNewLogFolderName & "\Panther\", True
		End If
		
		If oFSO.FileExists(oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\UnattendGC\setupact.log") Then

			oUtility.VerifyPathExists sNewLogFolderName & "\Panther\UnattendGC"
			OLogging.CreateEntry "Copying " & oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\UnattendGC\setupact.log to " & sNewLogFolderName & "\Panther\UnattendGC", LogTypeInfo
			oFSO.CopyFile oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\UnattendGC\setupact.log",  sNewLogFolderName & "\Panther\UnattendGC\", True
		End If

		If oFSO.FileExists(oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\UnattendGC\setuperr.log") Then

			oUtility.VerifyPathExists sNewLogFolderName & "\Panther\UnattendGC"
			OLogging.CreateEntry "Copying " & oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\UnattendGC\setuperr.log to " & sNewLogFolderName & "\Panther\UnattendGC", LogTypeInfo
			oFSO.CopyFile oEnvironment.Item("DestinationLogicalDrive") & "\$WINDOWS.~BT\Sources\Panther\UnattendGC\setuperr.log",  sNewLogFolderName & "\Panther\UnattendGC\", True

		End If					



		' Make sure we have a local log file; it might not exist if the disk isn't yet writable.

		sLogFile = LogPath & "\" & MasterLogFile
		If not oFSO.FileExists(sLogFile) then
			oLogging.CreateEntry "Master log file " & sLogFile & " was not found, unable to copy to " & sNewLogFolderName & "\BDD.LOG", LogTypeInfo
			Exit Function
		End if


		' Copy the file contents to the end of the network file. (It might already exist from a previous action, so append to it.)

		oLogging.CreateEntry "Copying log " & sLogFile & " contents to " & sNewLogFolderName & "\BDD.LOG", LogTypeInfo

		Set fptr1 = oFSO.OpenTextFile(sLogFile, ForReading, True)
		If Err then
			oLogging.CreateEntry "Unable to open " & sLogFile & " for reading: " & Err.Description & " (" & Err.Number & ")", LogTypeInfo
			Err.Clear
			Exit Function
		End if

		Set fptr2 = oFSO.OpenTextFile(sNewLogFolderName & "\BDD.LOG", ForAppending, True)
		If Err then
			oLogging.CreateEntry "Unable to open " & sNewLogFolderName & "\BDD.LOG for appending: " & Err.Description & " (" & Err.Number & ")", LogTypeInfo
			Err.Clear
			Exit Function
		End if

		Do while Not fptr1.AtEndOfStream
			sLine = fptr1.readline
			fptr2.writeline sLine
		Loop

		fptr1.Close
		fptr2.Close


		Err.Clear
		On Error Goto 0

	End Function


	Public Function ReportProgress(sMsg, iPercent)

		Dim iMaxPercent
		Dim oProgress
		Dim uStep
		Dim uMaxStep

		' Try to create the progress UI object

		On Error Resume Next
		Set oProgress = CreateObject("Microsoft.SMS.TSProgressUI")
		If Err then
			Err.Clear
			Exit Function
		End if
		On Error Goto 0


		' Update the progress

		On Error Resume Next

		iMaxPercent = 100
		uStep = CLng(oEnvironment.Item("_SMSTSNextInstructionPointer"))
		uMaxStep = CLng(oEnvironment.Item("_SMSTSInstructionTableSize"))
		Call oProgress.ShowActionProgress(oEnvironment.Item("_SMSTSOrgName"), oEnvironment.Item("_SMSTSPackageName"), oEnvironment.Item("_SMSTSCustomProgressDialogMessage"), oEnvironment.Item("_SMSTSCurrentActionName"), (uStep), (uMaxStep), sMsg, (iPercent), (iMaxPercent))
		If Err then
			CreateEntry "Unable to update progress: " & Err.Description & " (" & Err.Number & ")", LogTypeInfo
			ReportProgress = Failure
			Err.Clear
			Exit Function
		End if

		On Error Goto 0


		' Dispose of the object

		Set oProgress = Nothing

	End Function


	Property Get LogPath

		Dim iRetVal
		LogPath = oUtility.LogPath
		'Preserve the existing logpath before creating the new logpath
		
		If Ucase(oUtility.ScriptName) = "ZTIDISKPART" and Left(LogPath,2)<> "X:" Then
			oUtility.VerifyPathExistsEx "X:\MININT\SMSOSD\OSDLOGS", FALSE
			LogPath = "X:\MININT\SMSOSD\OSDLOGS"
		End If
		
		If oEnvironment.Item("LogPath") <> "" AND oEnvironment.Item("LogPath") <> LogPath And Ucase(oUtility.ScriptName) <> "LTICLEANUP" Then

			On Error Resume Next
			iRetVal = oShell.Run("xcopy """ & oEnvironment.Item("LogPath") & """ """ & LogPath & """ /D /s /e /h /y /c",0, true)
			oEnvironment.Item("LogPath") = LogPath
			on Error goto 0

		End If
		If oEnvironment.Item("LogPath") = "" And Ucase(oUtility.ScriptName) <> "LTICLEANUP" Then	

			oEnvironment.Item("LogPath") = LogPath

		End If

	End Property


	Function ReportFailure ( sMessage, iError )

		CreateEvent 41002, LogTypeError, oEnvironment.Substitute( "FAILURE ( " & FormatError(iError) & " ): " & sMessage ), Array(iError)

		' It is possible that we are are not running in the Wscript Host ( HTML Page ).
		on error resume next
			WScript.Quit iError
		on error goto 0

	End function


	' //
	' //  Perform an inline check of the condition and write out the message
	' //    iRC can be either SUCCESS (0), FAILURE(not 0) as defined above.
	' //    or iRC can be Boolean either TRUE or FALSE
	' //  
	' //  Anything other than Success or Failure will log a Warning or Error.
	' //
	Function TRACEEX( iRc, iError, sMessage, bFatal)

		TRACEEX = iRc

		If Err then

			' Error
			TRACEEX = Err.Number
			sMessage = sMessage & " - " & Err.Description
			CreateEntry oEnvironment.Substitute( "FAILURE (Err): " & FormatError(TRACEEX) & ": " & sMessage ), LogTypeWarning
			If bFatal then
				' Fatial Error
				ReportFailure sMessage, iError
			End if

		Elseif (VarType(iRC) = vbInteger) or (VarType(iRC) = vbLong) or (VarType(iRC) = vbBoolean) then

			' iRC is either a Variant Integer (Either SUCCESS or FAILURE), or a Variant Bool (Either True or False)
			
			If ( (iRC = Success or iRC = 3010) and ((VarType(iRC) = vbInteger) or (VarType(iRC) = vbLong)) ) or ( iRC = TRUE and VarType(iRC) = vbBoolean ) then
				CreateEntry oEnvironment.Substitute( "SUCCESS: " & FormatError(iRC) & ": " & sMessage ), LogTypeVerbose
			ElseIf bFatal then
				' Fatial Error
				ReportFailure  FormatError(iRC) & ": " & sMessage, iError
			Else
				CreateEntry oEnvironment.Substitute( "FAILURE: " & FormatError(iRC) & ": " & sMessage ), LogTypeWarning
			End if
			
		Elseif (VarType(iRC) <> vbEmpty) and (VarType(iRC) <> vbNull ) then

			CreateEntry oEnvironment.Substitute( "UNKNOWN: " & TypeName(iRc) & " = " & iRc & " : " & FormatError(iError) & ": " & sMessage ), LogTypeWarning

		End if

	End function


	Function ProcessResults( iRc )
		Dim iMainRc
		Dim sError
		Dim sMainRc

		iMainRc = iRc
		sMainRc = FormatError ( iMainRc )
		If Err then
			iMainRc = Err.Number
			sError = Err.Description
			sMainRc = FormatError ( iMainRc )
			CreateEvent 41002, LogTypeError, "ZTI ERROR - Unhandled error returned by " & oUtility.ScriptName & ": " & sError & " (" & sMainRc & ")", Array(iMainRc)
		ElseIf iRc <> Success then
			CreateEvent 41002, LogTypeError, "ZTI ERROR - Non-zero return code by " & oUtility.ScriptName & ", rc = " & sMainRc, Array(iMainRc)
		Else
			CreateEvent 41001, LogTypeInfo, oUtility.ScriptName & " processing completed successfully.", Array()
		End if
		WScript.Quit iMainRc

	End function

	Function FormatError( iError )

		' Error messages above 0x1000000 are most likely Hex values, print both Hex and Decimal
		If not isNumeric( iError ) then
			FormatError = ""
		ElseIf Abs(iError) >= &H1000000 then
			FormatError = cstr(iError) & "  0x" & right( "00000000" & hex ( iError ), 8 )
		Else
			FormatError = cstr(iError)
		End if

	End function


End Class

' //////////////////////////////////////////////////////




Class Environment

	' ***  Public variables ***

	Public PersistFile


	' ***  Private variables ***

	Private oVariables
	Private dLastModified
	Private osdV4


	' ***  Constructor and destructor ***

	Private Sub Class_Initialize

		PersistFile = "VARIABLES.DAT"


		On Error Resume Next
		Err.Clear
		Set oVariables = oUtility.CreateXMLDOMObject
		If Err then
			' Unable to create XML object
			Err.Clear
		End if
		On Error Goto 0
		dLastModified = 0


		' Create SMSv4 Task Sequence environment

		On Error Resume Next
		Err.Clear
		Set osdV4 = CreateObject("Microsoft.SMS.TSEnvironment")
		If Err then
			Set osdV4 = Nothing
			Err.Clear
		End if
		On Error Goto 0

		Err.Clear

	End Sub


	' ***  Private methods ***

	Function GetOSDV4(sVariable)

		GetOSDV4 = ""
		On Error Resume Next
		If osdV4 is Nothing then
			Exit Function
		Else
			GetOSDV4 = osdV4(sVariable)
			If Err then
				' oLogging.CreateEntry "WARNING - Unable to get SMSv4 Task Sequencer environment: " & Err.Description & " (" & Err.Number & ")", LogTypeWarning
			End if
		End if
		On Error Goto 0
		Err.Clear

	End Function

	Function SetOSDV4(sVariable, sNew)
		On Error Resume Next
		If osdV4 is Nothing then
			SetOSDV4 = False
			Exit Function
		Else
			osdV4(sVariable) = sNew
			If Err then
				' oLogging.CreateEntry "WARNING - Unable to get SMSv4 Task Sequencer environment: " & Err.Description & " (" & Err.Number & ")", LogTypeWarning
				SetOSDV4 = False
			End if
		End if
		On Error Goto 0
		Err.Clear
		SetOSDV4 = True

	End Function

	Property Get VariablesDat

		' If necessary, load the XML file

		If oFSO.FileExists(PersistPath & "\" & PersistFile) then
			If oFSO.GetFile(PersistPath & "\" & PersistFile).DateLastModified > dLastModified then
				oVariables.Load PersistPath & "\" & PersistFile
				dLastModified = oFSO.GetFile(PersistPath & "\" & PersistFile).DateLastModified
			End if
		ElseIf dLastModified = 0 then
			oVariables.LoadXml "<?xml version=""1.0"" ?><MediaVarList Version=""4.00.5345.0000""></MediaVarList>"
			dLastModified = Now  ' The file hasn't been saved yet, but we don't want to reset this
		End if


		' Set the return value to the XML document

		Set VariablesDat = oVariables

	End Property 

	Function GetDAT(sVariable)
		Dim oNode

		GetDAT = ""

		On Error Resume Next
		Set oNode = Nothing
		Set oNode = VariablesDat.DocumentElement.SelectSingleNode("//var[@name='" & UCase(sVariable) & "']")
		On Error Goto 0
		If not (oNode is Nothing) then
			GetDAT = oNode.Text
		End if

	End Function

	Function SetDAT(sVariable, sNew)
		Dim oNode, oCDATA
		If osdv4 is Nothing OR oEnvironment.Item("_SMSTSPACKAGENAME") = "Lite Touch Installation" Then
			If oFSO.FolderExists(PersistPath) then

				' See if the variable is already defined.  If not, append a new node

				Set oNode = VariablesDat.DocumentElement.SelectSingleNode("//var[@name='" & UCase(sVariable) & "']")
				If oNode is Nothing then
					Set oNode = VariablesDat.CreateElement("var")
					VariablesDat.DocumentElement.appendChild oNode
				Else
					If oNode.Text = sNew then
						' oLogging.CreateEntry "Not changing variable " & sVariable & " because the value was not changed.", LogTypeInfo
						SetDAT = True
						Exit Function
					End if
				End if


				' Set the name of the node

				oNode.SetAttribute "name", UCase(sVariable)


				' Set the value of the node

				Set oCDATA = VariablesDat.createCDATASection(sNew)
				If oNode.hasChildNodes then
					oNode.removeChild(oNode.childNodes.item(0))
				End if
				oNode.appendChild(oCDATA)


				' Save the updated XML file

				On Error Resume Next
				VariablesDat.Save PersistPath & "\" & PersistFile
				dLastModified = oFSO.GetFile(PersistPath & "\" & PersistFile).DateLastModified
				If Err then
					oLogging.CreateEntry "WARNING - Unable to persist items to " & PersistPath & "\" & PersistFile & ": " & Err.Description & " (" & Err.Number & ")", LogTypeWarning
					SetDat = False
					Err.Clear
				End if
				On Error Goto 0
			Else
				SetDat=False
			End if
		Else
			SetDat = False
		End If
		SetDat = True
	End Function


	Function ObfuscateEncode(sVariable, sNew)

		Select Case Ucase(sVariable)
		Case "USERID", "USERPASSWORD", "USERDOMAIN", "DOMAINADMIN", "DOMAINADMINPASSWORD", "DOMAINADMINDOMAIN", _
		 "ADMINPASSWORD", "BDEPIN", "TPMOWNERPASSWORD", "ADDSUSERNAME", "ADDSPASSWORD", _
		 "SAFEMODEADMINPASSWORD", "USERNAME", "USERPASSWORD", "PRODUCTKEY"
			ObfuscateEncode = oStrings.Base64Encode(sNew)
		Case Else
			ObfuscateEncode = sNew
		End Select

	End Function


	Function ObfuscateDecode(sVariable, sCurrent)

		Select Case Ucase(sVariable)
		Case "USERID", "USERPASSWORD", "USERDOMAIN", "DOMAINADMIN", "DOMAINADMINPASSWORD", "DOMAINADMINDOMAIN", _
		 "ADMINPASSWORD", "BDEPIN", "TPMOWNERPASSWORD", "ADDSUSERNAME", "ADDSPASSWORD", _
		 "SAFEMODEADMINPASSWORD", "USERNAME", "USERPASSWORD", "PRODUCTKEY"
			ObfuscateDecode = oStrings.Base64Decode(sCurrent)


			' If the variable wasn't a valid base64 string, an empty string will be returned.  Instead of
			' passing that back, return the current value.  (The Base64Decode method should log a warning.)

			If ObfuscateDecode = "" and sCurrent <> "" then
				ObfuscateDecode = sCurrent
			End if

		Case Else
			ObfuscateDecode = sCurrent
		End Select

	End Function


	' ***  Public methods ***

	Public Property Get Exists(sVariable)

		If Item(sVariable) <> "" then
			Exists = True
		Else
			Exists = False
		End if

	End Property


	Public Property Get Item(sVariable)

		Dim sOriginal
		Dim bSync


		' First try TS environment, then (for Lite Touch only) try the XML file
		
		Item = GetOSDV4(sVariable)
		If Item = "" and GetOSDV4("DeploymentMethod") <> "SCCM" then

			' No value retrieved from the task sequence, try the XML file

			Item = GetDat(sVariable)
			If Item <> "" then
				bSync = true
			End if

		End if


		' Decode and sync if not blank

		If Item <> "" then

			' Decode the variable

			sOriginal = Item
			Item = ObfuscateDecode(sVariable, sOriginal)


			' Try to set the value in the task sequence environment (sync)

			If bSync then
				SetOSDV4 sVariable, sOriginal
			End if

		End if

	End Property


	Public Property Let Item(sVariable, sNew)

		Dim sEncoded


		' Encode as required

		sEncoded = ObfuscateEncode(sVariable, sNew)


		' Save to all available environments

		If SetOSDV4(sVariable, sEncoded) or SetDat(sVariable, sEncoded) Then
			oLogging.CreateEntry "Property " & sVariable & " is now = " & sNew, LogTypeInfo
		End If


		' For completeness, set the variable in the process's environment as well

		oEnv(sVariable) = sNew

	End Property


	Public Property Get ListItem(sVariable)

		Dim i
		Dim sPadded

		Set ListItem = CreateObject("Scripting.Dictionary")
		For i = 1 to 999

			sPadded = sVariable & Right("000" & CStr(i), 3)
			If Item(sPadded) <> "" then
				If not ListItem.Exists(Item(sPadded)) then
					ListItem.Add Item(sPadded), ""
				End if
			ElseIf Item(sVariable & CStr(i)) <> "" then
				If not ListItem.Exists(Item(sVariable & CStr(i))) then
					ListItem.Add Item(sVariable & CStr(i)), ""
				End if
			Else
				Exit For  ' Exit on first "not found" entry
			End if

		Next

	End Property

	Public Sub SetListItemEx (sVariable, sNew)

		Dim sElement
		Dim i
		Dim sPadded

		i = 0
		For each sElement in sNew
			i = i + 1
			sPadded = sVariable & Right("000" & CStr(i), 3)
			Item(sPadded) = sElement
		Next


		' Blank out the next in case there was something there

		sPadded = sVariable & Right("000" & CStr(i+1), 3)
		If Exists(sPadded) then
			Item(sPadded) = ""
		End if


		' Blank out the non-list item if it was there

		If exists(sVariable) then
			Item(sVariable) = ""
		End if

	End sub

	Public Property Set ListItem(sVariable, sNew)

		SetListItemEx sVariable, sNew

	End Property

	Public Property Let ListItem(sVariable, sNew)

		SetListItemEx sVariable, sNew

	End Property


	Function Substitute(sVal)

		Dim sReplace, iPos, iEnd, sEval

		' Substitute the appropriate values

		iPos = Instr(sVal, "%")
		While iPos > 0

			' Find ending "%"
			iEnd = Instr(iPos+1, sVal, "%")
			If iEnd > 0 then

				sEval = Mid(sVal, iPos+1, iEnd - iPos - 1)

				sReplace = ""
				If oEnvironment.ListItem(sEval).Count > 0 then
					For each sReplace in oEnvironment.ListItem(sEval).Keys
						Exit For  ' Grab the first value
					Next
				ElseIf oEnvironment.Item(sEval) <> "" then
					sReplace = oEnvironment.Item(sEval)
				End if

				If sReplace <> "" then
					If iPos = 1 then
						sVal = CStr(sReplace) & Mid(sVal, iEnd + 1)
					ElseIf iEnd = Len(sVal) then
						sVal = Left(sVal, iPos - 1) & CStr(sReplace)
					Else
						sVal = Left(sVal, iPos - 1) & CStr(sReplace) & Mid(sVal, iEnd + 1)
					End if
					iPos = Instr(sVal, "%")
				Else
					iPos = iEnd
				End if
			Else
				iPos = iEnd
			End if
		WEnd


		' Expand any environment variables

		sVal = oShell.ExpandEnvironmentStrings(sVal)


		' Finally, look for evaluate blocks

		iPos = Instr(sVal, "#")
		While iPos > 0

			' Find ending "#"
			iEnd = Instr(iPos+1, sVal, "#")
			If iEnd > 0 then
			
				sEval = Mid(sVal, iPos+1, iEnd - iPos - 1)
				
				sReplace = empty
				On Error Resume Next
				sReplace = Eval(sEval)
				On Error Goto 0
				
				If not isEmpty(sReplace) then
					If iPos = 1 then
						sVal = sReplace & Mid(sVal, iEnd + 1)
					ElseIf iEnd = Len(sVal) then
						sVal = Left(sVal, iPos - 1) & sReplace
					Else
						sVal = Left(sVal, iPos - 1) & sReplace & Mid(sVal, iEnd + 1)
					End if
					iPos = Instr(sVal, "#")
				Else
					iPos = iEnd
				End if
			Else
				iPos = iEnd
			End if

		WEnd

		Substitute = Trim(sVal)

	End Function


	Public Property Get PersistPath
		Dim oOrigPersistFile
		Dim oNewPersistFile
		PersistPath = oUtility.LogPath

			If oFSO.FileExists("X:\minint\smsosd\osdlogs\variables.dat") Then
				Set oOrigPersistFile = oFSO.GetFile("X:\minint\smsosd\osdlogs\variables.dat")
				If oFSO.FileExists(PersistPath & "\VARIABLES.DAT") Then
					Set oNewPersistFile = oFSO.GetFile(PersistPath & "\VARIABLES.DAT")

					If oOrigPersistFile.Size > oNewPersistFile.Size Then
						oFSO.CopyFile "X:\minint\smsosd\osdlogs\variables.dat", PersistPath & "\VARIABLES.DAT", true
					End if
				Else
					oFSO.CopyFile "X:\minint\smsosd\osdlogs\variables.dat", PersistPath & "\VARIABLES.DAT", true
				End if

			End if



	End Property


	Public Function Release

		Set osdV4 = Nothing

	End Function

End Class


Class Utility

	' ***  Properties ***

	Public isHTML
	Public isWSH
	Public isMSHTA
	Public isCScript
	Public isWScript
	Public oMSHTA
	Public Arguments

	' ***  Private variables ***

	Dim dicNetworkConnections
	Dim sScriptDir
	Dim oBDDUtility
	Dim oSupportedPlatforms


	' ***  Constructor and destructor ***

	Private Sub Class_Initialize

		Dim re
		Dim arrDrives, i


		' Initialize the objects

		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set oShell = CreateObject("WScript.Shell")
		Set oEnv = oShell.Environment("PROCESS")
		Set oNetwork = CreateObject("WScript.Network")
		Set objWMI = Nothing
		On Error Resume Next
		Set objWMI = GetObject("winmgmts:")
		On Error Goto 0
		Set dicNetworkConnections = CreateObject("Scripting.Dictionary")
		dicNetworkConnections.CompareMode = TextCompare
		Set oBDDUtility = Nothing
		Set oSupportedPlatforms = Nothing


		isHTML = FALSE
		isWSH = FALSE   
		isMSHTA = FALSE
		isCScript = FALSE
		isWScript = FALSE
		set oMSHTA = nothing
		
		on error resume next
			isHTML = IsObject(window.location)
			isWSH = IsObject(WScript)
		on error goto 0

		If isHTML Then
			isMSHTA = document.all.tags("Application").length > 0
			On error resume next
			set oMSHTA = document.all.tags("Application")(0)
			if window.location.hostname <> "" then
				sScriptDir = oFSO.GetParentFolderName( unescape("\\" & window.location.hostname & window.location.pathname ) )
			else
				sScriptDir = oFSO.GetParentFolderName( unescape(window.location.pathname) )
			end if
			On error goto 0

		ElseIf isWSH Then
			isCScript = ucase(right(WScript.FullName,len("Xscript.exe"))) = "CSCRIPT.EXE"
			isWScript = ucase(right(WScript.FullName,len("Xscript.exe"))) = "WSCRIPT.EXE"
			sScriptDir = oFSO.GetParentFolderName(WScript.ScriptFullName)

		End if


		If Mid(sScriptDir, 2, 2) = ":\" then

			' Look to see if this is a mapped drive

			On Error Resume Next
			Set arrDrives = oNetwork.EnumNetworkDrives
			If Err then
				On Error Goto 0
				oLogging.CreateEntry "ERROR - Network is unavailable: " & Err.Description & " (" & Err.Number & ")", LogTypeError
			Else
				On Error Goto 0
				For i = 0 to arrDrives.Count - 1 Step 2
					If arrDrives.Item(i) = UCase(Left(sScriptDir,2)) then
						If Len(sScriptDir) > 3 then
							sScriptDir = arrDrives.Item(i+1) & Mid(sScriptDir, 3)
						Else
							sScriptDir = arrDrives.Item(i+1)
						End if
						Exit For
					End if
				Next
			End if
			On Error Goto 0

		End if

	End Sub


	' ***  Public methods ***

	Public Sub PrepareEnvironment

		Dim sArg

		Set oLogging = New Logging
		Set oEnvironment = New Environment
		Set oStrings = New Strings
		Set oFileHandling = New FileHandling

		set Arguments = GetArguments


		' Loop through all the parameters and turn them into environment variables.  Enforce debug values.

		On Error Resume Next

		For each sArg in Arguments
			If UCase(sArg) = "DEBUG" then
				If UCase(Arguments(sArg)) = "TRUE" or UCase(Arguments(sArg)) = "FALSE" then
					oLogging.CreateEntry "'debug' parameter was specified.", LogTypeInfo
				Else
					oLogging.CreateEntry "Invalid 'debug' parameter specified: " & Arguments(sArg), LogTypeError
					WScript.Quit Failure
				End if
				oEnvironment.Item(sArg) = UCase(Arguments(sArg))
			Else
				oEnvironment.Item(sArg) = Arguments(sArg)
			End if
		Next

		On Error Goto 0
		Err.Clear


		' Log the version

		oLogging.CreateEntry "Microsoft Deployment Toolkit version: " & Version, LogTypeInfo


		' Log where the SMSTS.LOG can be found

		If oEnvironment.Item("_SMSTSLogPath") <> "" then
			oEnvironment.SetDAT "SMSTSLogPath_Cache", oEnvironment.Item("_SMSTSLogPath")
			oLogging.CreateEntry "The task sequencer log is located at " & oEnvironment.Item("_SMSTSLogPath") & "\SMSTS.LOG.  For task sequence failures, please consult this log.", LogTypeInfo
		End if


		' Set a default for debug (if necessary)

		If oEnvironment.Item("Debug") = "" then
			oEnvironment.Item("Debug") = "FALSE"
		End if


		If UCase(oEnvironment.Item("Debug")) = "TRUE" then
			oLogging.Debug = True
		Else
			oLogging.Debug = False
		End if
		
		If oEnvironment.Item("SLShareDynamicLogging") <> "" then
		
			If oEnvironment.Item("UserID") <> "" and oEnvironment.Item("UserDomain") <> "" and oEnvironment.Item("UserPassword") <> "" then

				If oFSO.FolderExists( LogPath ) then

					oLogging.CreateEntry "Write all logging text to " & oEnvironment.Item("SLShareDynamicLogging") , LogTypeInfo
					ValidateConnection oEnvironment.Item("SLShareDynamicLogging")
					oUtility.VerifyPathExists oEnvironment.Item("SLShareDynamicLogging")
					
					If not oFSO.FolderExists(oEnvironment.Item("SLShareDynamicLogging")) then
						oLogging.CreateEntry "An invalid SLShareDynamicLogging value of " & oEnvironment.Item("SLShareDynamicLogging") & " was specified.", LogTypeWarning
					Else
						oLogging.NetworkLogging = oEnvironment.Item("SLShareDynamicLogging")
					End if
					
				End if
		
			End if
			
		End if
		

	End Sub


	Private Function GetArguments
	
		Dim RegExObj 
		Dim Match 

		set RegExObj = New RegExp
		RegExObj.Global = TRUE
		RegExObj.Multiline = TRUE
		RegExObj.IgnoreCase = TRUE
		RegExObj.Pattern = "\/([^\ ""\:\=]+)(?:(?:[\:\=]""([^""]+)"")|(?:[\:\=]([^\ ""]+)))?"
		
		if isWSH then
			set GetArguments = wscript.arguments.Named

		elseif isMSHTA and (not oMSHTA is nothing) then

			Set GetArguments = CreateObject("Scripting.Dictionary")
			GetArguments.CompareMode = vbTextCompare
			
			for each Match in RegExObj.Execute(oMSHTA.CommandLine)
				If not isempty(Match.submatches(0)) then
					if not GetArguments.Exists(Match.submatches(0)) then
						GetArguments.Add Match.submatches(0), Match.submatches(1) & Match.submatches(2) 
					End if
				End if
			next

		elseif IsHTML then
			Set GetArguments = CreateObject("Scripting.Dictionary")
			GetArguments.CompareMode = vbTextCompare
			
			for each Match in oStrings.ForceAsArray( window.location.search,"&")
				If Instr(1,Match,"=",vbTextCompare) <> 0 then
					If not GetArguments.Exists(left(Match,Instr(1,Match,"=",vbTextCompare)-1)) then
						GetArguments.Add left(Match,Instr(1,Match,"=",vbTextCompare)-1), mid(Match,Instr(1,Match,"=",vbTextCompare)+1)
					End if				
				End if
			next

		else 
			Set GetArguments = CreateObject("Scripting.Dictionary")
			GetArguments.CompareMode = vbTextCompare

		end if
	End function


	Function GetPA
		GetPA = Replace(oShell.Environment("PROCESS")("Processor_Architecture"),"amd64","x64",1,-1,vbTextCompare)
	end function


	Public Function RunNewInstanceEx ( sClassName, sClassInstance, sMain )
		Dim oScriptClass
		Dim iScriptRc
		
		' Disable On Error Resume next for advanced debugging
		
		If not oUtility.Arguments.Exists("DebugCapture") then
			On Error Resume next
		End if
		

		execute "Set " & sClassInstance & " = New " & sClassName
		TestAndFail SUCCESS, 5400, "Create object: Set " & sClassInstance & " = New " & sClassName

		If oUtility.Arguments.Exists("TestHook") then

			' Hook for Unit Test code. Allows us to run Unit Tests without modification of the Original script.

			oLogging.CreateEntry "Run Test Script [" & oUtility.Arguments.Item("TestHook") & "]...", LogTypeInfo
			Execute oFSO.OpenTextFile(oUtility.Arguments.Item("TestHook")).ReadAll
			oLogging.CreateEntry "Finished with Test Script! iScriptRc = " & iScriptRc, LogTypeInfo

		End if

		' Standard execution, call the main function in the created class.

		iScriptRc =  eval(sMain) 
		ProcessResults iScriptRc


		Wscript.quit iScriptRc

	end function 
	
	Function ConvertBooleanToString(bValue)

		Dim iRetVal 

		iRetVal = Failure
			
		If bValue = -1 Then
			iRetVal = "True"
		ElseIf bValue = 0 Then
			iRetVal = "False" 
		End If

		ConvertBooleanToString = iRetVal

	End Function	



	Property Get LocalRootPath

		If oEnvironment.GetOSDV4("_SMSTSBootImageID") <> "" then
			LocalRootPath = oEnvironment.GetOSDV4("_SMSTSMDataPath")
		Else
			LocalRootPath=""

			if oEnv("SystemDrive") = "X:" then   ' We're in PE
				If oFSO.GetDrive("X:").DriveType <> 2 then  ' We're definitely not running from a disk
					LocalRootPath = "C:\MININT"
				Else  ' We might be running from a disk

					If Instr(BootDevice, "RAMDISK") > 0 then  ' We're running in a PE 2.0 RAMdisk
						For each oDrive in oFSO.Drives
							If oDrive.DriveType = 2 and oDrive.DriveLetter<>"X" then
								If oDrive.IsReady Then
									If OFSO.FolderExists(ODrive.DriveLetter & ":\MININT") then
										LocalRootPath = oDrive.DriveLetter & ":\MININT"

										Exit For
									End if
								End If
							End if
						Next

						If LocalRootPath = "" Then
							For each oDrive in oFSO.Drives
								If oDrive.DriveType = 2 then
									If oDrive.IsReady Then
										If oFSO.FolderExists(oDrive.DriveLetter & ":\_SMSTaskSequence") then


											LocalRootPath = oDrive.DriveLetter & ":\MININT"
											Exit For

										Else

											LocalRootPath = "X:\MININT"
										End if
									End If
								End if
							Next

						End if

					Else   ' We're running from the disk
						LocalRootPath = "X:\MININT"
					End if
				End if
			Else  ' We're in a full OS
				For each oDrive in oFSO.Drives
					If oDrive.DriveType = 2 then
						If oDrive.IsReady Then
							If OFSO.FolderExists(ODrive.DriveLetter & ":\MININT") then
								LocalRootPath = oDrive.DriveLetter & ":\MININT"
								Exit For
							End if
						End If
					End if
				Next
				If LocalRootPath = "" Then
					For each oDrive in oFSO.Drives
						If oDrive.DriveType = 2 then
							If oDrive.IsReady Then
								If oFSO.FolderExists(oDrive.DriveLetter & ":\_SMSTaskSequence") then
									LocalRootPath = oDrive.DriveLetter & ":\MININT"
									Exit For

								Else
									LocalRootPath = "C:\MININT"
								End if
							End if
						End if
					Next

				End if
			End if
		End if

		oUtility.VerifyPathExists LocalRootPath


	End Property

	Property Get BootDevice
	
		BootDevice = UCase(oShell.Regread("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\SystemBootDevice"))

	End Property


	Property Get LogPath

		If oEnvironment.GetOSDV4("_SMSTSBootImageID") <> "" then
			LogPath = oEnvironment.GetOSDV4("_SMSTSLogPath")
		Else
			LogPath = LocalRootPath & "\SMSOSD\OSDLOGS"
		End if
		oUtility.VerifyPathExistsEx LogPath, False

	End Property


	Property Get StatePath

		StatePath = LocalRootPath & "\StateStore"
		oUtility.VerifyPathExists StatePath

	End Property


	Property Get ScriptName

		On Error Resume Next
		ScriptName = oFSO.GetBaseName(Wscript.ScriptName)
		If Err then
			ScriptName = oFSO.GetBaseName(Unescape(window.location.pathname))
		End if
		On Error Goto 0

	End Property


	Property Get ScriptDir

		ScriptDir = sScriptDir

	End Property


	Public Function ReadIni(file, section, item)

		Dim line, equalpos, leftstring, ini

		ReadIni = ""
		file = Trim(file)
		item = Trim(item)

		On Error Resume Next
		Set ini = oFSO.OpenTextFile( file, 1, False)
		If Err then
			Err.Clear
			Exit Function
		End if
		On Error Goto 0

		Do While (not ini.AtEndOfStream)
			line = ini.ReadLine
			line = oStrings.TrimAllWS(line)
			If LCase(line) = "[" & LCase(section) & "]" and (not ini.AtEndOfStream) Then
				line = ini.ReadLine
				line = oStrings.TrimAllWS(line)
				Do While Left( line, 1) <> "["
					'If InStr( 1, line, item & "=", 1) = 1 Then
					equalpos = InStr(1, line, "=", 1 )
					If equalpos > 0 Then
						leftstring = Left(line, equalpos - 1 )
						leftstring = oStrings.TrimAllWS(leftstring)
						If LCase(leftstring) = LCase(item) Then
							ReadIni = Mid( line, equalpos + 1 )
							ReadIni = oStrings.TrimAllWS(ReadIni)
							Exit Do
						End If
					End If

					If ini.AtEndOfStream Then Exit Do
					line = ini.ReadLine
					line = oStrings.TrimAllWS(line)
				Loop
				Exit Do
			End If
		Loop
		ini.Close

	End Function

	Public Sub WriteIni( file, section, item, myvalue )

		Dim in_section, section_exists, item_exists, wrote, itemtrimmed
		Dim read_ini, write_ini, temp_ini, linetrimmed, line, equalpos
		Dim leftstring

		in_section = False
		section_exists = False
		item_exists = ( ReadIni( file, section, item ) <> "" )
		wrote = False
		file = Trim(file)
		itemtrimmed = Trim(item)
		myvalue = Trim(myvalue)

		temp_ini = oFSO.GetParentFolderName(file) & "\" & oFSO.GetTempName

		Set read_ini = oFSO.OpenTextFile( file, 1, True, False )
		Set write_ini = oFSO.CreateTextFile( temp_ini, False)

		While read_ini.AtEndOfStream = False
			line = read_ini.ReadLine
			linetrimmed = Trim(line)
			If wrote = False Then
				If LCase(line) = "[" & LCase(section) & "]" Then
					section_exists = True
					in_section = True
				ElseIf InStr( line, "[" ) = 1 Then
					in_section = False
				End If
			End If

			If in_section Then
				If itemtrimmed = "" then
					' Do nothing: we want to wipe the section
				ElseIf item_exists = False Then
					write_ini.WriteLine line
					If myvalue <> "" then
						write_ini.WriteLine item & "=" & myvalue
					End if
					wrote = True
					in_section = False
				Else
					equalpos = InStr(1, line, "=", 1 )
					If equalpos > 0 Then
						leftstring = Left(line, equalpos - 1 )
						leftstring = Trim(leftstring)
						If LCase(leftstring) = LCase(item) Then
							If myvalue <> "" then
								write_ini.WriteLine itemtrimmed & "=" & myvalue
							End if
							wrote = True
							in_section = False
						End If
					End If
					If Not wrote Then
						write_ini.WriteLine line
					End If
				End If
			Else
				write_ini.WriteLine line
			End If
		Wend

		If section_exists = False and itemtrimmed <> "" Then ' section doesn't exist
			write_ini.WriteLine
			write_ini.WriteLine "[" & section & "]"
			If myvalue <> "" then
				write_ini.WriteLine itemtrimmed & "=" & myvalue
			End if
		End If

		read_ini.Close
		write_ini.Close
		If oFSO.FileExists(file) then
			oFSO.DeleteFile file, True
		End if
		oFSO.CopyFile temp_ini, file, true
		oFSO.DeleteFile temp_ini, True

	End Sub

	Public Function Sections(file)
		Dim oContents
		Dim line, equalpos, leftstring, ini

		Set oContents = CreateObject("Scripting.Dictionary")
		file = Trim(file)

		On Error Resume Next
		Set ini = oFSO.OpenTextFile( file, 1, False)
		If Err then
			Err.Clear
			Exit Function
		End if
		On Error Goto 0

		Do While ini.AtEndOfStream = False
			line = ini.ReadLine
			line = Trim(line)
			If Left(line,1) = "[" then
				equalpos = Instr(line, "]")
				leftstring = Mid(line, 2, equalpos - 2)
				oContents.Add leftstring, ""
			End if
		Loop

		Set Sections = oContents

	End Function

	Public Function SectionContents(file, section)
		Dim oContents
		Dim line, equalpos, leftstring, ini

		Set oContents = CreateObject("Scripting.Dictionary")
		file = Trim(file)
		section = Trim(section)

		On Error Resume Next
		Set ini = oFSO.OpenTextFile( file, 1, False)
		If Err then
			Err.Clear
			Exit Function
		End if
		On Error Goto 0

		Do While ini.AtEndOfStream = False
			line = ini.ReadLine
			line = Trim(line)
			If LCase(line) = "[" & LCase(section) & "]" Then
				line = ini.ReadLine
				line = Trim(line)
				Do While Left( line, 1) <> "["
					'If InStr( 1, line, item & "=", 1) = 1 Then
					equalpos = InStr(1, line, "=", 1 )
					If equalpos > 0 Then
						leftstring = Left(line, equalpos - 1 )
						leftstring = Trim(leftstring)
						oContents(leftstring) = Trim(Mid(line, equalpos + 1 ))
					End If

					If ini.AtEndOfStream Then Exit Do
					line = ini.ReadLine
					line = Trim(line)
				Loop
				Exit Do
			End If
		Loop
		ini.Close
		Set SectionContents = oContents

	End Function


	Function RegRead( sRegValue )

		on error resume next
		RegRead = oShell.Regread( sRegValue )      
		TestAndLog not isempty(RegRead),  "ReadRead(" & sRegValue & ")"
		on error goto 0

	end Function 

	Function RegWrite ( sKey, sValue )

		on error resume next

		if VarType(sValue) = vbString then
			RegWrite = oShell.RegWrite ( sKey, sValue, "REG_SZ"  )
		elseif VarType(sValue) = vbInteger or VarType(sValue) = vbLong or VarType(sValue) = vbBoolean then 
			RegWrite = oShell.RegWrite ( sKey, sValue, "REG_DWORD" )
		end if 

		TestAndLog err = 0, "RegWriteEx( " & sKey & " , " & sValue & " , REG_XXX )"

		on error goto 0
		
	end Function
	

	Function RunWithHeartbeat(sCmd)

		Dim oExec
		Dim lastHeartbeat
		Dim lastStart
		Dim iHeartbeat
		Dim iMinutes

		' Initialize the last heartbeat time (start the timer) and interval

		lastHeartbeat = Now
		iHeartbeat = 5

		' Start the command

		oLogging.CreateEntry "About to run command: " & sCmd, LogTypeInfo
		lastStart = Now
		Set oExec = oShell.Exec(sCmd)
		Do While oExec.Status = 0

			' Sleep

			WScript.Sleep 500


			' See if it is time for a heartbeat

			If iHeartbeat > 0 and DateDiff("n", lastHeartbeat, Now) > iHeartbeat then
				iMinutes = DateDiff("n", lastStart, Now)
				oLogging.CreateEvent 41003, LogTypeInfo, "ZTI Heartbeat: command has been running for " & CStr(iMinutes) & " minutes (process ID " & oExec.ProcessID & ")", Array(iMinutes)
				lastHeartbeat = Now
			End if

		Loop


		' Return the exit code to the caller

		oLogging.CreateEntry "Return code from command = " & oExec.ExitCode, LogTypeInfo
		RunWithHeartbeat = oExec.ExitCode


	End Function


	Function ComputerName

		Dim re
		Dim sComputer


		Set re = New RegExp


		' Figure out the computer name to include

                If oEnvironment.Item("_SMSTSMachineName") <> "" then
                	If oEnvironment.Item("OSDCOMPUTERNAME") <> "" and Instr(oEnvironment.Item("OSDCOMPUTERNAME"),":") = 0 then
				sComputer = oEnvironment.Item("OSDComputerName")
	                Else
        		         sComputer = oEnvironment.Item("_SMSTSMachineName")
	  	        End If
		ElseIf oEnvironment.Item("OSDCOMPUTERNAME") <> "" and Instr(oEnvironment.Item("OSDCOMPUTERNAME"),":") = 0 then
			sComputer = oEnvironment.Item("OSDCOMPUTERNAME")
		ElseIf oEnvironment.Item("OSDNEWMACHINENAME") <> "" then
			oEnvironment.Item("OSDComputerName") = oEnvironment.Item("OSDNEWMACHINENAME")
			sComputer = oEnvironment.Item("OSDCOMPUTERNAME")
		ElseIf oEnvironment.Item("ComputerName") <> "" then
			oEnvironment.Item("OSDComputerName") = oEnvironment.Item("ComputerName")
			sComputer = oEnvironment.Item("OSDComputerName")
		ElseIf oEnvironment.Item("OSVersion") = "WinPE" and oEnvironment.Item("OSDCOMPUTERNAME") <> "" then
			sComputer = oEnvironment.Item("OSDCOMPUTERNAME")
			re.Pattern = ":"
			sComputer = re.Replace(sComputer, "")
		Else
			oEnvironment.Item("OSDComputerName") = oEnvironment.Item("HostName")
			sComputer = oEnvironment.Item("OSDComputerName")
		End if

		ComputerName = oEnvironment.Substitute(sComputer)

	End Function


	Function FindFile(sFilename, sFoundPath)

		Dim iRetVal
		Dim sDir


		iRetVal = Failure
		sFoundPath = ""


		' Look through the standard locations

		If oEnvironment.Item("DeployRoot") <> "" then

			For each sDir in Array("\", "\Servicing\", "\Tools\", "\USMT\", "\Templates\", "\Scripts\", "\Control\","\x86\","\amd64\")

				If oFSO.FileExists(oEnvironment.Item("DeployRoot") & sDir & sFileName) then
					sFoundPath = oEnvironment.Item("DeployRoot") & sDir & sFileName
					iRetVal = Success
					Exit For
				ElseIf oEnvironment.Item("Architecture") <> "" and oFSO.FileExists(oEnvironment.Item("DeployRoot") & sDir & oEnvironment.Item("Architecture") & "\" & sFileName) then
					sFoundPath = oEnvironment.Item("DeployRoot") & sDir & oEnvironment.Item("Architecture") & "\" & sFileName
					iRetVal = Success
					Exit For

				' oEnvironment.Item("Architecture") *can* be blank, search in the %Processor_Architecture% directory.
				ElseIf ucase(oEnv("Processor_Architecture")) = "AMD64" and oFSO.FileExists(oEnvironment.Item("DeployRoot") & sDir & "x64\" & sFileName) then
					sFoundPath = oEnvironment.Item("DeployRoot") & sDir & "x64\" & sFileName
					iRetVal = Success
					Exit For
				ElseIf oFSO.FileExists(oEnvironment.Item("DeployRoot") & sDir & oEnv("Processor_Architecture") & "\" & sFileName) then
					sFoundPath = oEnvironment.Item("DeployRoot") & sDir & oEnv("Processor_Architecture") & "\" & sFileName
					iRetVal = Success
					Exit For
				End if

			Next

		End if
		If sFoundPath <> "" then
			FindFile = iRetVal
			Exit Function
		End if


		' Check resource root locations (normally OSD only)

		If oEnvironment.Item("ResourceRoot") <> "" and oEnvironment.Item("ResourceRoot") <> oEnvironment.Item("DeployRoot") then

			For each sDir in Array("\", "\Servicing\", "\Tools\", "\USMT\", "\Templates\", "\Scripts\", "\Control\")

				If oFSO.FileExists(oEnvironment.Item("ResourceRoot") & sDir & sFileName) then
					sFoundPath = oEnvironment.Item("ResourceRoot") & sDir & sFileName
					iRetVal = Success
					Exit For
				ElseIf oFSO.FileExists(oEnvironment.Item("ResourceRoot") & sDir & oEnvironment.Item("Architecture") & "\" & sFileName) then
					sFoundPath = oEnvironment.Item("ResourceRoot") & sDir & oEnvironment.Item("Architecture") & "\" & sFileName
					iRetVal = Success
					Exit For
				End if

			Next

		End if
		If sFoundPath <> "" then
			FindFile = iRetVal
			Exit Function
		End if


		' Now look in the log directory, this directory, working directory, and SYSTEM32 directory

		If oFSO.FileExists(oUtility.LogPath & "\" & sFilename) then
			sFoundPath = oUtility.LogPath & "\" & sFilename
			iRetVal = Success
		ElseIf oFSO.FileExists(oUtility.ScriptDir & "\" & sFilename) then
			sFoundPath = oUtility.ScriptDir & "\" & sFilename
			iRetVal = Success
		ElseIf oFSO.FileExists(oUtility.LocalRootPath & "\" & sFileName) then
			sFoundPath = oUtility.LocalRootPath & "\" & sFileName
			iRetVal = Success
		ElseIf oFSO.FileExists(".\" & sFileName) then
			sFoundPath = oShell.CurrentDirectory & "\" & sFilename
			iRetVal = Success
		ElseIf oFSO.FileExists(oEnvironment.Substitute("%WINDIR%\SYSTEM32\" & sFilename)) then
			sFoundPath = oEnvironment.Substitute("%WINDIR%\SYSTEM32\" & sFilename)
			iRetVal = Success
		ElseIf oEnvironment.Item("OSDPACKAGEPATH") <> "" and oFSO.FileExists(oEnvironment.Item("OSDPACKAGEPATH") & "\" & sFilename) then
			sFoundPath = oEnvironment.Item("OSDPACKAGEPATH") & "\" & sFilename
			iRetVal = Success
		ElseIf oEnvironment.Item("_SMSTSPackagePath") <> "" and oFSO.FileExists(oEnvironment.Item("_SMSTSPackagePath") & "\" & sFilename) then
			sFoundPath = oEnvironment.Item("_SMSTSPackagePath") & "\" & sFilename
			iRetVal = Success
		ElseIf oFSO.FileExists(".\x86\" & sFileName) then
			sFoundPath = oShell.CurrentDirectory & "\x86\" & sFilename
			iRetVal = Success
		ElseIf oFSO.FileExists(".\x64\" & sFileName) then
			sFoundPath = oShell.CurrentDirectory & "\x64\" & sFilename
			iRetVal = Success
		ElseIf oFSO.FileExists(".\amd64\" & sFileName) then
			sFoundPath = oShell.CurrentDirectory & "\amd64\" & sFilename
			iRetVal = Success
		End if

		If iRetVal <> Success then
			oLogging.CreateEntry "FindFile: The file " & sFilename & " could not be found in any standard locations.", LogTypeInfo
		End if

		FindFile = iRetVal

	End Function


	Function FindMappedDrive(sServerUNC)

		Dim arrSplit
		Dim sServerShare, sServerName
		Dim arrDrives, i


		' If the UNC isn't a UNC, just return the drive letter

		If Len(sServerUNC) < 3 then
			FindMappedDrive = ""
			EXIT FUNCTION
		End if
		If Mid(sServerUNC,2,2) = ":\" then
			FindMappedDrive = Left(sServerUNC, 2)
			EXIT FUNCTION
		End if


		' Build the UNC

		If Instr(Mid(sServerUNC, 3), "\") <= 0 or Left(sServerUNC, 2) <> "\\" then
			FindMappedDrive = ""
			EXIT FUNCTION
		End if
		arrSplit = Split(Mid(sServerUNC,3), "\",2)
		sServerName = arrSplit(0)
		sServerShare = "\\" & sServerName & "\" & arrSplit(1)


		' Look to see if this is a mapped drive

		On Error Resume Next
		Set arrDrives = oNetwork.EnumNetworkDrives
		If Err then
			FindMappedDrive = ""
			Exit Function
		Else
			On Error Goto 0
			For i = 0 to arrDrives.Count - 1 Step 2
				If UCase(sServerShare) = UCase(arrDrives.Item(i+1)) and arrDrives.Item(i) <> "" then
					FindMappedDrive = arrDrives.Item(i)
					Exit Function
				End if
			Next
		End if
		On Error Goto 0

		FindMappedDrive = ""

	End Function


	Function ValidateConnection(sServerUNC)
		ValidateConnection = ValidateConnectionEx(sServerUNC, False)
	End function
	
	Function ValidateConnectionEx(sServerUNC, bForceConnection)

		Dim iRetVal
		Dim arrSplit
		Dim sServerShare, sServerName, sFoundDrive
		Dim sOSDConnectToUNC, sRIPInfo, sCmd
		Dim sWizardHTA, sUserID
		Dim sCurrentServerName, i

		' Make sure a UNC is specified

		If sServerUNC = "" then
			oLogging.CreateEntry "WARNING - Unable to validation connection because a blank UNC was specified.", LogTypeWarning
			ValidateConnectionEx = Failure
			EXIT FUNCTION
		End if
		If Mid(sServerUNC,2,2) = ":\" then
			oLogging.CreateEntry "Using a local or mapped drive, no connection is required.", LogTypeInfo
			ValidateConnectionEx = Success
			EXIT FUNCTION
		End if

		oLogging.CreateEntry "Validating connection to " & sServerUNC, LogTypeInfo
		iRetVal = ValidateNetworkConnectivity
		If iRetVal <> Success then

			'blow up
			ValidateConnectionEx = Failure
			Exit Function
		End IF


		' See if we've already connected

		If bForceConnection then
			' When forcing a connection, map all the way down to the specified folder
			arrSplit = Split(Mid(sServerUNC,3), "\", 2)
		Else
			' When not forcing, split into more chunks so we only get server and share
			arrSplit = Split(Mid(sServerUNC,3), "\")
		End if
		sServerName = arrSplit(0)
		sServerShare = "\\" & sServerName & "\" & arrSplit(1)


		' This isn't necessary if we're trying to connect to the current DP, so check that.

		If Left(oUtility.ScriptDir, 2) = "\\" then

			arrSplit = Split(Mid(oUtility.ScriptDir, 3), "\")
			sCurrentServerName = arrSplit(0)
			If UCase(sServerName) = UCase(sCurrentServerName) and not bForceConnection then
				oLogging.CreateEntry "Already connected to server " & sServerName & " as that is where this script is running from.", LogTypeInfo
				ValidateConnectionEx = Success
				EXIT FUNCTION
			End if

		End if



		' Now see if there is already a server connection.
		' We may wish to use bForceConnection if we *Require* the Drive letter, 
		'   even though the computer may be connected to the sever via another share.

		If dicNetworkConnections.Exists(sServerName) and not bForceConnection then
			oLogging.CreateEntry "Already connected to server " & sServerName, LogTypeInfo
			ValidateConnectionEx = Success
			EXIT FUNCTION
		End if


		'Map the SCCM user variables to the CS.INI variables

		If oEnvironment.Item("_SMSTSReserved1") <> "" AND oEnvironment.Item("_SMSTSReserved2") <> "" Then
			OEnvironment.Item("UserDomain") = Left(oEnvironment.Item("_SMSTSReserved1"),Instr(oEnvironment.Item("_SMSTSReserved1"),"\")-1)
			oEnvironment.Item("UserID")= Mid(oEnvironment.Item("_SMSTSReserved1"),Instr(oEnvironment.Item("_SMSTSReserved1"),"\")+1)
			oEnvironment.Item("UserPassword")=oEnvironment.Item("_SMSTSReserved2")
		End if


		'  It is possible that the server allows anonymous connections, skip if the share is readable.
		'if oFso.FolderExists(sServerName) then
		'	oLogging.CreateEntry "Already connected to server " & sServerName, LogTypeInfo
		'	ValidateConnectionEx = Success
		'	EXIT FUNCTION
		'end if



		' If no credentials are available, prompt.

		If oEnvironment.Item("UserID") = "" or oEnvironment.Item("UserPassword") = "" or (  oEnvironment.Item("UserDomain") = "" and _
			instr(1,oEnvironment.Item("UserID"),"\",vbTextCompare) = 0 and instr(1,oEnvironment.Item("UserID"),"@",vbTextCompare) = 0 ) then

			' Find the HTA that prompts for credentials

			iRetVal = FindFile("Wizard.hta", sWizardHTA)
			If iRetVal <> Success then
				oLogging.CreateEntry "ERROR - Unable to find Wizard.hta, so it is impossible to prompt for credentials.", LogTypeError
				ValidateConnectionEx = Failure
				Exit Function
			End if

			' Execute the HTA

			oShell.Run "mshta.exe """ & sWizardHTA & """ /NotWizard /LeaveShareOpen /ValidateAgainstUNCPath:""" & sServerShare & """ /Definition:Credentials_ENU.xml", 1, true

			' See if the values are populated now

			If oEnvironment.Item("UserID") = "" then
				oLogging.CreateEntry "ERROR - no credentials were returned from LTICredentials.hta, so no connection is possible.", LogTypeError
				ValidateConnectionEx = Failure
				Exit Function
			End if

		End if


		' Map a drive

		If oEnvironment.Item("UserDomain") <> "" then
			sUserID = oEnvironment.Item("UserDomain") & "\" & oEnvironment.Item("UserID")
		Else
			sUserID = oEnvironment.Item("UserID")
		End if

		For i = 1 to 5
			If MapNetworkDrive(sServerShare, sUserID, oEnvironment.Item("UserPassword")) = Success then
				Exit for
			End If

			oLogging.CreateEntry "Unable to connect to " & sServerShare & ".  Sleeping for " & CStr(i * 5) & " seconds.", LogTypeInfo
			On error Resume next
			WScript.Sleep i * 5000
			on error goto 0
		Next



		If MapNetworkDrive(sServerShare, sUserID, oEnvironment.Item("UserPassword")) <> Success Then
			oLogging.CreateEntry "ERROR - Unable to map a network drive to " & sServerShare & ".", LogTypeError
			ValidateConnectionEx = Failure
			Exit Function
		End if


		' Record the mapped drive

		oLogging.CreateEntry "Successfully established connection using supplied credentials.", LogTypeInfo
		If not dicNetworkConnections.Exists(sServerName) then
			dicNetworkConnections.Add sServerName, sServerShare
		End if

		iRetVal = Success


		ValidateConnectionEx = iRetVal

	End Function


	Function ValidateNetworkConnectivity

		Dim Entity, ID, oAdapter, oAdapter2, colAdapters, colAdapters2, sIPConnectionMetric, sWirelessConnectionMetric
		Dim bValidIP

		' Check for networkadapters present

		bValidIP = True
		If objWMI.ExecQuery("select * from win32_NetworkAdapter where Installed = true and adaptertypeid = 0").Count = 0 then
			oLogging.CreateEntry "No networking adapters found, The network drivers for your device are not present",LogTypeError
			ValidateNetworkConnectivity = Failure
			exit function
	
		End if


		'Check for IP address

		Set colAdapters = objWMI.ExecQuery("select * from win32_NetworkAdapterConfiguration where IPEnabled=True")

		For Each oAdapter in colAdapters
			If oAdapter.DHCPEnabled = TRUE Then
				If oAdapter.DHCPServer = "255.255.255.255" Then
					bValidIP = False
				
				Else
					bValidIP = True
					Exit For
				End if
				
			Else
				oLogging.CreateEntry "DHCP is not enabled, assuming static IP address", LogTypeInfo
				bValidIP = True
				Exit For
			End if
		Next

		If bValidIP <> True Then
			'No IP Address, do an ipconfig /renew
			oShell.Run "ipconfig /renew",0,true
			on error resume next
			wscript.sleep 5000
			on error goto 0
			Set colAdapters2 = objWMI.ExecQuery("select * from win32_NetworkAdapterconfiguration where IPEnabled = True")
			For Each oAdapter2 in colAdapters2
				If oAdapter2.DHCPEnabled = TRUE Then
					If oAdapter2.DHCPServer = "255.255.255.255" Then
						oLogging.CreateEntry "IP Address could not be obtained",LogTypeError
						ValidateNetworkConnectivity = Failure
						Exit Function
					Else
						Exit for
					End if
				End if
			Next
		End If

		'Check for wireless connectivity

		Set colAdapters = objWMI.ExecQuery("select * from win32_NetworkAdapterconfiguration where IPEnabled = True")
		For Each oAdapter in colAdapters
			If Instr(UCase(oAdapter.Caption),"WIRELESS") = 0 Then
				If oAdapter.IPConnectionMetric < sIPConnectionMetric Or sIPConnectionMetric = "" Then
					sIPConnectionMetric = oAdapter.IPConnectionMetric
				End If
			End IF
			
			If Instr(UCase(oAdapter.Caption),"WIRELESS") Then
				sWirelessConnectionMetric = oAdapter.IPConnectionMetric
			
		
			End If
		
		Next
		
		If sIPConnectionMetric = "" Then
			oLogging.CreateEntry "No physical adapters present, cannot deploy over wireless", LogTypeError
			ValidatenetworkConnectivity = Failure
			Exit Function
		End IF


		'TODO Check for VPN connectivity

		ValidateNetworkConnectivity = Success

	End Function


	' For Backwards Compatiblity

	Function MapNetworkDrive (sShare, sDomID, sDomPwd )
		If Len(MapNetworkDriveEx (sShare, sDomID, sDomPwd, LogTypeError )) = 2 then
			MapNetworkDrive = Success
		Else
			MapNetworkDrive = Failure
		End if
	End function


	'
	' maps a drive letter to the sShare UNC path.
	'   Returns the drive letter example: "C:", otherwise returns an error string!
	'   sDomID and sDomPwd can be EMPTY.
	'
	Function MapNetworkDriveEx (sShare, sDomID, sDomPwd, iLogType )

		Dim sDrive
		Dim HasError
		Dim ErrDesc
		Dim i
		Dim arrDrives


		' Make sure networking is initialized

		On Error Resume Next
		Set arrDrives = oNetwork.EnumNetworkDrives
		If Err then
			oLogging.CreateEntry "Unable to enumerate network drives (is the network initialized?): " & Err.Description & " (" & Err.Number & ")", LogTypeWarning
		Else

			' Find any previous connections (skip connections without drive letters)

			For i = 0 to arrDrives.Count - 1 Step 2
				If UCase(sShare) = UCase(oNetwork.EnumNetworkDrives.Item(i+1)) and oNetwork.EnumNetworkDrives.Item(i) <> "" then
					MapNetworkDriveEx = oNetwork.EnumNetworkDrives.Item(i)
					oLogging.CreateEntry "Found Existing UNC Path " & MapNetworkDriveEx & "  = " & sShare , LogTypeInfo

					Exit function
				End if
			Next

		End if
		On Error Goto 0


		' Find the first avaiable drive letter

		For sDrive = asc("Z") to asc("C") step -1

			On Error Resume Next
			oNetwork.MapNetworkDrive  chr(sDrive)&":", sShare, False, sDomID, sDomPwd
			HasError = err.number
			ErrDesc = err.Description
			On Error Goto 0

			Select case HasError
			Case 0          ' No Error, SUCCESS
				MapNetworkDriveEx = chr(sDrive)&":"
				oLogging.CreateEntry "Mapped Network UNC Path " & MapNetworkDriveEx & "  = " & sShare , LogTypeInfo
				Exit function
			Case &h80070055 ' The local device name is already in use.
			Case &h800704B2 ' The local device name has a remembered connection to another network resource.
			' Case &h800704C3 ' Multiple connections to a server or shared resource by the same user, using more than one user name, are not allowed.
			' Case &h8007052E ' Logon failure: unknown user name or bad password.			
			Case Else 
				oLogging.CreateEntry "Unable to connect to share: " & ErrDesc & "( 0x" & hex(HasError) & " ) , trying to connect without username. ", LogTypeInfo
				Err.Clear
				On Error Resume Next
				oNetwork.MapNetworkDrive chr(sDrive)&":", sShare, False
				HasError = err.number
				ErrDesc = err.Description
				On Error Goto 0
				If HasError <> 0 Then

					' There was a some kind of fatal error.
					If ErrDesc <> "" then
						MapNetworkDriveEx = ErrDesc
					Else
						MapNetworkDriveEx = "Unable to map UNC Path " & sShare & " :" & "( 0x" & hex(HasError) & " ) "
					End if
					oLogging.CreateEntry MapNetworkDriveEx & "", iLogType
					Exit function
				Else
					MapNetworkDriveEx = chr(sDrive)&":"
					Exit Function
				End If		
			End select

		Next

		MapNetworkDriveEx = "Unable to map UNC Path " & sShare & " : No avaiable local device names! "
		oLogging.CreateEntry MapNetworkDriveEx , iLogType

	End function


	Public Function VerifyPathExistsEx(strPath, bAdjustSD)
		Dim oHelper, sSDDL, oLogicalFile, oDescriptor
		If strPath = "" then
			VerifyPathExistsEx = True
			Exit Function
		End if
		If oFSO.FolderExists(strPath) then
			VerifyPathExistsEx = true
			Exit Function
		Else
			VerifyPathExists oFSO.GetParentFolderName(strPath)
			'On Error Resume Next
			oFSO.CreateFolder strPath
			
			'Set Permissions if it is not a UNC path
			If Left(strPath, 2) <> "\\" and bAdjustSD Then
				strPath = Replace(strPath, "\", "\\")

				Set oHelper = objWMI.Get("Win32_SecurityDescriptorHelper")

				sSDDL = "O:BAG:SYD:PAI(A;OICI;FA;;;BA)(A;OICI;FA;;;SY)"
 
				oHelper.SDDLToWin32SD sSDDL, oDescriptor
				Set oLogicalFile = objWMI.Get("Win32_LogicalFileSecuritySetting.Path='" & strPath & "'")
				TestAndLog oLogicalFile.SetSecurityDescriptor(oDescriptor), "Set security on folder " & strPath
				
			End If
			On Error Goto 0
		End if
	End function 
	
	Public Function VerifyPathExists(strPath)
		VerifyPathExists = VerifyPathExistsEx(strPath, false)
	End Function


	Function GetAllFixedDrives(bReturnOnlyBootable)

		Dim oPartition
		Dim oLogical
		Dim oDisk
		Dim o

		oLogging.CreateEntry "GetAllFixedDrives(" & bReturnOnlyBootable & ")", LogTypeInfo

		for each oDisk in objWMI.ExecQuery("SELECT DeviceID FROM Win32_DiskDrive WHERE InterfaceType <> ""USB"" and InterfaceType <> ""1394""")
			for each oPartition in objWMI.ExecQuery( "ASSOCIATORS OF {" & oDisk.Path_ & "}  WHERE AssocCLass = Win32_DiskDriveToDiskPartition" )
				for each oLogical in objWMI.ExecQuery( "ASSOCIATORS OF {" & oPartition.Path_ & "}  WHERE AssocCLass = Win32_LogicalDiskToPartition" )
					If bReturnOnlyBootable then
						If oLogical.DriveType <> 3 then
							exit for
						ElseIf not oPartition.Bootable then
							exit for
						ElseIf not oFSO.FileExists( oLogical.DeviceID & "\ntldr" ) and not oFSO.FileExists( oLogical.DeviceID & "\bootmgr" ) and not oFSO.FileExists( oLogical.DeviceID & "\bootmgr.efi" ) then
							exit for
						End if
					End if 
					GetAllFixedDrives = GetAllFixedDrives & " " & oLogical.DeviceID
				next
			next
		next

		GetAllFixedDrives = split(trim(GetAllFixedDrives)," ")

	End function

	'
	' Create an XMLDOM Object
	'
	Function CreateXMLDOMObjectEx( sFileName )
	
		Dim bRetVal

		on error resume next
		
		Set CreateXMLDOMObjectEx = nothing
		Set CreateXMLDOMObjectEx = CreateObject("MSXML2.DOMDocument")
		TestAndFail not (CreateXMLDOMObjectEx is nothing), 5490, "Create MSXML2.DOMDocument."
		
		CreateXMLDOMObjectEx.Async = FALSE

		If sFileName <> "" then
		
			oLogging.CreateEntry "CreateXMLDOMObjectEx(" & sFileName & ")", LogTypeVerbose
			bRetVal = CreateXMLDOMObjectEx.Load (sFileName )
			TestAndLog bRetVal, "CreateXMLDOMObjectEx...Length = " & CreateXMLDOMObjectEx.documentElement.childNodes.Length
			With CreateXMLDOMObjectEx.ParseError
				If .errorCode <> 0 then
					oLogging.CreateEntry "File: " & sFileName & " Line: " & .Line & " - " & .Reason & " - " & .SrcText, LogTypeWarning
				End if
			End with
			
		End if
		
		on error goto 0

	End function
	

	' Create an XMLDOM Object, do not return if there is a parse error.
	Function CreateXMLDOMObjectSafe( sFileName )
	
		Set CreateXMLDOMObjectSafe = CreateXMLDOMObjectEx( sFileName )
		TestAndFail CreateXMLDOMObjectSafe.ParseError.ErrorCode, 5495, "Create MSXML2.DOMDocument  .ParseErr.ErrCode."
		
	End function
	

	
	Function CreateXMLDOMObject
		Set CreateXMLDOMObject = CreateXMLDOMObjectEx( empty )
	End function


	' Load an XML file via FindFile
	Function LoadConfigFile( sConfigFile )

		Set LoadConfigFile = LoadConfigFileEx( sConfigFile, false )

	End function
	
	
	' Load an XML file via FindFile, do not return if there is any error.
	Function LoadConfigFileSafe( sConfigFile )
	
		Set LoadConfigFileSafe = LoadConfigFileEx( sConfigFile, TRUE )

	End function
	

	' Load an XML file via FindFile, do not return if there is any error.
	Function LoadConfigFileEx( sConfigFile, bMustSucceed )
	
		Dim sFoundFile
		Dim iRetVal

		iRetVal = oUtility.FindFile( sConfigFile , sFoundFile)
		If bMustSucceed then
			TestAndFail iRetVal, 5496, "LoadControlFile.FindFile: " & sConfigFile
			set LoadConfigFileEx = CreateXMLDOMObjectSafe (sFoundFile)
		Else
			set LoadConfigFileEx = CreateXMLDOMObjectEx (sFoundFile)
		End if

	End function


	Function BDDUtility

		Dim sBDDUtility
		Dim iRetval
		Dim sProc


		' Already retrieved an instance?  Return it.
		If not (oBDDUtility is Nothing) then
			Set BDDUtility = oBDDUtility
			Exit Function
		End if


		' Already registered? Call Routine.
		on error resume next
		Set oBDDUtility = CreateObject("Microsoft.BDD.Utility")
		on error goto 0

		If not (oBDDUtility is Nothing) then
			Set BDDUtility = oBDDUtility
			Exit Function
		End if


		' Find each Microsoft.BDD.Utility.dll file, and copy locally if over the network.
		for each sProc in array("x86","x64")

			If not oFSO.FileExists(oEnv("TEMP") & "\Tools\" & sProc & "\Microsoft.BDD.Utility.dll") then

				iRetVal = FindFile( sProc & "\Microsoft.BDD.Utility.dll", sBDDUtility)
				oLogging.CreateEntry "FindFile(...\Microsoft.BDD.Utility.dll)  Result : " & iRetVal, LogTypeInfo

				If left(sBDDUtility,2) = "\\" then
					oUtility.VerifyPathExistsEx oEnv("TEMP") & "\Tools\" & sProc , false
					oFileHandling.CopyFile sBDDUtility, oEnv("TEMP") & "\Tools\" & sProc & "\", True
				End if

			End if
			
			If oFSO.FileExists(oEnv("TEMP") & "\Tools\" & sProc & "\Microsoft.BDD.Utility.dll") then
				sBDDUtility = oEnv("TEMP") & "\Tools\" & sProc & "\Microsoft.BDD.Utility.dll"
			End if

			' Register the DLL
			oLogging.CreateEntry "RUN: regsvr32.exe /s """ & sBDDUtility & """", LogTypeInfo
			oShell.Run "regsvr32.exe /s """ & sBDDUtility & """", 0, true ' Always returns 0 - Success

		next



		' Create an instance

		on error resume next
		Set oBDDUtility = CreateObject("Microsoft.BDD.Utility")
		TestAndLog not oBDDUtility is nothing, "CreateObject(Microsoft.BDD.Utility)"
		on error goto 0
		
		Set BDDUtility = oBDDUtility

	End Function


	Sub SetTaskSequenceProperties(tsID)

		Dim oTaskSequences
		Dim oTaskSequence
		Dim oOperatingSystems
		Dim oOS
		Dim oImageLang
		Dim sImageLang
		Dim oLanguage
		Dim sImagePath
		Dim oWDSServer
		Dim sWDSServer
		Dim objTmp
		Dim oTS
		Dim oOSGUID
		Dim oDestinationDisk,oDestinationPartition,ODestinationLogicalDrive
		Dim sSourcePath


		' If there is task sequence ID set, get the properties

		If tsID <> "" then

			' Get the build record
			tsID = Ucase(tsID)

			Set oTaskSequences = oUtility.CreateXMLDOMObjectEx(oEnvironment.Item("DeployRoot") & "\Control\TaskSequences.xml")
			Set oTaskSequence = oTaskSequences.selectSingleNode("//ts[ID='" & tsID & "']")
			If oTaskSequence is Nothing then
				oLogging.CreateEntry "ERROR: Invalid task sequence ID " & tsID & " specified", LogTypeError
				Exit Sub
			End if


			' Set the simple build properties

			oEnvironment.Item("TaskSequenceName") = oUtility.SelectSingleNodeString(oTaskSequence,"Name")
			on error resume next
			oEnvironment.Item("TaskSequenceVersion") = oTaskSequence.SelectSingleNode("Version").text
			on error goto 0

			' Load the TS.XML and get the OSGUID

			Set oTS = oUtility.CreateXMLDOMObjectEx(oEnvironment.Item("DeployRoot") & "\Control\" & tsID & "\TS.xml")
			Set oOSGUID = oTS.SelectSingleNode("//globalVarList/variable[@name='OSGUID']")
			If oOSGUID is Nothing then

				Exit Sub
			End if
			oEnvironment.Item("OSGUID")=oOSGUID.text
			If oEnvironment.Item("DestinationDisk") = "" and oEnvironment.Item("DestinationPartition") = "" Then

				on error resume next
				' Get the Disk Destination Variables
				Set oDestinationDisk = oTS.SelectSingleNode("//globalVarList/variable[@name='DestinationDisk']")
				Set oDestinationPartition = oTS.SelectSingleNode("//globalVarList/variable[@name='DestinationPartition']")
				If not oDestinationDisk is Nothing then
					oProperties.Add "DefaultDestinationDisk", oDestinationDisk.text
				End if
				If not oDestinationPartition is Nothing then
					oProperties.Add "DefaultDestinationPartition", oDestinationPartition.text
				End if
				on error goto 0
			End If
			
			' Get the OS record

			Set oOperatingSystems = oUtility.CreateXMLDOMObjectEx(oEnvironment.Item("DeployRoot") & "\Control\OperatingSystems.xml")
			Set oOS = oOperatingSystems.selectSingleNode("//os[@guid='" & oOSGUID.text & "']")
			If oOS is Nothing then
				oLogging.CreateEntry "ERROR: Invalid OS GUID " & oOSGUID.text & " specified for task sequence " & tsID & " specified", LogTypeInfo
				Exit Sub
			End if


			' Set the simple OS properties

			oEnvironment.Item("ImageIndex") = oUtility.SelectSingleNodeString(oOS,"ImageIndex")
			oEnvironment.Item("ImageSize") = oUtility.SelectSingleNodeString(oOS,"Size")
			on error resume next
			oEnvironment.Item("ImageFlags") = oOS.selectSingleNode("Flags").text
			on error goto 0
			oEnvironment.Item("ImageBuild") = oUtility.SelectSingleNodeString(oOS,"Build")
			oEnvironment.Item("ImageProcessor") = oUtility.SelectSingleNodeString(oOS,"Platform")


			' Get the languages

			Set oImageLang = oOS.selectNodes("Language")
			sImageLang = ""
			If not (oImageLang is Nothing) then
				For each oLanguage in oImageLang
					sImageLang = sImageLang & oLanguage.text & vbTab
				Next
			End if
			If right(sImageLang,1) = vbTab then
				sImageLang = Left(sImageLang, Len(sImageLang)-1)  ' Remove trailing tab
			End if
			oEnvironment.ListItem("ImageLanguage") = split(sImageLang, vbTab)


			' Set the image path

			If oOS.selectSingleNode("ImageFile") is nothing then
				sImagePath = "."
			Else
				sImagePath = oOS.selectSingleNode("ImageFile").Text
				If sImagePath = "" and left(oEnvironment.Item("ImageBuild"),1) = "5" then
					sImagePath = "."
				End if
			End if


			If Left(sImagePath, 1) = "." then

				' See if this is a WDS image

				Set oWDSServer = oOS.selectSingleNode("WDSServer")
				If not (oWDSServer is Nothing) then
					sWDSServer = oWDSServer.Text
				End if


				' Make sure that's where we want to pull it from

				If sWDSServer <> "" then
					If oEnvironment.Item("WDSServer") <> "" then
						sWDSServer = oEnvironment.Item("WDSServer")
					End if
				End if


				' Set the actual image path

				If sWDSServer <> "" then
					sImagePath = "\\" & sWDSServer & "\REMINST" & Mid(sImagePath, 2)
				Else
					sImagePath = oEnvironment.Item("DeployRoot") & Mid(sImagePath, 2)
				End if

			End if

			oLogging.CreateEntry "InstallFromPath: " & sImagePath, LogTypeInfo
			oEnvironment.Item("InstallFromPath") = oFileHandling.NormalizePath(sImagePath)

			sSourcePath = oUtility.SelectSingleNodeString(oOS,"Source")
			If Left(sSourcePath, 1) = "." then
				sSourcePath = oEnvironment.Item("DeployRoot") & Mid(sSourcePath, 2)
			End if
			oLogging.CreateEntry "SourcePath: " & sSourcePath, LogTypeInfo
			oEnvironment.Item("SourcePath") = sSourcePath

		End if

	End Sub


	Function IsSupportedPlatform(sPlatform)

		Dim iRetVal
		Dim sSupportedPlatforms
		Dim oPlatformNode
		Dim oNode
		Dim oResults
		Dim oResult
		Dim bFound


		' Special case: check for Windows PE

		If oEnvironment.Item("OSVersion") = "WinPE" then

			If sPlatform = "Windows PE" then
				IsSupportedPlatform = true
			Else
				IsSupportedPlatform = false
			End if

			Exit Function
		End if


		' Load the XML file if not yet loaded


		If oSupportedPlatforms is Nothing then

			iRetVal = oUtility.FindFile("ZTISupportedPlatforms.xml", sSupportedPlatforms)
			Set oSupportedPlatforms = CreateXMLDOMObjectEx(sSupportedPlatforms)

		End if


		' Find the selected platform

		Set oPlatformNode = oSupportedPlatforms.SelectSingleNode("//SupportedPlatform[@name='" & sPlatform & "']")
		If oPlatformNode is Nothing then
			oLogging.CreateEntry "Platform " & sPlatform & " is not found.", LogTypeInfo
			IsSupportedPlatform = False
			Exit Function
		End if


		' Check each of the expressions.  If any don't return a record, return false

		For each oNode in oPlatformNode.SelectNodes("Expression")

			bFound = false
			Set oResults = objWMI.ExecQuery(oNode.Text)
			For each oResult in oResults
				bFound = true
				Exit For
			Next

			If not bFound then
				oLogging.CreateEntry "Condition " & oNode.Text & " not satisfied, platform " & sPlatform & " is not supported.", LogTypeInfo
				IsSupportedPlatform = False
				Exit Function
			End if

		Next


		' All conditions satisfied, return true

		oLogging.CreateEntry "Platform " & sPlatform & " is supported on this computer.", LogTypeInfo
		IsSupportedPlatform = True

	End Function


	Function DetermineDriveFromDiskPart( sDestDisk, sDestPart )
	
		Dim oDiskPartition, oAssocs, Assoc

		' Determine if the Disk and Partition exists
		oLogging.CreateEntry "Verifying that the Disk and Partition exists",LogTypeInfo
		On Error Resume Next
		Set oDiskPartition = objWMI.ExecQuery("Select * From Win32_DiskPartition where deviceid = 'Disk #" & sDestDisk & ", Partition #" & sDestPart-1 & "'")
		On Error Goto 0
		If oDiskPartition.Count = 0 Then
			oLogging.CreateEntry "The specified Disk and Partition does not exist",logtypeInfo
			DetermineDriveFromDiskPart = Failure
		Else
			oLogging.CreateEntry "Found Disk and Partition",logTypeInfo
			On Error Resume Next
			Set oAssocs = objWMI.ExecQuery("Associators of {Win32_DiskPartition.DeviceID='Disk #" & sDestDisk & ", Partition #" & sDestPart-1 & "'} where ResultClass = Win32_LogicalDisk")
			On Error Goto 0 
			If oAssocs.Count = 0 Then
				oLogging.CreateEntry "The logical drive could not be determined from the disk and partition", logtypeInfo
				DetermineDriveFromDiskPart = Failure
			Else
				For Each Assoc in oAssocs
					DetermineDriveFromDiskPart = Assoc.DeviceID
					oLogging.CreateEntry "Logical Drive was determined to be " & DetermineDriveFromDiskPart, logTypeInfo
				Next
			End if

		End if

	End function
	

	Function DeterminePartition
		Dim iRetVal
		Dim sFile
		Dim oDiskPartition, oAssocs, DeviceID, Assoc

		DeterminePartition = Success


		If oEnvironment.Item("DestinationLogicalDrive") <> "" Then
			'Determine the Disk and Partition from the drive letter specified
			oLogging.CreateEntry "Determining the Disk and Partition Number from the Logical Drive", logTypeInfo
			On Error Resume Next
			Set oAssocs = objWMI.ExecQuery("ASSOCIATORS OF {Win32_LogicalDisk.DeviceID='" & oEnvironment.Item("DestinationLogicalDrive") & "'} WHERE ResultClass = Win32_DiskPartition")
			On Error Goto 0
			if oAssocs.Count = 0  Then
				oLogging.CreateEntry "The specified LogicalDrive does not exist", logtypeError
				DeterminePartition = Failure
				Exit Function
			Else
				For each Assoc in oAssocs
					DeviceID = Assoc.DeviceID
					oEnvironment.Item("DestinationDisk") = Mid(DeviceID, Instr(DeviceID,"#")+1,1)
					oEnvironment.Item("DestinationPartition") = Right(DeviceID,1)+1
					oLogging.CreateEntry "The logical drive was found on Disk " & oEnvironment.Item("DestinationDisk") & " and Partition " & oEnvironment.Item("DestinationPartition"), logTypeInfo
				Next
			End if

		End if

		If oEnvironment.Item("DestinationDisk") <> "" AND oEnvironment.Item("DestinationPartition") <> "" Then
			' Determine if the Disk and Partition exists
			oLogging.CreateEntry "Verifying that the Disk and Partition exists",LogTypeInfo
			On Error Resume Next
			Set oDiskPartition = objWMI.ExecQuery("Select * From Win32_DiskPartition where deviceid = 'Disk #" & oEnvironment.Item("DestinationDisk") & ", Partition #" & oEnvironment.Item("DestinationPartition")-1 & "'")
			On Error Goto 0
			If oDiskPartition.Count = 0 Then
				oLogging.CreateEntry "The specified Disk and Partition does not exist",logtypeError
				DeterminePartition = Failure
				Exit Function
			Else
				oLogging.CreateEntry "Found Disk and Partition",logTypeInfo
				If oEnvironment.Item("DestinationLogicalDrive") = "" Then
					On Error Resume Next
					Set oAssocs = objWMI.ExecQuery("Associators of {Win32_DiskPartition.DeviceID='Disk #" & oEnvironment.Item("DestinationDisk") & ", Partition #" & oEnvironment.Item("DestinationPartition")-1 & "'} where ResultClass = Win32_LogicalDisk")
					On Error Goto 0 
					If oAssocs.Count = 0 Then
						If Left(oEnvironment.Item("OSCurrentVersion"),3) >= "6.0" and oENV("SystemDrive") <> "X:" Then
							'There might be a hidden partition 
							oEnvironment.Item("DestinationDisk") = ""
							oEnvironment.Item("DestinationPartition") = ""
							oEnvironment.Item("DestinationLogicalDrive") = oENV("SystemDrive")
							DeterminePartition
						Else
							oLogging.CreateEntry "The logical drive could not be determined from the disk and partition", logTypeError
							DeterminePartition = Failure
							Exit Function
						End If
					Else
						For Each Assoc in oAssocs
							oEnvironment.Item("DestinationLogicalDrive") = Assoc.DeviceID
							oLogging.CreateEntry "Logical Drive was determined to be " & oEnvironment.Item("DestinationLogicalDrive"), logTypeInfo
						Next
					End if
				End if

			End if
		End if

		If oEnvironment.Item("DestinationLogicalDrive") = ""  Then
			oEnvironment.Item("DestinationDisk") = 0
			oEnvironment.Item("DestinationPartition") = 1
			oEnvironment.Item("DestinationLogicalDrive") = "C:"
		End If


	End Function


	Function SelectSingleNodeString( oXMLDomNode, sXPath )

		' TestAndLog not oXMLDomNode is nothing, "verify oXMLDomNode is object."
		If oXMLDomNode.SelectSingleNode(sXPath) is nothing then
			oLogging.CreateEntry "SelectSingleNodeString(" & sXPath & ") Missing Node.", LogTypeWarning
			SelectSingleNodeString = ""
		Else
			SelectSingleNodeString = oXMLDomNode.SelectSingleNode(sXPath).Text
		End if

	End function

	Function FindSysprepAnswerFile
		
		Dim sSysprepInf
		Dim sBuildPath
		Dim iRetVal
		
		iRetVal = SUCCESS
		If oEnvironment.Item("TaskSequenceID") = "" Then
			oLogging.CreateEntry "The TaskSequenceID is blank, possibly a misconfigured customsettings.ini or task sequence",LogTypeWarning
		End If
		sBuildPath = oEnvironment.Item("DeployRoot") & "\Control\" & oEnvironment.Item("TaskSequenceID")
		If not oFSO.FolderExists(sBuildPath) then
			sBuildPath = oEnvironment.Item("DeployRoot")

		End if

		oLogging.CreateEntry "Looking for Sysprep.inf in " & oEnvironment.Item("DestinationLogicalDrive") & "\sysprep\Sysprep.inf", LogTypeInfo
		
		If oFSO.FileExists(oEnvironment.Item("OSDAnswerFilePathSysprep")) Then
				
			FindSysprepAnswerFile = iRetVal
			Exit Function
		
		ElseIf oFSO.FileExists(oEnvironment.Item("DestinationLogicalDrive") & "\sysprep\Sysprep.inf") then

			sSysprepInf = oEnvironment.Item("DestinationLogicalDrive") & "\sysprep\Sysprep.inf"
			oLogging.CreateEntry "Found Sysprep.inf at " & sSysprepInf & ".", LogTypeInfo

		ElseIf oFSO.FileExists("x:\sysprep\Sysprep.inf") then

			sSysprepInf = "x:\sysprep\Sysprep.inf"
			oLogging.CreateEntry "Found Sysprep.inf at " & sSysprepInf & ".", LogTypeInfo

		ElseIf oFSO.FileExists(sBuildPath & "\Sysprep.inf") then

			' Copy it locally

			sSysprepInf = oEnvironment.Item("DestinationLogicalDrive") & "\sysprep\Sysprep.inf"
			oLogging.CreateEntry "Found Sysprep.inf at " & sBuildPath & "\Sysprep.inf, will copy to " & sSysprepInf, LogTypeInfo

			If not oFSO.FolderExists(oEnvironment.Item("DestinationLogicalDrive") & "\Sysprep") then
				oFSO.CreateFolder oEnvironment.Item("DestinationLogicalDrive") & "\Sysprep"
			End if

			oFSO.CopyFile sBuildPath & "\Sysprep.inf", sSysprepInf
			oFSO.GetFile(sSysprepInf).Attributes = 0
			oLogging.CreateEntry "Copied " & sBuildPath & "\Sysprep.inf to " & oEnvironment.Item("DestinationLogicalDrive") & "\sysprep", LogTypeInfo
		ElseIf oEnvironment.Item("OSDTargetSystemDrive")<> "" and oFSO.FileExists(oEnvironment.Item("OSDTargetSystemDrive") & "\sysprep\sysprep.inf") then
			
			sSysprepInf = oEnvironment.Item("OSDTargetSystemDrive") & "\sysprep\sysprep.inf"
			oLogging.CreateEntry "Found Sysprep.inf at " & sSysprepInf & ".", LogTypeInfo

		Else
			oLogging.CreateEntry "The sysprep.inf file was not found.", LogTypeInfo
		End if
		
		oEnvironment.Item("OSDAnswerFilePathSysprep") = sSysprepInf
		
		FindSysprepAnswerFile = iRetVal
			

	
	End Function
	
	Function FindUnattendAnswerFile
		
		Dim sUnattendXML, sUnattendTxt
		Dim sBuildPath
		Dim iRetVal
		
		iRetVal = SUCCESS
		If oEnvironment.Item("TaskSequenceID") = "" Then
			oLogging.CreateEntry "The TaskSequenceID is blank, possibly a misconfigured customsettings.ini or task sequence",LogTypeWarning
		End If

		sBuildPath = oEnvironment.Item("DeployRoot") & "\Control\" & oEnvironment.Item("TaskSequenceID")
		If not oFSO.FolderExists(sBuildPath) then
			sBuildPath = oEnvironment.Item("DeployRoot")

		End if

		If oFSO.FileExists(sBuildPath & "\Unattend.txt") then

			' Copy it locally

			sUnattendTxt = oUtility.LocalRootPath & "\unattend.txt"
			oLogging.CreateEntry "Found Unattend.txt at " & sBuildPath & "\Unattend.txt, will copy to " & sUnattendTxt, LogTypeInfo

			oFSO.CopyFile sBuildPath & "\Unattend.txt", sUnattendTxt, true
			oFSO.GetFile(sUnattendTxt).Attributes = 0
			oLogging.CreateEntry "Copied " & sBuildPath & "\Unattend.txt to " & sUnattendTxt, LogTypeInfo

		ElseIf oFso.FileExists(oEnvironment.Item("OSDAnswerFilePath")) and Instr(1,oEnvironment.Item("OSDAnswerFilePath"),".txt",vbTextCompare) >0 then
			sUnattendTxt= oEnvironment.Item("OSDAnswerFilePath")
			oLogging.CreateEntry "Found unattend.txt at " & oEnvironment.GetOSDV4("_SMSTSMDataPath") & "\unattend.txt",LogTypeInfo
		Else
			oLogging.CreateEntry "The unattend.txt file was not found.", LogTypeInfo
		End if


		' First see if there is already a local unattend.xml.  If not, copy one.

		
		If oFSO.FileExists(Left(oUtility.LocalRootPath, 2) & "\Windows\Panther\unattend\unattend.xml") then

			
			sUnattendXml = Left(oUtility.LocalRootPath, 2) & "\Windows\Panther\unattend\unattend.xml"
			oLogging.CreateEntry "Found existing unattend.xml at " & sUnattendXml, LogTypeInfo

		ElseIf oFSO.FileExists(oUtility.LocalRootPath & "\unattend.xml") then

			
			sUnattendXml = oUtility.LocalRootPath & "\unattend.xml"
			oLogging.CreateEntry "Found existing unattend.xml at " & sUnattendXml, LogTypeInfo

		ElseIf oFSO.FileExists(sBuildPath & "\Unattend.xml") then

			
			sUnattendXml = oUtility.LocalRootPath & "\Unattend.xml"

			oLogging.CreateEntry "Found unattend.xml at " & sBuildPath & "\Unattend.xml, will copy to " & sUnattendXml, LogTypeInfo
			oFSO.CopyFile sBuildPath & "\Unattend.xml", sUnattendXml, true
			oLogging.CreateEntry "Copied " & sBuildPath & "\Unattend.xml to " & sUnattendXml, LogTypeInfo
			oFSO.GetFile(sUnattendXml).Attributes = 0
		ElseIf oFso.FileExists(oEnvironment.Item("OSDAnswerFilePath")) and Instr(1,oEnvironment.Item("OSDAnswerFilePath"),".xml",vbTextCompare) >0 then
			sUnattendXML= oEnvironment.Item("OSDAnswerFilePath")



			oLogging.CreateEntry "Found existing unattend.xml at " & sUnattendXml, LogTypeInfo

		Else

			sUnattendXml = ""
			oLogging.CreateEntry "File " & sBuildPath & "\Unattend.xml does not exist, unable to copy", LogTypeInfo

		End if
		
		If sUnattendTxt <> "" AND sUnattendXML <> "" Then
			oLogging.CreateEntry "Found an unattend.xml and an unattend.txt file, Invalid configuration", LogTypeError
			iRetVal = FAILURE
			FindUnattendAnswerFile = iRetVal
			Exit Function
		ElseIf sUnattendTxt <> "" AND sUnattendXML = "" Then
			oEnvironment.Item("OSDAnswerFilePath") = sUnattendTxt
		ElseIF sUnattendXML <> "" AND sUnattendTxt = "" Then
			oEnvironment.Item("OSDAnswerFilePath") = sUnattendXML
		Else
			oLogging.CreateEntry "No answer file could be found",LogTypeWarning
		End If
		
		FindUnattendAnswerFile = iRetVal
		
	
	End Function
	
	Function IsHighEndSKUEx( sSKU )
	
		' Windows Ultimate/Enterprise and Server SKU's allow for some
		' higher-end features, like Bitlocker and Multiple Language Packs.
		
		select case (ucase(trim(sSKU)))
			case "ULTIMATE", "ULTIMATEE", "ULTIMATEN"
				IsHighEndSKUEx = TRUE
			case "ENTERPRISE", "ENTERPRISEE", "ENTERPRISEN"
				IsHighEndSKUEx = TRUE
			case "HYPERV"
				IsHighEndSKUEx = TRUE
			case else
				If Instr(1, ucase(trim(sSKU)), "SERVER", vbTextCompare) > 0 then
					IsHighEndSKUEx = TRUE
				Else
					IsHighEndSKUEx = FALSE
				End if
		End Select
		
	End function
	
	Function IsHighEndSKU
		TestAndLog oEnvironment.Item("OSSKU") <> "", "Verify %OSSKU% is defined."
		IsHighEndSKU = IsHighEndSKUEx( oEnvironment.Item("OSSKU") )
	End function 


End Class


'
'  Common String Processing Routines
'
Class Strings


	' Create a delimited list of items.
	Sub AddToList(byref List, Item, Delimiter)  
		if isempty(list) then 
			List = cstr(item)
		else 
			list = list & delimiter & cstr(Item)
		end if 
	end Sub 

	' Display a Hex value with width
	Function HexWidth ( Value, Width  ) 
		HexWidth = right( "00000000" & hex ( value  ), Width)
	end Function 

	' Display a Hex value with width
	Function HexWidthByte ( Value, Width  ) 
		HexWidthByte = right( "00000000" & hex ( ascb( value  ) ), Width ) 
	end Function 

	Function IsWhiteSpace (MyChar)
		' Whitespace defined as vtTab[9], vbLF[10], vbVerticalTab[11], vbFormFeed[12], vbCr[13]
		IsWhiteSpace = MyChar = " " or ( MyChar >= chr(9) and MyChar <= chr(13) ) or MyChar = chr(160)
		end Function 

	Function TrimAllWS( MyString )
		TrimAllWS = MyString
		While len(TrimAllWS) > 0 and IsWhiteSpace(left(TrimAllWS,1))
			TrimAllWS = Mid(TrimAllWS,2)
		wend
		While len(TrimAllWS) > 0 and IsWhiteSpace(right(TrimAllWS,1)) 
			TrimAllWS = Mid(TrimAllWS,1,len(TrimAllWS)-1)
		wend   
	end Function

	Function RightAlign( MyString, Width ) 
		RightAlign = Right( Space(Width) & MyString, Width )
	end Function

	Function LeftAlign( MyString, Width ) 
		LeftAlign = Left( MyString & Space(Width), Width )
	end Function

	'
	' Force a value to a string format.
	'   Non-printable types will return empty.
	'
	Function ForceAsString ( InputVar )
		dim InputType, Item
		
		InputType = VarType(InputVar)

		if isObject(InputVar) or isNull(InputVar) or isEmpty(InputVar) then
			ForceAsString = "" 
		
		elseif InputType = vbError or InputType = vbVariant or InputType = vbDataObject then
			ForceAsString = "" 
			
		elseif InputType = ( vbArray or vbByte ) then
			for item = 1 to LenB( InputVar )
				AddToList ForceAsString, ForceAsString( HexWidthByte(midb(InputVar,Item,1),2) ), ""
			next 

		elseif isArray(InputVar) then
			for each item in InputVar        
				AddToList ForceAsString, ForceAsString(Item), " "   ' recurse
			next    
			
		elseif InputType = vbByte then
			ForceAsString = HexWidthByte(InputVar,2)
			
		else
			ForceAsString = cstr(InputVar)
			
		end if
	end Function 

	'
	' Given a variant, will ensure the Function returns an array
	'
	' Parameters:
	'   InputVar - Variable to convert to an array
	'   sDelimiter - OPTIONAL delimiter to force spliting in the array
	'       If sDelimiter is empty or blank, Function will return an array of 1 element.
	' Notes

	Function ForceAsArray( InputVar, sDelimiter )
		dim i

		if isArray(InputVar) then        
			ForceAsArray = InputVar 
			
		elseif VarType(InputVar) = vbObject then
			redim newarray(InputVar.Count-1)
			for i = 0 to InputVar.Count-1
				newarray(i) = InputVar(i)         
			next
			ForceAsArray = NewArray
			
		elseif VarType(InputVar) = vbString and VarType(sDelimiter) = vbString then
			ForceAsArray = split(InputVar,sDelimiter, -1, vbTextCompare )
			
		else
			ForceAsArray = array(InputVar)
			
		end if

	end Function
	
	
	Function GenerateRandomGUID 
		GenerateRandomGUID = left( CreateObject("Scriptlet.TypeLib").GUID, 38 ) 
	end Function    
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'
	'  Base64 Functions
	'

	Function BASE64_TABLE
		BASE64_TABLE = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	End function 

	Function SafeAsc ( sPlainText, pos )
		if VarType(sPlainText) = (vbArray or vbByte) then
			SafeAsc =  cint(midb(sPlainText, pos + 1 , 1))
		elseif (pos + 1) <= len(sPlainText) then
			SafeAsc = asc(mid(sPlainText, pos + 1, 1))
		else
			SafeAsc = 0
		end if
	End Function

	Function SafeEnc( n, x )
		SafeEnc = mid(BASE64_TABLE, ((n\(2^x)) and 63) + 1,1)
	End Function

	Function base64Encode( sPlainText )
		Dim i, n

		if 0 < len(sPlainText) then
			for i = 0 to len(sPlainText) - 1 step 3
				' Add a new line ...
				if i > 0  and i mod 57  = 0 then
					base64Encode = base64Encode & vbNewLine
				end if
				' three 8-bit characters become one 24-bit number
				n = (SafeAsc(sPlainText,i)*&h10000 + SafeAsc(sPlainText,i+1)*&h100 + SafeAsc(sPlainText,i+2))

				' the 24-bit number becomes four 6-bit numbers
				base64Encode = base64Encode & SafeEnc( n, 18 ) & SafeEnc( n, 12 ) & SafeEnc( n, 6 ) & SafeEnc( n, 0 )
			next
		end if

		' Pad Text at End of String
		n = (3-(len(sPlainText)mod 3)) mod 3
		base64Encode = left ( base64Encode, len(base64Encode) - n ) + string( n, "=" )

	End Function

	Function SafeDecode( s, i, x )
		SafeDecode = ( InStr(1, BASE64_TABLE, mid(s,i,1), vbBinaryCompare) - 1) * (2 ^ x)
	End Function

	Function base64Decode( sEncodedText )

		Dim sEncText
		Dim regex
		Dim p, i, n

		' Remove all non base64 text
		set regex = new RegExp


		regex.pattern = "[^=" & BASE64_TABLE & "]"
		regex.global = true

		sEncText = regex.Replace(sEncodedText,"")
		sEncText = replace( sEncText, vbLF, "")
		sEncText = replace( sEncText, vbCR, "")

		' Verify String is in Base64 format (multiple of 4 chars)
		if len(sEncText) mod 4 <> 0 then
			oLogging.CreateEntry "Variable is not a valid string (not Base64 Format)", LogTypeInfo
			base64Decode = ""			
			exit function
		end if

		if right(sEncText,2) = "==" then
			p = 2
		elseif right(sEncText,1) = "=" then
			p = 1
		end if
		sEncText = left(sEncText,len(sEncText)-p) & string(p,"A")

		for i = 1 to len(sEncText) step 4
			' Convert four 6-bit numbers into one 24 bit value
			n = SafeDecode(sEncText,i+3,0) + SafeDecode(sEncText,i+2,6) + SafeDecode(sEncText,i+1,12) + SafeDecode(sEncText,i+0,18)

			' Convert the 24-bit value back into three 8-bit values.
			base64Decode = base64Decode & chr( (n \ (2^16)) and 255 ) & chr( (n \ (2^8)) and 255 ) & chr( n and 255 )

		next

		' Trim off any excess space.
		base64Decode = left(base64Decode,len(base64Decode)-p)
	End Function



End Class


Class FileHandling


	Function RemoveFolder(sPath)
		RemoveFolder = RemoveFolderEx (sPath, TRUE)
	End Function 
	
	Function RemoveFolderEx(sPath, bLogging )

		Dim oFile, oFolder

		If bLogging then
			oLogging.CreateEntry "Remove Folder: " & sPath, LogTypeInfo
		End if 

		' Make sure the folder exists

		If not oFSO.FolderExists(sPath) then
			Exit Function
		End if


		' First try to remove any files

		For each oFile in oFSO.GetFolder(sPath).Files
			DeleteFileEx oFile.Path, bLogging
		Next


		' Then take care of subfolders

		For each oFolder in oFSO.GetFolder(sPath).Subfolders
			RemoveFolderEx oFolder.Path, bLogging
		Next


		' Now try to remove the folder

		On Error Resume Next
		oFSO.DeleteFolder sPath, true
		If bLogging then
			TestAndLog Err = 0, "Delete Folder: " & sPath
		End if 
		On Error Goto 0

	End Function
	
	Function DeleteFile(sFile)
		DeleteFile = DeleteFileEx ( sFile, TRUE )
	End function 
	
	Function DeleteFileEx(sFile, bLogging)
	
		If bLogging then
			oLogging.CreateEntry "Delete File: " & sFile, LogTypeInfo
		End if 
		On Error Resume Next
		oFSO.DeleteFile sFile, TRUE
		If bLogging then
			TestAndLog Err = 0, "Delete File: " & sFile
		End if 
		Err.Clear
		On Error Goto 0
		
	End function
	
	Function MoveFile(sFile,sDest)
		MoveFile = MoveFileEx(sFile,sDest,True)
	End function 
	
	Function MoveFileEx(sFile,sDest,bLogging)
	
		If bLogging then
			oLogging.CreateEntry "Move File: " & sFile & " to " & sDest	, LogTypeInfo
		End if 
		On Error Resume Next
		oFSO.MoveFile sFile, sDest
		If bLogging then
			TestAndLog Err = 0, "Move File: " & sFile & " to " & sDest		
		End if 
		Err.Clear
		On Error Goto 0
		
	End function
	
	Function CopyFile(sFile,sDest, bOverwrite)
		CopyFile = CopyFileEx(sFile,sDest, bOverwrite,True)
	End function 

	Function CopyFileEx(sFile,sDest, bOverwrite,bLogging)
	
		If bLogging then
			oLogging.CreateEntry "Copy File: " & sFile & " to " & sDest	, LogTypeInfo
		End if 
		On Error Resume Next
		oFSO.CopyFile sFile, sDest, bOverwrite
		If bLogging then
			TestAndLog Err = 0, "Copy File: " & sFile & " to " & sDest		
		End if 
		Err.Clear
		On Error Goto 0
		
	End function
	
	Function CopyFolder (sSource, sDest, bOverwrite)
		CopyFolder = CopyFolderEx(sSource, sDest, bOverwrite, True)
	End function
	
	Function CopyFolderEx (sSource, sDest, bOverwrite,bLogging)
	
		If bLogging then
			oLogging.CreateEntry "Copy Folder: " & sSource & " to " & sDest	, LogTypeInfo
		End if 
		On Error Resume Next
		oFSO.CopyFolder  sSource, sDest, bOverwrite
		If bLogging then
			TestAndLog Err = 0, "Copy Folder: " & sSource & " to " & sDest		
		End if 
		Err.Clear
		On Error Goto 0
		
	End function
	
	Function MoveFolder (sSource, sDest ) 
		MoveFolder = MoveFolderEx (sSource, sDest, True ) 
	End function 
	
	Function MoveFolderEx (sSource, sDest, bLogging ) 

		If bLogging then
			oLogging.CreateEntry "Move Folder " & sSource & " to " &  sDest , LogTypeInfo
		End if 
		On Error Resume Next
		oFSO.MoveFolder sSource,sDest
		If bLogging then
			TestAndLog Err = 0, "Move Folder " & sSource & " to " &  sDest 
		End if 
		On Error Goto 0

	End function 

	Function NormalizePath( Path )
		Dim i, j
		Dim SplitPath

		SplitPath = Split(Replace(Path, """", "" ), "\" )

		For i = lBound(SplitPath) to UBound(SplitPath)
			If SplitPath(i) = "." then
				SplitPath(i) = empty
			Elseif SplitPath(i) = ".." then
				SplitPath(i) = empty
				For j = i-1 to 0  step -1
					If SplitPath(j) <> empty then
						SplitPath(j) = empty
						Exit for
					End if
				Next
			End if
		Next

		Normalizepath = Join(SplitPath, "\")

		Do while Instr(2, NormalizePath, "\\", vbTextCompare) <> 0
			If left(NormalizePath,2) = "\\" then
				NormalizePath = "\" & Replace(NormalizePath, "\\", "\", 2, -1)
			Else
				NormalizePath = Replace(NormalizePath, "\\", "\", 1, -1)
			End if
		Loop

	End function
	
End Class


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Regular Expression Helper functions
'    http://msdn.microsoft.com/en-us/library/6wzad2b2(VS.85).aspx
'
' Usage: 
'  for each Match in oRegEx.GetRegExMatchesFromFile( "\[([^\]]*)\]","c:\boot.ini" ) 
'     Wscript.echo "Match: " & Match
'  next
'

dim g_oRegEx
Function oRegEx
	if isempty(g_oRegEx) then
		set g_oRegEx = new clsRegEx
	end if 
	set oRegEx = g_oRegEx
end Function 

class clsRegEx

	private Function RegExObj ( SearchPattern, IgnoreCase, IsGlobal ) 
		set RegExObj = New RegExp
		RegExObj.Global = IsGlobal
		RegExObj.Multiline = TRUE
		RegExObj.IgnoreCase = IgnoreCase
		RegExObj.Pattern = SearchPattern   
	end Function

	Function GetRegExMatches ( SearchPattern, Buffer ) 
		set GetRegExMatches = RegExObj ( SearchPattern, TRUE, TRUE ).Execute(Buffer)
	end Function

	Function TestRegEx ( SearchPattern, Buffer ) 
		TEstRegEx = RegExObj ( SearchPattern, TRUE, TRUE ).Test(Buffer)
	end Function 

	Function ReplaceRegEx ( SearchPattern, ReplacementString, Buffer ) 
		ReplaceRegEx = RegExObj ( SearchPattern, TRUE, TRUE ).Replace(Buffer, ReplacementString)
	end Function 
	

	Function GetRegExMatchesFromFile ( SearchPattern, FileName ) 
		with fso.OpenTextFile(oShlUtl.EES(FileName), ForReading, FALSE )
			set GetRegExMatchesFromFile = GetRegExMatches ( SearchPattern, .ReadAll ) 
		end with 
	end Function

	Function TestRegExFromFile ( SearchPattern, FileName ) 
		with fso.OpenTextFile(oShlUtl.EES(FileName), ForReading, FALSE )
			TestRegExFromFile = TestRegEx ( SearchPattern, .ReadAll ) 
		end with 
	end Function

	Function ReplaceRegExFromFile ( SearchPattern, ReplacementString, FileName ) 
		with fso.OpenTextFile(oShlUtl.EES(FileName), ForReading, FALSE )
			ReplaceRegExFromFile = ReplaceRegEx ( SearchPattern, ReplacementString, .ReadAll )
		end with 
	end Function

	Function FindStrInFile( SearchStr, FileName )
		with fso.OpenTextFile(oShlUtl.EES(FileName), ForReading, FALSE )
			FindStrInFile = InStr(1, .ReadAll, SearchStr, vbTextCompare)
		end with 
	end Function 

end class

