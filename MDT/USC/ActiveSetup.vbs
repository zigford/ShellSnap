	wscript.echo "Running ActiveSetup.wsf"
	wscript.echo "Will enumerate all HKLM Active Setup commands and execute as local Administrator account"
Const HKLM = &H80000002

Const REG_SZ = 1 
Const REG_EXPAND_SZ = 2 
Const REG_BINARY = 3 
Const REG_DWORD = 4 
Const REG_MULTI_SZ = 7 
  
computer = "."
Set shell = CreateObject("WScript.Shell") 
set objEnv = shell.Environment("PROCESS")
Set registry = GetObject("winmgmts:\\" & computer & "\root\default:StdRegProv") 

wscript.echo "Setting SEE_MASK_NOZONECHECKS to 1"
objEnv("SEE_MASK_NOZONECHECKS") = 1

keyPath = "SOFTWARE\Microsoft\Active Setup\Installed Components" 
psexecCMD = "psexec.exe /i /accepteula /h /u Administrator /p Osorio "
registry.EnumKey HKLM, keyPath, keyNames

If Not IsNull(keyNames) then
	For Each Key in KeyNames
	NewKeyPath = keyPath & "\" & Key
		registry.EnumValues HKLM, NewKeyPath, valueNames, valueTypes
		If Not IsNull(valueNames) then
			For i=0 to UBound(valueNames)
				valueName = valueNames(i)
				If valueName = "StubPath" then
					registry.GetStringValue HKLM, NewKeyPath, valueName, value
					'wscript.echo "Going to run: " & value
					returnCode = shell.Run(psexecCMD & value, 1, True)
					wscript.echo Key & " RC: " & returnCode 
					'registry.DeleteValue HKLM, keyPath, valueName
					'wscript.echo value & " Success!"
					'Now create local user component
					registry.EnumValues HKLM, NewKeyPath, SuccessNames, SuccessTypes
					If Not IsNull(SuccessNames) then
						'wscript.echo "Finding relevant values to copy"
						For b=0 to UBound(SuccessNames)
							SuccessName = SuccessNames(b)	
							If ((SuccessName = "Locale") Or (SuccessName = "Version")) then
								'wscript.echo "relevant value found"
								registry.GetStringValue HKLM, NewKeyPath, SuccessName, SuccessValue
								regCMD = "reg ADD ""HKCU\" & NewKeyPath & """ /v " & SuccessName & " /d " & SuccessValue & " /f"
								'wscript.echo "Going to run " & regCMD
								returnCode = shell.Run(psexecCMD & regCMD, 1, True)
								wscript.echo "Active Setup population RC: " & returnCode
							End If
						Next
					End If
				End If 
			Next
		End If
	Next
End If
wscript.echo "Removing SEE_MASK_NOZONECHECKS"
objEnv.Remove("SEE_MASK_NOZONECHECKS")