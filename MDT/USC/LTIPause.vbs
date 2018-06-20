Set env = CreateObject("Microsoft.SMS.TSEnvironment") 
MsgBox "Pausing Task Sequence, click OK to continue…" & chr(13) & _
" " & chr(13) & _
"Computername:" & chr(9) & chr(9) & env("OSDComputername") & chr(13) & _
"IPAddress001:" & chr(9) & chr(9) & env("IPAddress001") & chr(13) & _
"DefaultGateWay001:" & chr(9) & env("DefaultGateway001") & chr(13) & _
"Phase:" & chr(9) & chr(9) & chr(9) & env("Phase") & chr(13) & _
"Make:" & chr(9) & chr(9) & chr(9) & env("Make") & chr(13) & _
"Model:" & chr(9) & chr(9) & chr(9) & env("Model") & chr(13) & _
"VMPlatform:" & chr(9) & chr(9) & env("VMPlatform") & chr(13) & _
"TaskSequenceID:" & chr(9) & chr(9) & env("TaskSequenceID") & chr(13) & _
"TaskSequenceName:" & chr(9) & env("TaskSequenceName") & chr(13) & _
"TaskSequenceVersion:" & chr(9) & env("TaskSequenceVersion") & chr(13) & _
"ViaServerConfig:" & chr(9) & chr(9) & env("ViaServerConfig") & chr(13) & _
"ImageFlags:" & chr(9) & chr(9) & env("ImageFlags") & chr(13) & _
"ImageBuild:" & chr(9) & chr(9) & env("ImageBuild") & chr(13) & _
"SLShare:" & chr(9) & chr(9) & chr(9) & env("SLShare") & chr(13) & _
"SLShareDynamicLogging:" & chr(9) & env("SLShareDynamicLogging") & chr(13) & _
"DeployRoot:" & chr(9) & chr(9) & env("DeployRoot") & chr(13) & _
"DriverGroup001:" & chr(9) & chr(9) & env("DriverGroup001") & chr(13) & _
"DriverGroup002:" & chr(9) & chr(9) & env("DriverGroup002") & chr(13)_
, 0, "LTIPause"