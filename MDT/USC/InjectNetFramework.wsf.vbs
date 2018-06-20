<job id="Roles-InjectNetFx35">
	<script language="VBScript" src="ZTIUtility.vbs"/>
	<script language="VBScript">

	oLogging.CreateEntry "Getting OS drive letter", LogTypeInfo
	sSystemDrive = oEnvironment.Item("OSDISK")
	oLogging.CreateEntry "OS drive letter is " & sSystemDrive, LogTypeInfo
	sScratchDir = sSystemDrive + "\Windows\temp"
	oLogging.CreateEntry "ScratchDir is " + sScratchDir, LogTypeInfo
	sNetFxSource = oEnvironment.Item("SourcePath") + "\sources\sxs"
	oLogging.CreateEntry "SXS source is " + sNetFxSource, LogTypeInfo
	sCmd = "dism.exe /Image:""" + sSystemDrive + """  /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:""" + sNetFxSource _
	""" /ScratchDir:""" + sScratchDir """"
	oLogging.CreateEntry "about to run: " + sCmd, LogTypeInfo
	oUtility.RunWithConsoleLogging(sCmd)
	oLogging.CreateEntry "done", LogTypeInfo	
	</script>
</job>
