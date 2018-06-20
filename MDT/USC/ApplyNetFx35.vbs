wscript.echo "Getting current OS drive letter"
sSystemDrive = "c:"
wscript.echo "OS drive letter is " & sSystemDrive
wscript.echo "About to inject roles offline"
sScratchDir = sSystemDrive & "\Windows\temp"
wscript.echo "ScratchDir is " & sScratchDir
sNetFxSource = "deployroot\windows00\sources\sxs"
wscript.echo "SXS source is " & sNetFxSource 
sCmd = "dism.exe /Image:""" & sSystemDrive & _
"""  /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:""" & _
sNetFxSource & """ /ScratchDir:""" & sScratchDir & """"
wscript.echo "about to run: " & sCmd
wscript.echo "done"