Option Explicit

Dim strModel
Dim objWMIServices
Dim colComputer
Dim objComp
Dim WshShell
Dim Shell

Set WshShell = CreateObject("Wscript.Shell")
Set objWMIServices = GetObject("winmgmts:")
Set colComputer = objWMIServices.Execquery("select * from win32_ComputerSystem")
Set Shell = WshShell.Environment("User")
  
For each objComp in colComputer
    Shell( "ComputerModel" ) = objComp.Model
Next
