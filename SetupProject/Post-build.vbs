Option Explicit

If WScript.arguments.Count <> 1 Then
	MsgBox "64-bit Installer Post-build scripting: Wrong number of args."
	WScript.Quit
End IF

Const msiOpenDatabaseModeTransact     = 1
Const msiViewModifyAssign         = 3
Const msiOpenDatabaseModeDirect = 2

Dim installer : Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer")

Dim sqlQuery : sqlQuery = "SELECT `Name`, `Data` FROM `Binary` WHERE `Name`='InstallUtil'"

Dim database : Set database = installer.OpenDatabase(WScript.arguments(0), msiOpenDatabaseModeDirect)
Dim view     : Set view = database.OpenView(sqlQuery)
Dim record

view.Execute

Set record = view.Fetch()

record.SetStream 2, "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\InstallUtilLib.dll"

view.Modify msiViewModifyAssign, record 
database.Commit 
Set view = Nothing
Set database = Nothing