' BET PRO - pornire automata (o data pe zi, din Task Scheduler)
' Pune acest fisier in ACELASI folder cu BET_PRO.xlsm.
Option Explicit
Dim fso, folder, cale, xl, wb
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
cale = folder & "\BET_PRO.xlsm"
If Not fso.FileExists(cale) Then
    WScript.Quit 1
End If
Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AutomationSecurity = 1
Set wb = xl.Workbooks.Open(cale)
On Error Resume Next
xl.Run "'" & wb.Name & "'!RuleazaAutomat"
wb.Save
wb.Close False
xl.Quit
Set wb = Nothing
Set xl = Nothing
