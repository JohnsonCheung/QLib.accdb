Attribute VB_Name = "MxInstall"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxInstall."
Function TreeWsMdLines$()
Erase XX
X "Option Explicit"
X "Sub Worksheet_Change(ByVal Target As Range)"
X "MTreeWs.Change Target"
X "End Sub"

X "Sub Worksheet_SelectionChange(ByVal Target As Range)"
X "MTreeWs.SelectionChange Target"
X "End Sub"
TreeWsMdLines = JnCrLf(XX)
Erase XX
End Function

Sub InstallTreeWs()
Dim Ws, Wb
For Each Ws In Itr(TreeWsAy)
    InstallTreeWsz CvWs(Ws)
Next
For Each Wb In Itr(TreeWbAy)
    InstallTreeWbz CvWb(Wb)
Next
End Sub
Sub InstallTreeWbz(Wb As Workbook)
Exit Sub
'Wb.VBProject.References.AddFromFile ShpCstFxa
End Sub
Function IsTreeWb(A As Workbook) As Boolean
Dim Ws As Worksheet
For Each Ws In A.Sheets
    If Ws.Name = "TreeWs" Then IsTreeWb = True: Exit Function
Next
End Function
Function TreeWbAy() As Workbook()
Dim Wb As Workbook, Ws As Worksheet
For Each Wb In Xls.Workbooks
    If IsTreeWb(Wb) Then PushObj TreeWbAy, Wb
Next

End Function
Function TreeWsAy() As Worksheet()
Dim Wb As Workbook, Ws As Worksheet
For Each Wb In Xls.Workbooks
    For Each Ws In Wb.Sheets
        If Ws.Name = "TreeWs" Then PushObj TreeWsAy, Ws
    Next
Next
End Function

Sub InstallTreeWsz(A As Worksheet)
Dim Md As CodeModule
Set Md = MdzWs(A)
Stop
If Md.CountOfLines = 0 Then
    Md.AddFromString TreeWsMdLines
    InfLin CSub, "TreeWs in Wb is installed with code", "Wb", WbnzWs(A)
Else
    InfLin CSub, "TreeWs in Wb already has code", "Wb", WbnzWs(A)
End If
End Sub
