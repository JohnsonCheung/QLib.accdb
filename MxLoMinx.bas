Attribute VB_Name = "MxLoMinx"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxLoMinx."
Sub MinxLo(A As ListObject)
If Fst2Chr(A.Name) <> "T_" Then Exit Sub
Dim R1 As Range
Set R1 = A.DataBodyRange
If R1.Rows.Count >= 2 Then
    RgRR(R1, 2, R1.Rows.Count).EntireRow.Delete
End If
End Sub

Sub MinxLozWs(A As Worksheet)
If A.CodeName = "WsIdx" Then Exit Sub
If Fst2Chr(A.CodeName) <> "Ws" Then Exit Sub
Dim L As ListObject
For Each L In A.ListObjects
    MinxLo L
Next
End Sub

Sub MinxLozWb(A As Workbook)
Dim Ws As Worksheet
For Each Ws In A.Sheets
    MinxLozWs Ws
Next
End Sub

Sub MinxLozFx(Fx)
Dim O As Workbook: Set O = WbzFx(Fx)
MinxLozWb O
SavWb O
ClsWbNoSav O
End Sub

