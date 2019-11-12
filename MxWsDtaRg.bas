Attribute VB_Name = "MxWsDtaRg"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxWsDtarg."
Sub ClrDtarg(A As Worksheet)
DtaDtarg(A).Clear
End Sub

Function DtaDtarg(Ws As Excel.Worksheet) As Range
Set DtaDtarg = Ws.Range(A1zWs(Ws), LasCell(Ws))
End Function


