Attribute VB_Name = "MxWsRg"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxWsRg."
Function WsC(A As Worksheet, C) As Range
Dim R As Range
Set R = A.Columns(C)
Set WsC = R.EntireColumn
End Function

Function WsCC(A As Worksheet, C1, C2) As Range
Set WsCC = WsRCC(A, 1, C1, C2).EntireColumn
End Function

Function WsCRR(A As Worksheet, C, R1, R2) As Range
Set WsCRR = WsRCRC(A, R1, C, R2, C)
End Function

Function WsRC(A As Worksheet, R, C) As Range
Set WsRC = A.Cells(R, C)
End Function

Function WsRCC(A As Worksheet, R, C1, C2) As Range
Set WsRCC = WsRCRC(A, R, C1, R, C2)
End Function

Function WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = A.Range(WsRC(A, R1, C1), WsRC(A, R2, C2))
End Function

Function WsRR(A As Worksheet, R1, R2) As Range
Set WsRR = A.Range(WsRC(A, R1, 1), WsRC(A, R2, 1)).EntireRow
End Function

