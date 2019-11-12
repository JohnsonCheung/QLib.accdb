Attribute VB_Name = "MxXlsA1"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsA1."
Function A1zWb(A As Workbook, Optional NewWsn$) As Range ' Return A1 of a new Ws (with NewWsn) in Wb
Set A1zWb = A1zWs(AddWs(A, NewWsn))
End Function


Function A1zLo(Lo As ListObject) As Range
Set A1zLo = RgRC(Lo.ListColumns(1).Range, 2, 1)
End Function


Function A1zRg(A As Range) As Range
Set A1zRg = RgRC(A, 1, 1)
End Function

Function A2zWs(A As Worksheet) As Range
Set A2zWs = A.Range("A2")
End Function

Function A1zWs(A As Worksheet) As Range
Set A1zWs = A.Range("A1")
End Function

