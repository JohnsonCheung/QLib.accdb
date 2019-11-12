Attribute VB_Name = "MxLoDta"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxLoDta."
#Const Doc = False
#If Doc Then
Return Dta from Lo
#End If
Function DrzLoCell(Lo As ListObject, Cell As Range) As Variant()
Dim Ix&: Ix = LoRno(Lo, Cell): If Ix = -1 Then Exit Function
DrzLoCell = VvyzRg(Lo.ListRows(Ix).Range)
End Function
