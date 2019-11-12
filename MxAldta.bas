Attribute VB_Name = "MxAldta"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxAldta."
Function RgzAldta(S As Worksheet) As Range
Set RgzAldta = S.Range(S.Cells(1, 1), LasCell(S))
End Function

Function SqzAldta(S As Worksheet) As Variant()
SqzAldta = RgzAldta(S).Value
End Function

Function DrszAldta(S As Worksheet) As Drs
DrszAldta = DrszSq(SqzAldta(S))
End Function
