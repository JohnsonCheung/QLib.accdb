Attribute VB_Name = "MxDft"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxDft."
Function Dft(V, DftV)
If IsEmp(V) Then
   Dft = DftV
Else
   Dft = V
End If
End Function

Function DftStr$(Str, Dft)
DftStr = IIf(Str = "", Dft, Str)
End Function

Function Limit(V, A, b)
Select Case V
Case V > b: Limit = b
Case V < A: Limit = A
Case Else: Limit = V
End Select
End Function
