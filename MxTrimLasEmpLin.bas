Attribute VB_Name = "MxTrimLasEmpLin"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxTrimLasEmpLin."
Function HasLasEmpLin(Ly$()) As Boolean
Dim N&: N = Si(Ly)
If N = 0 Then Exit Function
Dim O As Boolean: O = Ly(N - 1) = ""
HasLasEmpLin = O
End Function

Sub TrimLasEmpLinzFt(Ft$)
Dim Ly$(): Ly = LyzFt(Ft)
If HasLasEmpLin(Ly) Then
    WrtAy RmvLasEle(Ly), Ft
End If
End Sub
