Attribute VB_Name = "MxAyEle"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAyEle."
Function LasSndEle(Ay)
Dim N&: N = Si(Ay)
If N <= 1 Then
    Thw CSub, "Only 1 or no ele in Ay"
Else
    Asg Ay(N - 2), LasSndEle
End If
End Function

