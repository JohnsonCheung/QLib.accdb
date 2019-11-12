Attribute VB_Name = "MxAyDry"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAyDry."

Function DyoAddAyC(Ay, C) As Variant()
'XCDy is AyItmX-Const-Dy
Dim I
For Each I In Itr(Ay)
    PushI DyoAddAyC, Array(I, C)
Next
End Function

Function DyoCAddAy(A, C) As Variant()
'CDyzS12s is Const-AyItmX-Dy
Dim I
For Each I In Itr(A)
    PushI DyoCAddAy, Array(C, I)
Next
End Function

Function DyoAyzTyNmVal(Ay) As Variant()
Dim I
For Each I In Itr(Ay)
    PushI DyoAyzTyNmVal, Array(TypeName(I), I)
Next
End Function

Sub DmpAyzTyNmVal(Ay)
DmpDy DyoAyzTyNmVal(Ay)
End Sub
