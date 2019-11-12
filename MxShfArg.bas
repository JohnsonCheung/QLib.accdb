Attribute VB_Name = "MxShfArg"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxShfArg."
Function ShfArgMdy$(OArg$)
ShfArgMdy = ShfPfxAyS(OArg, ArgMdyAy)
End Function
Function ShfArgSfx$(OLin$)
Dim P%: P = InStr(OLin, "=")
If P > 0 Then
    ShfArgSfx = Left(OLin, P - 2)
    OLin = Mid(OLin, P - 1)
    Exit Function
Else
    ShfArgSfx = OLin
    OLin = ""
End If
End Function

Function ShfShtArgSfx$(OLin$)
ShfShtArgSfx = ShtArgSfx(ShfArgSfx(OLin))
End Function
