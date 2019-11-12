Attribute VB_Name = "MxLnoStr"
Option Explicit
Option Compare Text
Const CLib$ = "QTp."
Const CMod$ = CLib & "MxLnoStr."
Const NLnoStrDig As Byte = 2

Function Lnoss_FmLnoCol_WhStrCol_HasS$(LnoCol&(), StrCol$(), S)
Dim OLno$()
Dim J&: For J = 0 To UB(LnoCol)
    If StrCol(J) = S Then
        PushI OLno, LnoStr(LnoCol(J))
    End If
Next
Lnoss_FmLnoCol_WhStrCol_HasS = JnSpc(OLno)
End Function

Function Lnoss_FmLnoCol_WhSyCol_HasS$(LnoCol&(), SyCol(), S)
Dim OLno$()
Dim J&: For J = 0 To UB(LnoCol)
    If HasEle(SyCol(J), S) Then
        PushI OLno, LnoStr(LnoCol(J))
    End If
Next
Lnoss_FmLnoCol_WhSyCol_HasS = JnSpc(OLno)
End Function

Function LnoStr$(L&)
LnoStr = AlignR(L, NLnoStrDig)
End Function
