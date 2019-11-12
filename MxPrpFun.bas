Attribute VB_Name = "MxPrpFun"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxPrpFun."
Dim Inf$()
Function IsLinPrpFun(Lin) As Boolean
Dim L$, b$
L = RmvMdy(Lin)
If ShfMthTy(L) <> "Function" Then Exit Function
If ShfNm(L) = "" Then Exit Function
ShfTyChr L
IsLinPrpFun = Left(L, 2) = "()"
End Function

Function PrpFunLnoAy(M As CodeModule) As Long()
Dim J&, L
For Each L In Src(M)
    J = J + 1
    If IsLinPrpFun(L) Then PushI PrpFunLnoAy, J
Next
End Function


Sub EnsPjFunzMd(Md As CodeModule, Optional WhatIf As Boolean)
Dim L
For Each L In Itr(PrpFunLnoAy(Md))
    EnsPrpFunMdLno Md, L, WhatIf
Next
End Sub

Sub EnsPrpFunzP(Pj As VBProject, Optional WhatIf As Boolean)
Dim I
Erase Inf
'For Each I In MdItr(Pj)
'    EnsPjFunMd CvMd(I), WhatIf
'Next
Brw Inf
End Sub

Sub EnsPrpFun()
EnsPjFunzMd CMd
End Sub

Sub EnsPrpFunMdLno(M As CodeModule, Lno, Optional WhatIf As Boolean)
Dim OldLin
Dim NewLin
    OldLin = M.Lines(Lno, 1)
    NewLin = Replace(M.Lines(Lno, 1), "Function", "Property Get")
If Not WhatIf Then M.ReplaceLine Lno, NewLin
PushI Inf, "EnsPrpFun:EnsPrpFunMdLno NewLin: " & OldLin
PushI Inf, "                 OldLin: " & NewLin
End Sub
