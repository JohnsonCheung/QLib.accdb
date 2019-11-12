Attribute VB_Name = "MxChkIde"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxChkIde."
Function ChkMdn(P As VBProject, Mdn) As Boolean
If Has_Mdn_InPj(P, Mdn) Then Exit Function
MsgBox FmtQQ("Mdn not found: ?|In Pj: ?", Mdn, P.Name), vbCritical
ChkMdn = True
End Function
