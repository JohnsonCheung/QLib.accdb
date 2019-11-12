Attribute VB_Name = "MxResOp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxResOp."
Sub EdtRes(ResFn$, Optional Pseg$)
Dim F$: F = ResFfn(ResFn, Pseg): If NoFfn(F) Then EnsFt F
VcFt F
End Sub
