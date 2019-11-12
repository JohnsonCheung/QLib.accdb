Attribute VB_Name = "MxCmp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCmp."
Function IsMod(A As VBComponent) As Boolean
IsMod = A.Type = vbext_ct_StdModule
End Function

Function IsCls(A As VBComponent) As Boolean
IsCls = A.Type = vbext_ct_ClassModule
End Function

Function IsMd(A As VBComponent) As Boolean
Select Case A.Type
Case vbext_ct_ClassModule, vbext_ct_StdModule: IsMd = True
End Select
End Function
