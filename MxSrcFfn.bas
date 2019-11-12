Attribute VB_Name = "MxSrcFfn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxSrcFfn."

Function SrcFfnM$()
SrcFfnM = SrcFfn(CCmp)
End Function

Function SrcFfn$(A As VBComponent)
SrcFfn = SrcpzCmp(A) & SrcFn(A)
End Function

Function SrcFn$(A As VBComponent)
SrcFn = A.Name & ".bas"
End Function

Function ExtzCmpTy$(A As vbext_ComponentType)
Dim O$
Select Case A
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case vbext_ct_MSForm: O = ".cls"
Case Else: Err.Raise 1, , "SrcExt: Unexpected Md_CmpTy.  Should be [Class or Module or Document]"
End Select
ExtzCmpTy = O
End Function

Function SrcFfnzMdn$(Mdn$)
SrcFfnzMdn = SrcFfn(Cmp(Mdn))
End Function

Function SrcFfnzM$(M As CodeModule)
SrcFfnzM = SrcFfn(M.Parent)
End Function
