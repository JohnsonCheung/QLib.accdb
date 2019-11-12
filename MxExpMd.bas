Attribute VB_Name = "MxExpMd"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxExpMd."

Sub ExpMdM()
ExpMd CMd
End Sub

Sub ExpMd(M As CodeModule)
EndTrimMd M
Dim F$: F = SrcFfnzM(M)
M.Parent.Export F
'TrimLasEmpLinzFt F
End Sub
