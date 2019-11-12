Attribute VB_Name = "MxCpyCmp"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCpyCmp."
Sub ThwIf_NotCls(M As CodeModule, Fun$)
If M.Parent.Type = vbext_ct_ClassModule Then Thw Fun, "Should be a Cls", "ShtCmpTy", ShtCmpTy(M.Parent.Type)
End Sub

Sub CpyCls(M As CodeModule, ToPj As VBProject)
Const CSub$ = CMod & "CpyCls"
ThwIf_NotCls M, CSub
If HasCmpzP(ToPj, M.Name) Then
    InfLin CSub, "Cls is fnd in ToPj", "Cls ToPj", Mdn(M), ToPj.Name
    Exit Sub
End If
Dim T$: T = TmpFt
M.Parent.Export T
ToPj.VBComponents.Import T
Kill T
End Sub

Sub CpyCmp(A As VBComponent, ToPj As VBProject)
If IsCls(A) Then
    CpyCls A.CodeModule, ToPj 'If ClassModule need to export and import due to the Public/Private class property can only the set by Export/Import
Else
    CpyMod A.CodeModule, ToPj
End If
End Sub
Function NotMod(M As CodeModule, Fun$) As Boolean
If M.Parent.Type <> vbext_ct_StdModule Then
    NotMod = True
    InfLin CSub, "Should be a mod", "Mdn", Mdn(M)
End If
End Function

Sub CpyMod(M As CodeModule, ToPj As VBProject)
AddCmpzL ToPj, M.Name, Srcl(M)
End Sub


Function CpyMd(M As CodeModule, ToM As CodeModule) As Boolean
'Ret : Cpy @M to @ToM and  both must exist @@
CpyMd = RplMd(ToM, Srcl(M))
End Function

