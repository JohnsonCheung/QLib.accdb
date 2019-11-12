Attribute VB_Name = "MxCnt"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCnt."

Sub CntCmpP()
BrwDrs DCmpCntP
End Sub

Function DCmpCntP() As Drs
DCmpCntP = DCmpCntzP(CPj)
End Function

Function DCmpCntzP(P As VBProject) As Drs
DCmpCntzP = DrszFF("Pj Tot Mod Cls Doc Frm Oth", Av(DrCmpCntzP(P)))
End Function

Function DrCmpCntzP(P As VBProject) As Variant()
Dim NCls%, NDoc%, NFrm%, NMod%, NOth%, NTot%
Dim C As VBComponent
For Each C In P.VBComponents
    Select Case C.Type
    Case vbext_ct_ClassModule:  NCls = NCls + 1
    Case vbext_ct_Document:     NDoc = NDoc + 1
    Case vbext_ct_MSForm:       NFrm = NFrm + 1
    Case vbext_ct_StdModule:    NMod = NMod + 1
    Case Else:                  NOth = NOth + 1
    End Select
    NTot = NTot + 1
Next
DrCmpCntzP = Array(P.Name, NTot, NMod, NCls, NDoc, NFrm, NOth)
End Function


Sub CntCmpzP(P As VBProject)
Brw LinzDrsR(DCmpCntzP(P))
End Sub

Function Cmp(Cmpn) As VBComponent
Set Cmp = CPj.VBComponents(Cmpn)
End Function

Function NCls%()
NCls = NClszP(CPj)
End Function

Function NCmpP%()
NCmpP = CPj.VBComponents.Count
End Function


Function NModPj%()
NModPj = NModzP(CPj)
End Function
'===============================================

Function NCmpzP%(P As VBProject)
If P.Protection = vbext_pp_locked Then Exit Function
NCmpzP = P.VBComponents.Count
End Function


Function NModzP%(P As VBProject)
NModzP = NCmpzTy(P, vbext_ct_StdModule)
End Function

Function NClszP%(P As VBProject)
NClszP = NCmpzTy(P, vbext_ct_ClassModule)
End Function

Function NDoczP%(P As VBProject)
NDoczP = NCmpzTy(P, vbext_ct_Document)
End Function

Function NCmpzTy%(P As VBProject, Ty As vbext_ComponentType)
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
Dim O%
For Each C In P.VBComponents
    If C.Type = Ty Then O = O + 1
Next
NCmpzTy = O
End Function

Function NOthCmpzP%(P As VBProject)
NOthCmpzP = NCmpzP(P) - NClszP(P) - NModzP(P) - NDoczP(P)
End Function
