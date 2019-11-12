Attribute VB_Name = "MxDoDcl"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoDcl."

Function DoDclP() As Drs
DoDclP = DoDclzP(CPj)
End Function

Function DoDclzP(P As VBProject) As Drs
Dim Dy(), Pjn$
Pjn = P.Name
Dim C As VBComponent: For Each C In P.VBComponents
    Dim M As CodeModule: Set M = C.CodeModule
    If M.CountOfDeclarationLines > 0 Then
        Dim L$: L = DcllzM(M)
        PushI Dy, Array(Pjn, C.Name, L, L)
    End If
Next
DoDclzP = DrszFF(FFoDcl, Dy)
End Function


