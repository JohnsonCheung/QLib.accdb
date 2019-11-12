Attribute VB_Name = "MxEns3Cnst"
Option Explicit
Option Compare Text
Const CNs$ = "AA"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxEns3Cnst."


Sub EnsCLibzM(M As CodeModule, CLibv$)
If Not IsMd(M.Parent) Then Exit Sub
EnsCnstLin M, CLibLin(CLibv)
End Sub

Sub EnsCNsLin(M As CodeModule, Ns$)
If Not IsMd(M.Parent) Then Exit Sub
EnsCnstLin M, CNsLin(Ns)
End Sub

Sub EnsCNszM(M As CodeModule, Ns$)
EnsCnstLin M, CNsLin(Ns)
End Sub

Sub EnsCModM()
EnsCModzM CMd
End Sub

Sub EnsCModP()
EnsCModzP CPj
End Sub

Sub EnsCModzM(M As CodeModule)
EnsCnstLinAft M, CModLin(M), "CLib", IsPrvOnly:=True
End Sub

Sub EnsCModzP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    EnsCModzM C.CodeModule
Next
End Sub
