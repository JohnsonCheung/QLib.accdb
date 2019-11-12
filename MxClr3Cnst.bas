Attribute VB_Name = "MxClr3Cnst"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxClr3Cnst."


Sub ClrCModM()
ClrCModzM CMd
End Sub

Sub ClrCModzM(M As CodeModule)
ClrCnstLin M, "CMod"
End Sub

Sub ClrCLibzM(M As CodeModule)
ClrCnstLin M, "CMod"
End Sub

Sub RmvCModM()
RmvCModzM CMd
End Sub

Sub RmvCModzM(M As CodeModule)
RmvCnstLin M, "CMod"
End Sub

Sub RmvCLibM()
RmvCLibzM CMd
End Sub

Sub RmvCLibzM(M As CodeModule)
RmvCnstLin M, "CLib"
End Sub

Sub RmvCLibP()
RmvCLibzP CPj
End Sub

Sub RmvCLibzP(P As VBProject)
RmvCnstLinzP P, "CLib", IsPrvOnly:=True
End Sub
Sub RmvCModP()
RmvCModzP CPj
End Sub

Sub RmvCModzP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    RmvCModzM C.CodeModule
Next
End Sub

