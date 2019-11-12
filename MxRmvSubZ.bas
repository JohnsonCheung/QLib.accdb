Attribute VB_Name = "MxRmvSubZ"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxRmvSubZ."
Sub RmvSubZM()
RmvSubZzM CMd
End Sub
Sub RmvSubZzM(M As CodeModule)
RmvMth M, "Z"
End Sub
Sub RmvSubZP()
RmvSubZzP CPj
End Sub
Sub RmvSubZzP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    RmvSubZzM C.CodeModule
Next
End Sub

