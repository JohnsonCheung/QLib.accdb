Attribute VB_Name = "MxEntCLib"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxEntCLib."
Sub EntCLibP()
EntCLibzP CPj
End Sub
Sub EntCLibzP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    If EntCLibzM(C.CodeModule) Then Exit Sub
Next
End Sub

Sub EntCLibM()
EntCLibzM CMd
End Sub

Function EntCLibzM(M As CodeModule) As Boolean
Static LasCModv$
If Not IsMd(M.Parent) Then Exit Function
Dim V$: V = CLibv(DclzM(M)): If V <> "" Then Exit Function
V = InputBox("Enter CLibv: ", "For Md: " & M.Name, LasCModv)
If V = "" Then EntCLibzM = True: Exit Function
If FstChr(V) <> "Q" Then Exit Function
EnsCnstLin M, CLibLin(V)
LasCModv = V
End Function

