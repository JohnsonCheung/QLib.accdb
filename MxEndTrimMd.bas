Attribute VB_Name = "MxEndTrimMd"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxEndTrimMd."
Function EndTrimMdLasLin(M As CodeModule) As Boolean
Dim N&: N = M.CountOfLines
    If N = 0 Then Exit Function
    If IsLinCd(M.Lines(N, 1)) Then Exit Function
EndTrimMdLasLin = True
M.DeleteLines N, 1
End Function

Sub EndTrimMdM()
EndTrimMd CMd
End Sub

Sub EndTrimMd(M As CodeModule)
Dim J%, Trimmed As Boolean
While EndTrimMdLasLin(M)
    J = J + 1: If J > 10000 Then ThwLoopingTooMuch CSub
    Trimmed = True
Wend
If Trimmed Then Debug.Print "EndTrimMd: Module is trimmed [" & Mdn(M) & "]"
End Sub

Sub EndTrimMdzP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    EndTrimMd C.CodeModule
Next
End Sub

Sub EndTrimMdP()
EndTrimMdzP CPj
End Sub
