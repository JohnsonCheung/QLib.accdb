Attribute VB_Name = "MxSrcPth"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxSrcPth."

Function HasExtss(Ffn, ExtSsLin) As Boolean
Dim E$: E = Ext(Ffn)
Dim Sy$(): Sy = SyzSS(ExtSsLin)
HasExtss = HasEleS(Sy, E)
End Function
Function IsSrcp(Pth) As Boolean
Dim F$: F = Fdr(Pth)
If Not HasExtss(F, ".xlam .accdb") Then Exit Function
IsSrcp = Fdr(ParPth(Pth)) = ".Src"
End Function

Sub ThwNotSrcp(Srcp$)
If Not IsSrcp(Srcp) Then Err.Raise 1, , "Not Srcp:" & vbCrLf & Srcp
End Sub

Function IsInstScrp(Pth) As Boolean
If NoPth(Pth) Then Exit Function
If Fdr(Pth) <> "Src" Then Exit Function
Dim P$: P = ParPth(Pth)
If Not IsTimStr(Fdr(P)) Then Exit Function
IsInstScrp = True
End Function

Sub ThwNotInstScrp(Pth)
If Not IsInstScrp(Pth) Then Err.Raise 1, , "Not InstScrp(" & Pth & ")"
End Sub
