Attribute VB_Name = "MxSrcp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxSrcp."

Function SrcpzCmp$(A As VBComponent)
SrcpzCmp = SrcpzP(PjzC(A))
End Function

Function SrcpzPjf$(Pjf, Optional Libv$)
SrcpzPjf = EnsPth(Pjf & ".src")
End Function

Sub EnsSrcp(P As VBProject)
EnsPthAll SrcpzP(P)
End Sub

Function SrcpzDistPj$(DistPj As VBProject)
Dim P$: P = Pjp(DistPj)
SrcpzDistPj = AddFdrAp(UpPth(P, 1), ".Src", Fdr(P))
End Function

Function SrcpP$(Optional Libv$)
SrcpP = SrcpzP(CPj, Libv)
End Function

Function SrcpzP$(P As VBProject, Optional Libv$)
SrcpzP = SrcpzPjf(Pjf(P), Libv)
End Function

Sub BrwSrcpP()
BrwPth SrcpP
End Sub

Function IsSrcp(Pth) As Boolean
Dim F$: F = Fdr(Pth)
If Not HasExtss(F, ".xlam .accdb") Then Exit Function
IsSrcp = Fdr(ParPth(Pth)) = ".Src"
End Function

Sub ThwNotSrcp(Srcp$)
If Not IsSrcp(Srcp) Then Err.Raise 1, , "Not Srcp:" & vbCrLf & Srcp
End Sub

