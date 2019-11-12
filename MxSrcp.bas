Attribute VB_Name = "MxSrcp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxSrcp."

Function SrcpzCmp$(A As VBComponent)
SrcpzCmp = SrcpzP(PjzC(A))
End Function

Function SrcpzPjf$(Pjf)
SrcpzPjf = EnsPth(Pjf & ".src")
End Function

Sub EnsSrcp(P As VBProject)
EnsPthAll SrcpzP(P)
End Sub

Function SrcpzDistPj$(DistPj As VBProject)
Dim P$: P = Pjp(DistPj)
SrcpzDistPj = AddFdrAp(UpPth(P, 1), ".Src", Fdr(P))
End Function

Function SrcpP$()
SrcpP = SrcpzP(CPj)
End Function

Function SrcpzP$(P As VBProject)
SrcpzP = SrcpzPjf(Pjf(P))
End Function
Sub BrwSrcpP()
BrwPth SrcpP
End Sub

