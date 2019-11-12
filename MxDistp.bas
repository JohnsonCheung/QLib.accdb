Attribute VB_Name = "MxDistp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDistp."

Function Distp$(Srcp)
Dim A$: A = RmvPthSfx(Srcp)
If HasPfx(A, ".Src") Then Thw CSub, "Given @Srcp is not a Src path (fdr with .src at end)", "Given-Src-Pth", Srcp
Distp = EnsPth(RmvExt(A) & ".dist")
End Function

Function DistFbaP$()
DistFbaP = DistFba(SrcpzP(CPj))
End Function

Function DistFba$(Srcp)
DistFba = DistPjfzSrcp(Srcp, ".accdb")
End Function

Function DistPjfzSrcp(Srcp, Ext)
Dim P$:   P = Distp(Srcp)
Dim F1$: F1 = RplExt(Fdr(ParPth(P)), Ext)
Dim F2$: F2 = NxtFfnzNotIn(F1, PjfnAyV)
Dim F$:   F = NxtFfnzAva(P & F2)
DistPjfzSrcp = F
End Function

Function DistpP$() 'Distribution Path
DistpP = Distp(SrcpP)
End Function


Function DistFxazSrcp$(Srcp)
DistFxazSrcp = DistPjfzSrcp(Srcp, ".xlam")
End Function


