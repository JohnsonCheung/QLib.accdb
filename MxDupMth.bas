Attribute VB_Name = "MxDupMth"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDupMth."
Public Const FFoDupMth$ = FFoMthc

Function DupMthLy() As String()
Dim D As Drs: D = DoDupMthP
Dim Md$():   Md = AmAlignR(StrCol(D, "Mdn"))
Dim Mth$(): Mth = AmAlign(StrCol(D, "Mthn"))
Dim A$():     A = AyabJnDot(Md, Mth)
Dim MthLin$(): MthLin = StrCol(D, "MthLin")
                   DupMthLy = AyabJnSngQ(A, MthLin)
End Function

Function DupMthWsP() As Worksheet
Set DupMthWsP = DupMthWszP(CPj)
End Function

Function DoDupMthP(Optional InclPrv As Boolean, Optional IsExactDup As Boolean) As Drs
DoDupMthP = DoDupMthzP(CPj, InclPrv, IsExactDup)
End Function

Function DoDupMthzP(P As VBProject, Optional InclPrv As Boolean, Optional IsExactDup As Boolean) As Drs
Dim A As Drs: A = DwEQ(DoMthczP(P), "MdTy", "Std")
Dim A1 As Drs:
    If InclPrv Then
        A1 = A
    Else
        A1 = DwNe(A, "Mdy", "Prv")
    End If
Dim B As Drs: B = DwDup(A1, "Mthn")
Dim C As Drs: C = SrtDrs(B, "Mthn")
Dim D As Drs: D = AddColzValIdqCnt(C, "Mthl")
If IsExactDup Then
    DoDupMthzP = DwDup(D, "MthlId")
Else
    DoDupMthzP = D
End If
End Function

Function FmtDupMthWs(DupMthWs As Worksheet) As Worksheet
Dim Lo As ListObject: Set Lo = FstLo(DupMthWs)
SetLcWdt Lo, "MthL", 10
SetLcWrp Lo, "MthL", False
End Function

Function DupMthWszP(P As VBProject) As Worksheet
Set DupMthWszP = FmtDupMthWs(WszDrs(DoDupMthzP(P), "DupMth"))
End Function

Sub Z_DoDupMthP()
BrwDrs DoDupMthP
End Sub

