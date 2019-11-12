Attribute VB_Name = "MxMthl"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMthl."

Function MthLyzM(M As CodeModule, Mthn) As String()
MthLyzM = SplitCrLf(MthlzM(M, Mthn))
End Function

Function MthlzNmTy$(M As CodeModule, Mthn, ShtMthTy$)
Dim S$(): S = Src(M)
Dim Ix&: Ix = MthIxzNmTy(S, Mthn, ShtMthTy)
MthlzNmTy = MthlzIx(S, Ix)
End Function
Function MthlzPN$(P As VBProject, Mthn)
MthlzPN = MthlzM(MdzMthn(P, Mthn), Mthn)
End Function

Function MthLyzIx(Src$(), MthIx) As String()
Dim EIx&:       EIx = EndLix(Src, MthIx)
MthLyzIx = AwFT(Src, MthIx, EIx)
End Function
Function MthlzIx$(Src$(), MthIx)
MthlzIx = JnCrLf(MthLyzIx(Src, MthIx))
End Function

Function MthlzNm$(Src$(), Mthn)
Dim Ix, O$()
For Each Ix In Itr(MthIxyzN(Src, Mthn))
    PushI O, MthlzIx(Src, Ix)
Next
MthlzNm = JnDblCrLf(O)
End Function

Function MthlzSTN$(Src$(), ShtMthTy$, Mthn)
Dim Ix&: Ix = MthIxzNmTy(Src, Mthn, ShtMthTy)
MthlzSTN = MthlzIx(Src, Ix)
End Function

Function MthlzN$(Mthn)
MthlzN = MthlzPN(CPj, Mthn)
End Function

Function MthlzM$(M As CodeModule, Mthn)
MthlzM = MthlzNm(Src(M), Mthn)
End Function

Function MthLyM(Mthn) As String()
MthLyM = MthLyzM(CMd, Mthn)
End Function

