Attribute VB_Name = "MxDcl"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDcl."
Public Const FFoDcl$ = "Pjn Mdn Dcll Dcll1"

Function FstDclLno&(M As CodeModule)

End Function
Function DclItmAyzDimLin(DimLin) As String()
Dim L$: L = DimLin
If Not ShfPfx(L, "Dim ") Then Exit Function
DclItmAyzDimLin = SplitCommaSpc(L)
End Function

Function DclNy(DcitmAy$()) As String()
Dim Dcitm
For Each Dcitm In Itr(DcitmAy)
    PushI DclNy, Dcn(Dcitm)
Next
End Function

Function NTyMd%(M As CodeModule)
NTyMd = NTySrc(DclzM(M))
End Function

Function EnmBdyLyzS(Src$(), Enmn) As String()
EnmBdyLyzS = EnmBdyLy(EnmLy(Src, Enmn))
End Function

Function EnmBdyLy(EnmLy$()) As String()

End Function

Function EnmFei(Src$(), Enmn) As Fei
Dim Fm&: Fm = EnmFmIx(Src, Enmn)
EnmFei = Fei(Fm, EndLix(Src, Fm))
End Function

Function EnmLy(Src$(), Enmn) As String()
EnmLy = AwFei(Src, EnmFei(Src, Enmn))
End Function

Function EnmFmIx&(Src$(), Enmn)
Dim J&, L, Lin$
For Each L In Itr(Src)
    Lin = RmvMdy(L)
    If ShfTermEnm(Lin) Then
        If Nm(Lin) = Enmn Then
            EnmFmIx = J
            Exit Function
        End If
    End If
    If IsLinMth(Lin) Then Exit For
    J = J + 1
Next
EnmFmIx = -1
End Function

Function EnmNyMd(M As CodeModule) As String()
EnmNyMd = EnmNy(DclzM(M))
End Function
Function EnmNyPj(Pj As VBProject) As String()
Dim M
For Each M In MdItr(Pj)
    PushIAy EnmNyPj, EnmNyMd(CvMd(M))
Next
End Function
Function EnmNy(Src$()) As String()
Dim L
For Each L In Itr(Src)
   PushNB EnmNy, Enmn(L)
Next
End Function

Function HasTyn(Src$(), Nm$) As Boolean
Dim L
For Each L In Itr(Src)
    If Mtyn(L) = Nm Then HasTyn = True: Exit Function
Next
End Function

Function NEnm%(Src$())
Dim L, O%
For Each L In Itr(Src)
   If IsLinEmn(L) Then O = O + 1
Next
NEnm = O
End Function

Function TyFei(Dcl$(), TyNm$) As Fei
Dim FmI&: FmI = TyFmIx(Dcl, TyNm)
Dim ToI&: ToI = EndLix(Dcl, FmI)
TyFei = Fei(FmI, ToI)
End Function
Function MthELno&(M As CodeModule, MthLno&)
Dim MLin$: MLin = M.Lines(MthLno, 1)
Dim ELin$: ELin = MthELin(MLin)
Dim O&: For O = MthLno + 1 To M.CountOfLines
    If HasPfx(M.Lines(O, 1), ELin) Then MthELno = O: Exit Function
Next
End Function
Function MthELin$(MthLin)
Dim K$: K = MthKd(MthLin)
If K = "" Then Thw CSub, "MthLin Error", "MthLin", MthLin
MthELin = "End " & K
End Function

Function TyLines$(Dcl$(), Mtyn$)
TyLines = JnCrLf(TyLy(Dcl, Mtyn))
End Function

Function TyLy(Dcl$(), TyNm$) As String()
TyLy = AwFei(Dcl, TyFei(Dcl, TyNm))
End Function

Function TyFmIx&(Src$(), TyNm)
Dim J%
For J = 0 To UB(Src)
   If IsLinTy(Src(J)) = TyNm Then TyFmIx = J: Exit Function
   If IsLinMth(Src(J)) Then Exit For
Next
TyFmIx = -1
End Function

Function TyNy(DclLy$()) As String()
Dim L
For Each L In Itr(DclLy)
    PushNB TyNy, Mtyn(L)
Next
End Function
Function Enmn(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfPfx(L, "Enum ") Then Enmn = Nm(L)
End Function

Function Mtyn$(Lin)
':Mtyn: :Nm #Type-Name# ! Vb Type Name of @Lin
Dim L$: L = RmvMdy(Lin)
If ShfPfx(L, "Type ") Then Mtyn = Nm(L)
End Function

Function EnmLyzMN(M As CodeModule, Enmn) As String()
EnmLyzMN = EnmLy(DclzM(M), Enmn)
End Function

Function NEnmMbrzMN%(M As CodeModule, Enmn)
NEnmMbrzMN = Si(EnmMbrLyzMN(M, Enmn))
End Function

Function CdLyzL(Lin) As String()
Dim L$: L = Trim(Lin)
If L = "" Then Exit Function
If FstChr(L) = "'" Then Exit Function
CdLyzL = SyzTrim(Split(Lin, ":"))
End Function
Sub Z_Srcc()
Brw Srcc(SrczP(CPj))
End Sub

Function EnmMbrLyzMN(M As CodeModule, Enmn) As String()
EnmMbrLyzMN = Srcc(EnmLyzMN(M, Enmn))
End Function

Function NEnmzM%(M As CodeModule)
NEnmzM = NEnm(DclzM(M))
End Function

Function TyNyzM(M As CodeModule) As String()
TyNyzM = TyNy(DclzM(M))
End Function

Function TyNyP() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = TyNyzP(CPj)
End If
TyNyP = Y
End Function

Function TyNyzP(P As VBProject) As String()
Dim I, C As VBComponent
For Each C In P.VBComponents
    PushIAy TyNyzP, TyNyzM(C.CodeModule)
Next
End Function

Function ShfTermEnm(OLin$) As Boolean
ShfTermEnm = ShfPfx(OLin, "Enum")
End Function


Sub Z_NEnmMbrzMN()
Ass NEnmMbrzMN(Md("Ide"), "AA") = 1
End Sub

Sub BrwNDclLinP()
BrwDy DyoNDclLin(CPj)
End Sub


Function IxoPrvCdLin&(Src$(), Fm)
Dim O&
For O = Fm - 1 To 0 Step -1
    If IsLinCd(Src(O)) Then IxoPrvCdLin = O: Exit Function
Next
IxoPrvCdLin = -1
End Function
Function DclDicP() As Dictionary
Set DclDicP = DclDiczP(CPj)
End Function

Function DclDiczP(P As VBProject) As Dictionary
If P.Protection = vbext_pp_locked Then Set DclDiczP = New Dictionary: Exit Function
Dim C As VBComponent, M As CodeModule
Set DclDiczP = New Dictionary
For Each C In P.VBComponents
    Set M = C.CodeModule
    Dim Dcl$(): Dcl = DclzM(M)
    If Si(Dcl) = 0 Then
        DclDiczP.Add MdDn(M), Dcl
    End If
Next
End Function

Function DclP() As String()
DclP = DclzP(CPj)
End Function

Function DclzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy DclzP, DclzM(C.CodeModule)
Next
End Function

Function DclItr(M As CodeModule)
':DclItr: :LinItr #Dcl-Lin-Itr#
Asg Itr(DclzM(M)), DclItr
End Function

Function Dcll$(Src$())
':Dcll: :Lines ! comes fm a module
Dcll = JnCrLf(Dcl(Src))
End Function

Function HasDclLin(M As CodeModule, DclLin$) As Boolean
Dim J&: For J = 1 To M.CountOfDeclarationLines
    If M.Lines(J, 1) = DclLin Then HasDclLin = True: Exit Function
Next
End Function

Function Dcl(Src$()) As String()
If Si(Src) = 0 Then Exit Function
Dim N&: N = NDclLinzS(Src)
If N <= 0 Then Exit Function
Dcl = FstNEle(Src, N)
End Function

Function MdlzN$(P As VBProject, Mdn)
MdlzN = Mdl(P.VBComponents(Mdn))
End Function

Function MdlAy(P As VBProject, MdNy$()) As String()
Dim N: For Each N In Itr(MdNy)
    PushI MdlAy, MdlzN(P, N)
Next
End Function

Function Mdl$(M As CodeModule)
Dim Cnt&: Cnt = M.CountOfLines
If Cnt = 0 Then Exit Function
Mdl = M.Lines(1, Cnt)
End Function

Function MdLines$(M As CodeModule, Lno&, Cnt&)
If Lno <= 0 Then Exit Function
If Cnt <= 0 Then Exit Function
If Lno > M.CountOfLines Then Exit Function
MdLines = M.Lines(Lno, Cnt)
End Function

Sub Z_DclzM()
Dim O$(), C As VBComponent
For Each C In CPj.VBComponents
    PushNB O, DclzM(C.CodeModule)
Next
VcLinesAy O
End Sub
Function DcllzM$(M As CodeModule)
DcllzM = JnCrLf(DclzM(M))
End Function

Function DclzM(M As CodeModule) As String()
Dim N&: N = M.CountOfDeclarationLines
If N > 0 Then
    DclzM = EndTrimVbRmk(SplitCrLf(M.Lines(1, N)))
End If
End Function
