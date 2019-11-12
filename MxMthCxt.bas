Attribute VB_Name = "MxMthCxt"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMthCxt."
Type Fc
    FmLno As Long
    Cnt As Long
End Type
Type Fcs: N As Long: Ay() As Fc: End Type
Public Const FFoMthn$ = "MdTy Mthn Mdn Mdy Ty"

Sub UnRmkMth(M As CodeModule, Mthn)
UnRmkMdzFcs M, MthCxtFcs(Src(M), Mthn)
End Sub

Function Fc(FmLno, Cnt) As Fc
If Cnt <= 0 Then Exit Function
If FmLno <= 0 Then Exit Function
Fc.FmLno = FmLno
Fc.Cnt = Cnt
End Function

Sub RmkMdeszFc(M As CodeModule, Fc As Fc)
RmkNLin M, Fc.FmLno, Fc.Cnt
End Sub

Function MthCxtFcs(Src$(), Mthn) As Fcs
MthCxtFcs = MthCxtFcszzSM(Src, MthCxtFcs(Src, Mthn))
End Function

Function MthCxtFcszzSM(Src$(), Mth As Fcs) As Fcs
Dim J&
For J = 0 To Mth.N - 1
    PushFc MthCxtFcszzSM, MthCxtFczzSM(Src, Mth.Ay(J))
Next
End Function

Function NContLin(Src$(), MthIx) As Byte
Dim J&, O&
For J = MthIx To UB(Src)
    O = O + 1
    If LasChr(Src(J)) <> "_" Then NContLin = O: Exit Function
Next
Thw CSub, "LasEle of Src has LasChr = _", "Src", Src
End Function

Function AddFc(A As Fc, B As Fc) As Fcs
PushFc AddFc, A
PushFc AddFc, B
End Function

Function FmtFcs$(A As Fcs)
Dim O$(), J&
For J = 0 To A.N - 1
    PushI O, FmtFc(A.Ay(J))
Next
FmtFcs = JnCrLf(O)
End Function

Function FmtFc$(Fc As Fc)
With Fc
FmtFc = "Fc " & .FmLno & " " & .Cnt
End With
End Function

Function MthCxtFczzSM(Src$(), Mth As Fc) As Fc
With Mth
Dim N%: N = NContLin(Src, .FmLno)
MthCxtFczzSM = Fc(.FmLno - N, .Cnt - N - 1)
End With
End Function
Function MthCxtFc(M As CodeModule, Mthn) As Fc

End Function
Sub RmkMthzN(M As CodeModule, Mthn)
RmkMdeszFc M, MthCxtFc(M, Mthn)
End Sub

Function MthFcs(M As CodeModule, Mthn) As Fcs
Dim Ix, S$()
S = Src(M)
For Each Ix In Itr(MthIxyzN(S, Mthn))
    PushFc MthFcs, Fc(Ix + 1, ContLinCnt(S, Ix))
Next
End Function

Sub PushFc(O As Fcs, M As Fc)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Sub RmkNLin(M As CodeModule, Lno&, N&)
Dim J&
For J = Lno To Lno + N - 1
    RmkLin M, J
Next
End Sub

Sub RmkLin(M As CodeModule, Lno&)
M.ReplaceLine M, "'" & M.Lines(Lno, 1)
End Sub

Sub RmkMth()
RmkMthzN CMd, CMthn
End Sub

Sub Z_RmkMth()
Dim Md As CodeModule, Mthn
'            Ass LineszVbl(MthL(M)) = "Property Get ZZA()|End Property||Sub SetYYA(V)||End Property"
'RmkMth M:   Ass LineszVbl(MthL(M)) = "Property Get ZZA()|Stop '|End Property||Sub SetYYA(V)|Stop '|'|End Property"
'UnRmkMth M: Ass LineszVbl(MthL(M)) = "Property Get ZZA()|End Property||Sub SetYYA(V)||End Property"
End Sub
Function NxtMdLno(M As CodeModule, Lno)
Const CSub$ = CMod & "NxtMdLno"
Dim J&
For J = Lno To M.CountOfLines
    If LasChr(M.Lines(Lno, 1)) <> "_" Then
        NxtMdLno = J
        Exit Function
    End If
Next
Thw CSub, "All line From Lno has _ as LasChr", "Lno Md Src", Lno, Mdn(M), AddIxPfx(Src(M), 1)
End Function

Sub UnRmkMdzFcs(M As CodeModule, B As Fcs)
Dim J&: For J = 0 To B.N - 1
    UnRmkMdzFc M, B.Ay(J)
Next
End Sub

Sub UnRmkMdzFc(M As CodeModule, B As Fc)
'If Not IsRmkzS(LyzMdFei(A, B)) Then Exit Sub
Stop
Dim J%, L$
'For J = NxtMdLno(A, B.FmNo) To B.ToNo - 1
    L = M.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    M.ReplaceLine J, Mid(L, 2)
'Next
End Sub

Sub RmkMdzFcs(M As CodeModule, B As Fcs)
Dim J%
For J = 0 To B.N - 1
    RmkMdzFc M, B.Ay(J)
Next
End Sub

Sub RmkMdzFc(M As CodeModule, B As Fc)
If IsRmkzFc(M, B) Then Exit Sub
Dim J&: For J = 0 To B.Cnt - 1
    M.ReplaceLine J, "'" & M.Lines(B.FmLno + J, 1)
Next
End Sub

Function IsRmkzFc(M As CodeModule, B As Fc) As Boolean
IsRmkzFc = IsRmkzS(SrczFc(M, B))
End Function

Function DoMthCxt() As Drs
DoMthCxt = DoMthCxtzML(CMd, CMthLno)
End Function


Function DoMthnzM(M As CodeModule) As Drs
DoMthnzM = DwEQ(SelDrs(DoMthnP, FFoMthn), "Mthn", Mdn(M))
End Function

Function AddColzBetBkt(D As Drs, ColnAs$, Optional IsDrp As Boolean) As Drs
Dim BetColn$, NewC$: AsgBrk1 ColnAs, ":", BetColn, NewC
If NewC = "" Then NewC = BetColn & "InsideBkt"
Dim Ix%: Ix = IxzAy(D.Fny, BetColn)
Dim Dr, Dy(): For Each Dr In Itr(D.Dy)
    PushI Dr, BetBkt(Dr(Ix))
    PushI Dy, Dr
Next
Dim O As Drs: O = AddColzFFDy(D, NewC, Dy)
If IsDrp Then O = DrpCol(O, BetColn)
AddColzBetBkt = O
End Function


Function IsRetObj(RetSfx$) As Boolean
':IsRetObj: :B ! False if @RetSfx (isBlnk | IsAy | IsPrimTy | Is in TyNyP)
If RetSfx = "" Then Exit Function
If HasSfx(RetSfx, "()") Then Exit Function
If IsPrimTy(RetSfx) Then Exit Function
If HasEle(TyNyP, RetSfx) Then Exit Function
IsRetObj = True
End Function

Function AddColzRetAs(DoMthLin As Drs) As Drs
'Fm DoMthLin : ..MthLin..
'Ret        : ..RetSfx  @@
Dim IxMthLin%: IxMthLin = IxzAy(DoMthLin.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In Itr(DoMthLin.Dy)
    Dim MthLin$: MthLin = Dr(IxMthLin)
    Dim R$: R = RetSfx(MthLin)
    PushI Dr, R
    PushI Dy, Dr
Next
AddColzRetAs = AddColzFFDy(DoMthLin, "RetSfx", Dy)
End Function

Function DoMthCxtzML(M As CodeModule, MthLno&) As Drs
'Ret DoMthCxt : L Lin
Dim Dy(), L&, ELin$, MthLin$, Lin$
MthLin = M.Lines(MthLno, 1)
ELin = MthELin(MthLin)
For L = NxtLnozML(M, MthLno) To M.CountOfLines
    Lin = M.Lines(L, 1)
    If Lin = ELin Then
        GoTo X
    End If
    Lin = M.Lines(L, 1)
    PushI Dy, Array(L, Lin)
Next
ThwImpossible CSub
X:
DoMthCxtzML = DrszFF("L MthLin", Dy)
End Function

Function IsRmkzMthLy(MthLy$()) As Boolean
If Si(MthLy) = 0 Then Exit Function
If Not HasPfx(MthLy(0), "Stop '") Then Exit Function
Dim L
For Each L In MthLy
    If Left(L, 1) <> "'" Then Exit Function
Next
IsRmkzMthLy = True
End Function
Function MthCxtFe(MthLy$(), Fe As Fei) As Fei
MthCxtFe = Fei(NxtIxzSrc(MthLy, Fe.FmIx), Fe.EIx - 1)
End Function

Function MthCxt$(MthLy$())
MthCxt = JnCrLf(MthCxtLy(MthLy))
End Function

Function MthCxtLy(MthLy$()) As String()
If Si(MthLy) = 0 Then Exit Function
Dim L&: L = FstMthIx(MthLy): If L = -1 Then Thw CSub, "Given MthLy is not MthLy", "MthLy", MthLy
Dim J%
For J = NxtIxzSrc(MthLy, L) To UB(MthLy) - 1
    PushI MthCxtLy, MthLy(J)
Next
End Function

Sub Z_MthCxtFeis _
 _
()
Stop
Dim I
'For Each I In MthCxtFeis(CSrc, CurMthn)
    'With CvFei(I)
'        Debug.Print .FmNo, .ToNo
    'End With
'Next
End Sub
