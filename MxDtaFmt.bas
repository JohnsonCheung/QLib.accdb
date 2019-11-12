Attribute VB_Name = "MxDtaFmt"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaFmt."

Function AlignAy(Ay, Optional W0%) As String()
Dim W%: If W0 <= 0 Then W = WdtzAy(Ay) Else W = W0
Dim S: For Each S In Itr(Ay)
    PushI AlignAy, AlignL(S, W)
Next
End Function


Function AlignQteSq(Fny$()) As String()
AlignQteSq = AlignAy(SyzQteSq(Fny))
End Function

Function AlignRzAy(Ay, Optional W0%) As String() 'Fmt-Dr-ToWdt
Dim W%: If W0 <= 0 Then W = WdtzAy(Ay) Else W = W0
Dim I
For Each I In Itr(Ay)
    PushI AlignRzAy, AlignR(I, W)
Next
End Function

Function DyzSSAy(SSAy$()) As Variant()
Dim Dr: For Each Dr In Itr(SSAy)
    PushI DyzSSAy, SyzSS(Dr)
Next
End Function

Function AlignSSAy(SSAy$()) As String()
Dim Dy(): Dy = DyzSSAy(SSAy)
Dim Dy1(): Dy1 = AlignDy(Dy)
AlignSSAy = JnDy(Dy1)
End Function

Function AlignLyzSepss(Ly$(), SepSS$) As String()
AlignLyzSepss = JnDy(AlignDy(DyoSySep(Ly, SyzSS(SepSS))))
End Function

Function BrkLin(Lin, Sep$(), Optional IsRmvSep As Boolean) As String()
'Ret : seg ay of a lin sep by @Sep.  Si of seg ret = si of @sep + 1.  Each will have its own sep, expt fst.
'      Segs are not trim and wi/wo by @IsRmvSep.  If not @IsRmvSep, Jn(@Rslt) will eq @Lin @@
Dim L$: L = Lin
Dim O$()
Dim S: For Each S In Sep
    PushI O, ShfBef(L, S)
Next
PushI O, L
If IsRmvSep Then
    Dim J&, Seg: For Each Seg In O
        PushI BrkLin, RmvPfx(Seg, Sep(J))
        J = J + 1
    Next
Else
    BrkLin = O
End If
End Function


Sub BrwDrs(A As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EmIxCol.EiBeg1, _
Optional Fmt As EmTblFmt = EiSSFmt, _
Optional FnPfx$, Optional OupTy As EmOupTy = EmOupTy.EiOtBrw)
BrwAy FmtDrs(A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt), FnPfx, OupTy
End Sub

Sub BrwDrs2(A As Drs, B As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EiBeg1, _
Optional Fmt As EmTblFmt = EiTblFmt, Optional NN$, Optional Tit$ = "Brw 2 Drs", _
Optional FnPfx$, Optional OupTy As EmOupTy = EmOupTy.EiOtBrw)
Dim Ay$(), AyA$(), AyB$(), N1$, N2$, T$()
N1 = DftStr(BefSpc(NN), "Drs-A")
N2 = DftStr(AftSpc(NN), " Drs-B")
AyA = FmtDrszNoRdu(A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N1)
AyB = FmtDrszNoRdu(B, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N2)
T = Sy(Tit, ULinDbl(Tit))
Ay = Sy(T, AyA, AyB)
Brw Ay, FnPfx, OupTy:=OupTy
End Sub

Sub BrwDrs3(A As Drs, B As Drs, C As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EmIxCol.EiBeg1, _
Optional Fmt As EmTblFmt = EiTblFmt, Optional NN$, Optional Tit$ = "Brw 3 Drs", _
Optional FnPfx$, Optional OupTy As EmOupTy = EmOupTy.EiOtBrw)
Dim Ay$(), AyA$(), AyB$(), AyC$(), N1$, N2$, N3$, T$()
N1 = DftStr(T1(NN), "Drs-A")
N2 = DftStr(T2(NN), " Drs-B")
N3 = DftStr(RmvTT(NN), " Drs-C")
AyA = FmtDrszNoRdu(A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N1)
AyB = FmtDrszNoRdu(B, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N2)
AyC = FmtDrszNoRdu(C, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N3)
T = Sy(Tit, ULinDbl(Tit))
Ay = Sy(T, AyA, AyB, AyC)
Brw Ay, FnPfx, OupTy:=OupTy
End Sub

Sub BrwDrs4(A As Drs, B As Drs, C As Drs, D As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EiBeg1, _
Optional Fmt As EmTblFmt = EiTblFmt, _
Optional FnPfx$, Optional OupTy As EmOupTy = EmOupTy.EiOtBrw)
Dim Ay$(), AyA$(), AyB$(), AyC$(), AyD$()
AyA = FmtDrszNoRdu(A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt)
AyB = FmtDrszNoRdu(B, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt)
AyC = FmtDrszNoRdu(C, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt)
AyD = FmtDrszNoRdu(D, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt)
Ay = Sy(AyA, AyB, AyC, AyD)
Brw Ay, FnPfx, OupTy:=OupTy
End Sub

Sub BrwDy(A(), Optional MaxColWdt% = 100, Optional BrkCCIxy, Optional ShwZer As Boolean, Optional Fmt As EmTblFmt = EmTblFmt.EiTblFmt)
BrwAy FmtDy(A, MaxColWdt, BrkCCIxy, ShwZer, Fmt)
End Sub

Sub BrwDyoSpc(A(), Optional MaxColWdt% = 100, Optional BrkCCIxy, Optional ShwZer As Boolean)
BrwAy FmtDy(A, MaxColWdt, BrkCCIxy, ShwZer, Fmt:=EiSSFmt)
End Sub

Function Cell$(V, Optional ShwZer As Boolean, Optional MaxWdt0% = 30)
':Cell: :SCell-or-WCell
':SCell: :S      ! can fill in a cell without wrap
':WCell: :Lines  ! can fill in a cell with wrap
Dim O$, W%: W = EnsBet(MaxWdt0, 1, 1000)
Select Case True
Case IsLines(V):   O = CellzLines(V, W)
Case IsStr(V):     O = CellzS(V, W)
Case IsBool(V):    O = IIf(V, "True", "")
Case IsNumeric(V): O = CellzN(V, W, ShwZer)
Case IsPrim(V):    O = V
Case IsEmp(V):     O = "#Emp#"
Case IsNull(V):    O = "#Null#"
Case IsArray(V):   O = Cell = "*[" & Si(V) & "]"
Case IsDic(V):     O = "#Dic:Cnt(" & CvDic(V).Count & ")"
Case IsObject(V):  O = "#O:" & TypeName(V)
Case IsErObj(V)
Case Else:         O = V
End Select
Cell = O
End Function

Function CellzLines$(Lines, W%)
'Ret : each lin in @Lines will be cut to @W and jn it back
Dim O$(), S: For Each S In Itr(SplitCrLf(Lines))
    PushI O, CellzS(S, W)
Next
CellzLines = JnCrLf(O)
End Function

Function CellzN$(N, MaxW%, ShwZer As Boolean)
Select Case True
Case N = 0: If ShwZer Then CellzN = "0"
Case Else:  CellzN = N
End Select
End Function

Function CellzS$(S, W%)
CellzS = SlashCrLf(Left(S, W))
End Function

Sub DmpDy(Dy(), _
Optional MaxColWdt% = 100, _
Optional BrkCCIxy0, _
Optional ShwZer As Boolean, _
Optional Fmt As EmTblFmt)
D FmtDy(Dy, MaxColWdt, BrkCCIxy0, ShwZer, Fmt)
End Sub

Sub DmpDyoSpc(Dy(), _
Optional MaxColWdt% = 100, _
Optional BrkCCIxy0, _
Optional ShwZer As Boolean)
D FmtDy(Dy, MaxColWdt, BrkCCIxy0, ShwZer, Fmt:=EiSSFmt)
End Sub

Function DrszFmtg(DrsFmtg$()) As Drs
Dim TitLin$: TitLin = DrsFmtg(1)
Dim Fny$(): Fny = AeFstLas(SyzTrim(Split(TitLin, "|")))
Dim Dy()
    Dim J&
    For J = 3 To UB(DrsFmtg) - 1
        PushI Dy, AvzAy(AeFstLas(RmvFstChrzAy(RSyzTrim(Split(DrsFmtg(J), "|")))))
    Next
DrszFmtg = Drs(Fny, Dy)
End Function

Function DyoInsIx(Dy()) As Variant()
' Ret Dy with each row has ix run from 0..{N-1} in front
Dim Ix&, Dr: For Each Dr In Itr(Dy)
    Dr = InsEle(Dr, Ix)
    PushI DyoInsIx, Dr
    Ix = Ix + 1
Next
End Function

Function DyoSySep(Sy$(), Sep$()) As Variant()
'Ret : a dry wi each rec as a sy of brkg one lin of @Sy.  Each lin is brk by @Sep using fun-BrkLin @@
Dim I, Lin
For Each I In Itr(Sy)
    Lin = I
    PushI DyoSySep, BrkLin(Lin, Sep)
Next
End Function

Function DyoSySepss(Ly$(), SepSS$) As Variant()
DyoSySepss = DyoSySep(Sy, SyzSS(SepSS))
End Function

Function FmtDrszNoRdu(D As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EmIxCol.EiBeg1, _
Optional Fmt As EmTblFmt = EiTblFmt, Optional Nm$) As String()
'Fm IsSum    : If true all num col will have a sum as las lin in the fmt
'Fm BrkColnn : if changed, insert a break line if BrkColNm is given
Dim NmBox$(): If Nm <> "" Then NmBox = Box(Nm)
If NoReczDrs(D) Then FmtDrszNoRdu = MsgNoRec(D, NmBox): Exit Function
Dim IxD As Drs:    IxD = AddColzIx(D, IxCol)                     ' Add Col-Ix
Dim IxyB&():      IxyB = Ixy(IxD.Fny, TermAy(BrkColnn))          ' Ixy-Of-BrkCol
Dim Dy():           Dy = AddEle(IxD.Dy, IxD.Fny)                 ' Dy<Bdy-Fny-Sep>
Dim Bdy$():        Bdy = FmtDy(Dy, MaxColWdt, IxyB, ShwZer, Fmt) ' Ly<Bdy-Fny-Sep-?Sum>
Dim Sep$:          Sep = Pop(Bdy)                                ' Sep-Lin
Dim Hdr$:          Hdr = Pop(Bdy)                                ' Hdr-Lin
Dim O$():            O = Sy(NmBox, Sep, Hdr, Bdy, Sep)
                FmtDrszNoRdu = O
End Function

Function MsgNoRec(D As Drs, NmBox$()) As String()
Dim S$:        S = JnSpc(D.Fny)
Dim S1$:           If S = "" Then S1 = " (No Fny)" Else S1 = S
Dim Lin$:    Lin = "(NoRec) " & S1
   MsgNoRec = Sy(NmBox, Lin)
End Function

Function FmtDt(A As Dt, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EiBeg1) As String()
PushI FmtDt, "*Tbl " & A.DtNm
PushIAy FmtDt, FmtDrszNoRdu(DrszDt(A), MaxColWdt, BrkColNm, ShwZer, IxCol)
End Function

Function IsEqAyzIxy(A, B, Ixy&()) As Boolean
Dim J%
For J = 0 To UB(Ixy)
    If A(Ixy(J)) <> B(Ixy(J)) Then Exit Function
Next
IsEqAyzIxy = True
End Function
Function JnDySpc(Dy()) As String()
JnDySpc = JnDy(Dy, " ")
End Function
Function JnDyDot(Dy()) As String()
JnDyDot = JnDy(Dy, ".")
End Function
Function JnDy(Dy(), Optional Sep$ = " ") As String()
'Ret: :Ly by joining each :Dr in @Dy by @Sep
Dim Dr: For Each Dr In Itr(Dy)
    PushI JnDy, Jn(Dr, Sep)
Next
End Function

Function LinzDr(Dr, Optional Sep$ = " ", Optional QteStr$)
'Ret : ret a lin from Dr-QteStr-Sep
LinzDr = Qte(Jn(Dr, Sep), QteStr)
End Function

Function LinzDrsR(A As Drs, Optional Nm$) As String()
If NoReczDrs(A) Then Exit Function
Dim AFny$(): AFny = Sy("#", AlignAy(A.Fny))

Dim Ly$(), Lixy&()
    Dim N&: N = Si(A.Dy)
    Dim Dr, J&: For Each Dr In Itr(A.Dy)
        J = J + 1
        PushI Ly, Empty
        PushI Lixy, UB(Ly)
        Dim I$: I = J & " of " & N
        Dim Av(): Av = AddAy(Array(I), Dr)
        PushIAy Ly, LyzNyAv(AFny, Av)
    Next
Dim Align$(): Align = AlignAy(Ly)
Dim Q$()
    Dim L: For Each L In Align
        Push Q, "| " & L & " |"
    Next
'== Oup ===
Dim O$(): O = Q
Dim W%:   W = Len(Align(0))
Dim Lin$:   Lin = "|-" & Dup("-", W) & "-|"
Dim Ix: For Each Ix In Itr(Lixy)
    O(Ix) = Lin
Next
PushI O, Lin
LinzDrsR = O
End Function

Function LinzSep$(W%())
LinzSep = LinzDr(SepDr(W), "-|-", "|-*-|")
End Function

Function SepDr(W%()) As String()
Dim I: For Each I In W
    Push SepDr, Dup("-", I)
Next
End Function

Function SslSyzDy(Dy()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    Push SslSyzDy, SslzDr(Dr) ' Fmtss(X)
Next
End Function

Sub VcDrs(A As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EiBeg1, _
Optional Fmt As EmTblFmt, _
Optional FnPfx$, Optional UseVc As Boolean)
BrwDrs A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, FnPfx, OupTy:=EiOtVc
End Sub
Function WdtAyzDy_AllCol(Dy()) As Integer()
':CellDy: :Dy ! Each cell is a Str or Lines
Dim J&
For J = 0 To NColzDy(Dy) - 1
    Push WdtAyzDy_AllCol, WdtzAy(StrColzDy(Dy, J))
Next
End Function

Function WdtAyzDy_FstNCol(Dy(), FstNCol%) As Integer()
':CellDy: :Dy ! Each cell is a Str or Lines
Dim J&, O%()
ReDim O(FstNCol - 1)
For J = 0 To FstNCol - 1
    O(J) = WdtzAy(StrColzDy(Dy, J))
Next
End Function
Function WdtAyzDy(Dy(), Optional FstNCol%) As Integer()
If FstNCol <= 0 Then
    WdtAyzDy = WdtAyzDy_AllCol(Dy)
Else
    WdtAyzDy = WdtAyzDy_FstNCol(Dy, FstNCol)
End If
End Function


Sub Z_DyoSySepss()
Dim Ly$(), Sep$
GoSub T0
Exit Sub
T0:
    Sep = ". . . . . ."
    Ly = Sy("AStkShpCst_Rpt.OupFx.Fun.")
    Ept = Sy("AStkShpCst_Rpt", ".OupFx", ".Fun", ".")
    GoTo Tst
Tst:
    BrwDy DyoSySepss(Sy, Sep)
    C
    Return
End Sub

Sub Z_FmtDrszNoRdu()
Dim A As Drs, MaxColWdt%, BrkColVbl$, ShwZer As Boolean, IxCol As EmIxCol
GoSub Z
Exit Sub
T1:
    A = SampDrs
    GoSub Tst
Tst:
    Act = FmtDrszNoRdu(A, MaxColWdt, BrkColVbl, ShwZer, IxCol)
    Brw Act: Stop
    C
    Return
Z:
    DmpAy FmtDrszNoRdu(SampDrs1)
    Return
End Sub

Sub Z_FmtDt()
Dim A As Dt, MaxColWdt%, BrkColNm$, ShwZer As Boolean
'--
A = SampDt1
'Ept = Z_TimStrpt1
GoSub Tst
'--
Exit Sub
Tst:
    Act = FmtDt(A, MaxColWdt, BrkColNm, ShwZer)
    C
    Return
End Sub
