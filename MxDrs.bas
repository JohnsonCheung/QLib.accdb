Attribute VB_Name = "MxDrs"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDrs."
Const CNs$ = "Dta.Ds"
Enum EmTblFmt
    EiTblFmt = 0
    EiSSFmt = 1
End Enum
Type Drs: Fny() As String: Dy() As Variant: End Type
Type Drss: N As Long: Ay() As Drs: End Type
Enum EmCnt
    EiCntAll
    EiCntDup
    EiCntSng
End Enum
Enum EmCntSrtOpt
    eNoSrt
    eSrtByCnt
    eSrtByItm
End Enum
Private Type GpDrs
    GpDrs As Drs
    RLvlGpIx() As Long
End Type

Function AddColzCol(A As Drs, C$, Col) As Drs
Dim Dy(): Dy = A.Dy
If Si(Dy) <> Si(Col) Then Thw CSub, "@Drs.Dy and @Col Should be same size", "Drs-Dy-Si Col-Si", Si(Dy), Si(Col)
Dim ODy()
    Dim Dr, J&: For Each Dr In Itr(A.Dy)
        PushI Dr, Col(J)
        PushI ODy, Dr
        J = J + 1
    Next
AddColzCol = AddColzFFDy(A, C, ODy)
End Function

Function CpyCol(A As Drs, FmColn$, AsColn$) As Drs
Dim Col1(): Col1 = Col(A, FmColn)
CpyCol = AddColzCol(A, AsColn, Col1)
Stop
End Function


Function AddColzExiB(A As Drs, b As Drs, Jn$, ExiB_FldNm$) As Drs
Dim IA&(), IB&(), Dr, KA(), BKeyDy(), ODy()
IA = IxyzJnA(A.Fny, Jn)
IB = IxyzJnB(b.Fny, Jn)
BKeyDy = SelDy(b.Dy, IB)
For Each Dr In Itr(A.Dy)
    KA = AwIxy(Dr, IA)
    If HasDr(BKeyDy, KA) Then
        PushI Dr, True
    Else
        PushI Dr, False
    End If
    PushI ODy, Dr
Next
AddColzExiB = Drs(AddSS(A.Fny, ExiB_FldNm), ODy)
End Function

Function AddColzFFDy(A As Drs, FF$, NewDy()) As Drs
AddColzFFDy = Drs(AddSS(A.Fny, FF), NewDy)
End Function

Function AddColzFiller(A As Drs, CC$) As Drs
Dim O As Drs: O = A
Dim C
For Each C In SyzSS(CC)
    O = AddColzFillerC(O, C)
Next
AddColzFiller = O
End Function

Function AddColzFillerC(A As Drs, C) As Drs
'@  A : ..{C}.. ! @A should have col @C
'@  C : #Coln.
'Ret    : a new drs with addition col @F where F = "F" & C and value eq Len-of-Col-@C
If NoReczDrs(A) Then Stop
Dim W%: W = WdtzAy(StrCol(A, C))
Dim I%: I = IxzAy(A.Fny, C)
Dim ODy(): ODy = A.Dy
Dim Dr, J&
For Each Dr In Itr(ODy)
    PushI Dr, W - Len(Dr(I))
    ODy(J) = Dr
    J = J + 1
Next
AddColzFillerC = Drs(AddSS(A.Fny, "F" & C), ODy)
End Function

Function AddColzIx(D As Drs, IxCol As EmIxCol) As Drs
Dim J&, Fny$(), Dy(), I, Dr
Select Case True
Case IxCol = EiNoIx: AddColzIx = D: Exit Function
Case IxCol = EiBeg0
Case IxCol = EiBeg1: J = 1
End Select
Fny = InsEle(D.Fny, "Ix")
For Each Dr In Itr(D.Dy)
    Push Dy, InsEle(Dr, J)
    J = J + 1
Next
AddColzIx = Drs(Fny, Dy)
End Function

Function AddColzLen(D As Drs, AsCol$) As Drs
'Fm AsCol : If no as, {Col}Len will be used
'Ret      : add a len col at end using LenCol @@
Dim C$:       C = BefOrAll(AsCol, ":")
Dim LenC$: LenC = AftOrAll(AsCol, ":")
                  If LenC = C Then LenC = C & "Len"
Dim Ix&: Ix = IxzAy(D.Fny, C)
Dim Dy(), Dr: For Each Dr In Itr(D.Dy)
    Dim L%: L = Len(Dr(Ix))
    PushI Dr, L
    PushI Dy, Dr
Next
AddColzLen = AddColzFFDy(D, LenC, Dy)
End Function

Function AddDrs(A As Drs, b As Drs) As Drs
If IsEmpDrs(A) Then AddDrs = b: Exit Function
If IsEmpDrs(b) Then AddDrs = A: Exit Function
If Not IsEqAy(A.Fny, b.Fny) Then Thw CSub, "Dif Fny: Cannot add", "A-Fny B-Fny", A.Fny, b.Fny
AddDrs = Drs(A.Fny, AddAv(A.Dy, b.Dy))
End Function

Function AddDrs3(A As Drs, b As Drs, C As Drs) As Drs
Dim O As Drs: O = AddDrs(A, b)
          AddDrs3 = AddDrs(O, C)
End Function

Function AgrCntzDy(Dy()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    PushI Dr, Si(Pop(Dr))
    PushI AgrCntzDy, Dr
Next
End Function

Function AlignDrs(D As Drs, Gpcc$, CC$) As Drs
'Fm D : ..@Gpcc..@CC.. ! It has a str col @CC to be alignL and @Gpcc to be gp
'Ret  : @D             ! all col @CC are aligned within the gp and rec will gp together.
If NoReczDrs(D) Then AlignDrs = D: Exit Function
Dim Gix&(): Gix = IxyzCC(D, Gpcc)       ' The grouping col ix ay
Dim Aix&(): Aix = IxyzCC(D, CC)         ' The aligning col ix ay
Dim G(): G = GpAsAyDy(D, Gpcc)          ' each gp (an ele of #G) is a dry-of-@D with sam @Gpcc col val
Dim Dy, ADy(), ODy(): For Each Dy In G  'For each data-gp (@Dy)
    ADy = AlignDyzCix(CvAv(Dy), Aix)  ' aligning it (@Dy) into (@ADy)
    PushIAy ODy, ADy              ' pushing the rslt (@ADy-aligned-dry) to oup (@ODy)
Next
AlignDrs = Drs(D.Fny, ODy)
End Function

Sub ApdDrs(O As Drs, M As Drs)
If Not IsEqAy(O.Fny, M.Fny) Then Thw CSub, "Fny are dif", "O.Fny M.Fny", O.Fny, M.Fny
Dim UO&, UM&, U&, J&
UO = UB(O.Dy)
UM = UB(M.Dy)
U = UO + UM + 1
ReDim Preserve O.Dy(U)
For J = UO + 1 To U
    O.Dy(J) = M.Dy(J - UO - 1)
Next
End Sub

Sub ApdDrsSub(O As Drs, M As Drs)
Dim Ixy&(): Ixy = IxyWiNegzSupSubAy(O.Fny, M.Fny)
Dim ODy(): ODy = O.Dy
Dim Dr
For Each Dr In Itr(M.Dy)
    PushI ODy, SelDr(CvAv(Dr), Ixy)
Next
O.Dy = ODy
End Sub

Sub AsgCol(A As Drs, CC$, ParamArray OColAp())
Dim OColAv(), J%, Col, C$()
OColAv = OColAp
C = SyzSS(CC)
For J = 0 To UB(OColAv)
    Col = IntozDrsC(OColAv(J), A, C(J))
    OColAp(J) = Col
Next
End Sub

Sub AsgColDist(A As Drs, CC$, ParamArray OColAp())
Dim OColAv(), J%, Col, b As Drs, C$()
b = DwDist(A, CC)
OColAv = OColAp
C = SyzSS(CC)
For J = 0 To UB(OColAv)
    Col = IntozDrsC(OColAv(J), b, C(J))
    OColAp(J) = Col
Next
End Sub

Function AvDrsC(A As Drs, C) As Variant()
AvDrsC = IntozDrsC(Array(), A, C)
End Function

Sub BrwCnt(Ay, Optional Opt As EmCnt)
Brw FmtDiKqCnt(DiKqCnt(Ay, Opt))
End Sub

Function CntLy(Ay, Optional Opt As EmCnt, Optional SrtOpt As EmCntSrtOpt, Optional IsDesc As Boolean) As String()
Dim D As Dictionary: Set D = DiKqCnt(Ay, Opt)
Dim K
Dim W%: W = DicItmWdt(D)
Dim O$()
Select Case SrtOpt
Case eNoSrt
    CntLy = CntLyzDiKqCnt(D, W)
Case eSrtByCnt
    CntLy = AySrtQ(CntLyzDiKqCnt(D, W), IsDesc)
Case eSrtByItm
    CntLy = CntLyzDiKqCnt(SrtDic(D, IsDesc), W)
Case Else
    Thw CSub, "Invalid SrtOpt", "SrtOpt", SrtOpt
End Select
End Function

Function CntLyzDiKqCnt(DiKqCnt As Dictionary, CntWdt%) As String()
Dim K
For Each K In DiKqCnt.Keys
    PushI CntLyzDiKqCnt, AlignR(DiKqCnt(K), CntWdt) & " " & K
Next
End Function

Sub ColApzDrs(A As Drs, CC$, ParamArray OColAp())
Dim Av(): Av = OColAp
Dim C$(): C = SyzSS(CC)
Dim J%, O
For J = 0 To UB(Av)
    O = OColAp(J)
    O = IntozDrsC(O, A, C(J)) 'Must put into O first!!
                              'This will die: OColAp(J) = IntozDrsC(O, A, C(J))
    OColAp(J) = O
Next
End Sub

Function ColGp(Col(), RLvlGpIx&()) As Variant()
'Fm Col      : Col to gp
'Fm RLvlGpIx : Each V in Col is mapped to GpIx by this RLvlGpix @@
ThwIf_DifSi Col, RLvlGpIx, CSub
Dim MaxGpIx&: MaxGpIx = AyMax(RLvlGpIx)
Dim O(): ReDim O(MaxGpIx)
Dim I&: For I = 0 To MaxGpIx
    O(I) = Array()
Next
I = 0
Dim V: For Each V In Itr(Col)
    Dim GpIx&: GpIx = RLvlGpIx(I)
    PushI O(GpIx), V
    I = I + 1
Next
ColGp = O
End Function

Sub CrtTblzAlignCC(D As Database, T$, Fm$, Gpcc$, CC$)
'Fm T  : @Gpcc @CC
'Fm Fm : ..@Gpcc..@CC..
'Ret   :                ! Crt @T in @D @Fm.  @T will has sam rec as @Fm.  Each gp of rec the @CC will be align and the las chr of each col is [.].
'                       ! because the txt col of a tbl will always RTrim.
Dim GFny$(): GFny = SyzSS(Gpcc)

Dim CFny$(): CFny = SyzSS(CC)

Dim N&:         N = Si(CFny)
Dim SeqN%(): SeqN = IntSeqzN(N, 1)        ' Seq
Dim WFny$(): WFny = SyzQAy("#W?", SeqN) ' Wdt of each CC

Dim WFnySq$(): WFnySq = SyzQteSq(WFny)
Dim MFny$():     MFny = SyzQAy("#M?", SeqN) ' Max of Wdt
Dim FF$:           FF = Gpcc & " " & CC
Dim WMFny$():   WMFny = AddSy(WFny, MFny)

Dim WMFnySq$(): WMFnySq = AddSy(WFnySq, SyzQteSq(MFny))
Dim ColAy$():     ColAy = AmAddSfx(WMFnySq, " Integer")

Dim Ey$(): Ey = SyzQAy("Len(?)", CFny)

Dim Eq0$():  Eq0 = AmAddSfx(WFnySq, " = 0")
Dim Bexp$:  Bexp = JnAnd(Eq0)

Dim TG$: TG = TmpNm("TGp_")

Dim StrQ$:     StrQ = "Max(x.[#W?]) as [#M?]"
Dim Max$():     Max = SyzQAy(StrQ, SeqN)
Dim SelMax$: SelMax = JnCommaSpc(AddSy(GFny, Max))

Dim Gp$: Gp = JnCommaSpc(GFny)

Dim AliEy$(): AliEy = SyzMacro("[{F}] & Space([#M{N}]-[#W{N}]) & ' .'", CFny, SeqN)

Dim QO$:           QO = SqlSelzIntoCpy(T, Fm)
Dim QOAddWM$: QOAddWM = SqlAddColzAy(T, ColAy)            '  ! Gpcc W* M*
Dim QOUpdW$:   QOUpdW = SqlUpdzEy(T, WFny, Ey)            '  ! Update Wcc as Len(CC)
Dim QONoW0$:   QONoW0 = SqlDlt(T, Bexp)                   '  ! Rmv rec with all rmv = 0
Dim QG$:           QG = SqlSelzIntoFmX(TG, SelMax, T, Gp) '  ! G : Gpcc Mcc              Mcc is Max-Wcc gp by Gpcc
Dim QOUpdM$:   QOUpdM = SqlUpdzJn(T, TG, GFny, MFny)
Dim QOAli$:     QOAli = SqlUpdzEy(T, CFny, AliEy)         '  ! align CC
Dim QODrpWM$: QODrpWM = SqlDrpFld(T, WMFny)               '  ! O M : Sam as T              Use T.Gpcc and T3.CC

D.Execute QO        ' Cpy @Fm into @T
D.Execute QOAddWM   ' Add col-W* & M*, where W* is from CC
D.Execute QOUpdW
D.Execute QONoW0
D.Execute QG
D.Execute QOUpdM
D.Execute QOAli
D.Execute QODrpWM
DrpT D, TG
End Sub

Function DicItmWdt%(A As Dictionary)
Dim I, O%
For Each I In A.Items
    O = Max(Len(I), O)
Next
DicItmWdt = O
End Function

Function DiczRenFF(RenFF$) As Dictionary
Set DiczRenFF = New Dictionary
Dim Ay$(): Ay = SyzSS(RenFF)
Dim V: For Each V In SyzSS(RenFF)
    If HasSubStr(V, ":") Then
        DiczRenFF.Add Bef(V, ":"), Aft(V, ":")
    Else
        Thw CSub, "Invalid RenFF.  all Sterm has have [:]", "RenFF", RenFF
    End If
Next
End Function

Sub DmpDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional Nm$, Optional IxCol As EmIxCol, Optional Fmt As EmTblFmt = EiTblFmt, Optional IsSum As Boolean)
DmpAy FmtDrszNoRdu(A, MaxColWdt, BrkColnn$, ShwZer, IxCol, Fmt, Nm)
End Sub

Function DrpColzFny(D As Drs, Fny$()) As Drs
Dim IxAll&(): IxAll = LngSeqzU(UB(D.Fny))
Dim IxToExl&():      IxToExl = Ixy(D.Fny, Fny)
Dim IxSel&(): IxSel = AyMinus(IxAll, IxToExl)
Dim ODy(): ODy = SelDy(D.Dy, IxSel)
DrpColzFny = Drs(SyMinus(D.Fny, Fny), ODy)
End Function

Function Drs(Fny$(), Dy()) As Drs
With Drs
    .Fny = Fny
    .Dy = Dy
End With
End Function

Function DrsInsCV(A As Drs, C$, V) As Drs
DrsInsCV = Drs(CvSy(InsEle(A.Fny, C)), InsColzDy(A.Dy, V, IxzAy(A.Fny, C)))
End Function

Function DrsInsCVAft(A As Drs, C$, V, AftFldNm$) As Drs
DrsInsCVAft = DrsInsCVIsAftFld(A, C, V, True, AftFldNm)
End Function

Function DrsInsCVBef(A As Drs, C$, V, BefFldNm$) As Drs
DrsInsCVBef = DrsInsCVIsAftFld(A, C, V, False, BefFldNm)
End Function

Function DrsInsCVIsAftFld(A As Drs, C$, V, IsAft As Boolean, FldNm$) As Drs
Dim Fny$(), Dy(), Ix&, Fny1$()
Fny = A.Fny
Ix = IxzAy(Fny, C): If Ix = -1 Then Stop
If IsAft Then
    Ix = Ix + 1
End If
Fny1 = InsEle(Fny, FldNm, CLng(Ix))
Dy = InsColzDy(A.Dy, V, Ix)
DrsInsCVIsAftFld = Drs(Fny1, Dy)
End Function

Function Drsz4TRstLy(T4RstLy$(), FF$) As Drs
Dim I, Dy(): For Each I In Itr(T4RstLy)
    PushI Dy, T4Rst(T4RstLy)
Next
Drsz4TRstLy = DrszFF(FF, Dy)
End Function

Function DrszF(FF$) As Drs
DrszF = DrszFF(FF, EmpAv)
End Function

Function DrszFF(FF$, Dy()) As Drs
DrszFF = Drs(TermAy(FF), Dy)
End Function

Function DrszFillLasIfB(D As Drs, C$) As Drs
'Fm D : It has a str col C
'Ret  : Fill in the blank col-C val by las val
Dim LasV$
Dim Fst As Boolean: Fst = True
Dim Ix%: Ix = IxzAy(D.Fny, C)
Dim Dr, Dy(): For Each Dr In Itr(D.Dy)
    Dim V$: V = Dr(Ix)
    If Fst Then
        LasV = V
        Fst = False
    End If
    If V = "" Then
        Dr(Ix) = LasV
    Else
        LasV = V
    End If
    PushI Dy, Dr
Next
DrszFillLasIfB = Drs(D.Fny, Dy)
End Function

Function DrszMapAy(Ay, MapFunNN$, Optional FF$, Optional ValNm$ = "V") As Drs
DrszMapAy = DrszMapItr(Itr(Ay), MapFunNN, FF, ValNm)
End Function

Function DrszMapItr(Itr, MapFunNN$, Optional FF0$, Optional ValNm$ = "V") As Drs
Dim Dy(), V: For Each V In Itr
    Dim Dr(): Dr = Array(V)
    Dim F: For Each F In ItrzSS(MapFunNN)
        PushI Dr, Run(F, V)
    Next
    PushI Dy, Dr
Next
Dim FF$
    If FF0 = "" Then
        FF = ValNm & " " & MapFunNN
    Else
        FF = FF0
    End If
Stop
DrszMapItr = DrszFF(FF, Dy)
End Function

Function DrszRen(D As Drs, RenFF$) As Drs
DrszRen = Drs(FnyzRen(D.Fny, RenFF), D.Dy)
End Function

Function DrszSplitSS(D As Drs, SSCol$) As Drs
'Fm D     : It has a col @SSCol
'Fm SSCol : It is a col nm in @D whose value is SS.
'Ret  : a drs of sam ret but more rec by split@SSCol col to multi record
Dim I%: I = IxzAy(D.Fny, SSCol)
Dim Dr, Dy(): For Each Dr In Itr(D.Dy)
    Dim S: For Each S In Itr(SyzSS(Dr(I)))
        Dr(I) = S
        PushI Dy, Dr
    Next
Next
DrszSplitSS = Drs(D.Fny, Dy)
End Function

Function DrszSSAy(SSAy$(), FF$) As Drs
DrszSSAy = DrszFF(FF, DyoSSAy(SSAy))
End Function

Function DrszTRstLy(TRstLy$(), FF$) As Drs
Dim I, Dy(): For Each I In Itr(TRstLy)
    PushI Dy, SyzTRst(I)
Next
DrszTRstLy = DrszFF(FF, Dy)
End Function

Function DrwIxy(Dr(), Ixy&())
Dim U&: U = MaxEle(Ixy)
Dim O: O = Dr
If UB(O) < U Then
    ReDim Preserve O(U)
End If
DrwIxy = AwIxy(O, Ixy)
End Function

Function DwDist(A As Drs, CC$) As Drs
DwDist = DrszFF(CC, DywDist(SelDrs(A, CC).Dy))
End Function

Function DwInsFF(A As Drs, FF$, NewDy()) As Drs
DwInsFF = Drs(AddSy(SyzSS(FF), A.Fny), NewDy)
End Function

Function DyoSSAy(SSAy$()) As Variant()
Dim SS
For Each SS In Itr(SSAy)
    PushI DyoSSAy, TermAy(SS)
Next
End Function

Function EmpLNewO() As Drs
EmpLNewO = LNewO(EmpAv())
End Function

Function EnsColTyzInt(D As Drs, C) As Drs
If NoReczDrs(D) Then EnsColTyzInt = D: Exit Function
Dim O As Drs, J&, Ix%, Dr
Ix = IxzAy(D.Fny, C)
O = D
If IsSy(O.Dy(0)) Then Stop
For Each Dr In Itr(O.Dy)
    Dr(Ix) = CInt(Dr(Ix))
    O.Dy(J) = Dr
    J = J + 1
Next
EnsColTyzInt = O
End Function

Function FmtLNewO(L_NewL_OldL As Drs, Org_L_Lin As Drs) As String()
'@ : L_NewL_OldL ! Assume all NewL and OldL are nonEmp and <>
'Ret : LinesAy !
Dim SDy(): SDy = SelDy(Org_L_Lin.Dy, LngAy(0, 1))
Dim S As Drs: S = DrszFF("L Lin", SDy)
Dim D As Drs: D = DeCeqC(L_NewL_OldL, "NewL OldL")
Dim NewL As Drs: NewL = LDrszJn(S, D, "L", "NewL")
Dim Gpno As Drs: Gpno = FmtLNewOGpno(NewL)
Dim NLin As Drs: NLin = FmtLNewONLin(Gpno)
Dim Lines As Drs: Lines = FmtLNewOLines(NLin)
Dim OneG As Drs: OneG = FmtNewOneG(NLin)
FmtLNewO = StrCol(OneG, "Lines")
End Function

Function FmtLNewOGpno(NewL As Drs) As Drs
'@ NewL: L Lin NewL ! NewL may empty, when non-Emp, NewL <> Lin
'Ret D: L Lin NewL Gpno ! Gpno is running from 1:
'                      !   all conseq Lin with Emp-NewL is one group
'                      !   each non-Emp-NewL is one gp
Dim IGpno&, Dr, Dy(), Lin, NewL_, LasEmp As Boolean, Emp As Boolean

'For Each Dr In Itr(NewL.Dy)
'    PushI Dr, IsEmpty(Dr(2))
'    PushI Dy, Dr
'Next
'BrwDy Dy
'Erase Dy
'Stop
LasEmp = True
IGpno = 0
For Each Dr In Itr(NewL.Dy)
    Lin = Dr(1)
    NewL_ = Dr(2)
    Emp = IsEmpty(NewL_)
    If Not Emp Then If Lin = NewL_ Then Stop
    If IsEmpty(Lin) Then Stop
    Select Case True
    Case Not Emp: IGpno = IGpno + 1
    Case Emp And Not LasEmp: IGpno = IGpno + 1
    Case Else
    End Select
    PushI Dr, IGpno
    PushI Dy, Dr
    LasEmp = Emp
Next
FmtLNewOGpno = DrszFF("L Lin NewL Gpno", Dy)
End Function

Function FmtLNewOLines(NLin As Drs) As Drs
'Fm NLin: L Gpno NLin SNewL
'Ret Lines: L Gpno Lines
Dim Dr, L&, Gpno&, Lines$, NLin_$, SNewL
Dim Dy()
'Insp SNewL should have some Emp
'    Erase Dy
'    For Each Dr In NLin.Dy
'        PushI Dr, IsEmpty(Dr(2))
'        PushI Dy, Dr
'    Next
'    BrwDrs DrszFF("L Gpno NLin SNewL Emp", Dy)
'    Erase Dy
For Each Dr In Itr(NLin.Dy)
    AsgAp Dr, L, Gpno, NLin_, SNewL
    If IsEmpty(SNewL) Then
        Lines = NLin_
    Else
        Lines = NLin_ & vbCrLf & SNewL
    End If
    PushI Dy, Array(L, Gpno, Lines)
Next
FmtLNewOLines = DrszFF("L Gpno Lines", Dy)
'BrwDrs FmtLNewOLines: Stop
End Function

Function FmtLNewONLin(Gpno As Drs) As Drs
'@Gpno: L Lin NewL Gpno
'Ret E: L Gpno NLin SNewL ! NLin=L# is in front; SNewL = Spc is in front, only when nonEmp
Dim MaxL&: MaxL = AyMax(LngAyzDrs(Gpno, "L"))
Dim NDig%: NDig = Len(CStr(MaxL))
Dim S$: S = Space(NDig + 1)
Dim Dy(), Dr, L&, Lin$, NewL, IGpno&, NLin$, SNewL
For Each Dr In Itr(Gpno.Dy)
    AsgAp Dr, L, Lin, NewL, IGpno
    NLin = AlignR(L, NDig) & " " & Lin
    If IsEmpty(NewL) Then
        SNewL = Empty
    Else
        SNewL = S & NewL
    End If
    PushI Dy, Array(L, IGpno, NLin, SNewL)
Next
FmtLNewONLin = DrszFF("L Gpno NLin SNewL", Dy)
End Function

Function FmtNewOneG(NLin As Drs) As Drs
'@D: L Gpno NLin SNewL !
'Ret E: Gpno Lines ! Gpno now become uniq
Dim O$(), L&, LasG&, Dr, Dy(), Gpno&, NLin_$, SNewL
If NoReczDrs(NLin) Then Exit Function
LasG = NLin.Dy(0)(1)
For Each Dr In Itr(NLin.Dy)
    AsgAp Dr, L, Gpno, NLin_, SNewL
    If LasG <> Gpno Then
        PushI Dy, Array(Gpno, JnCrLf(O))
        Erase O
        LasG = Gpno
    End If
    PushI O, NLin_
    If Not IsEmpty(SNewL) Then PushI O, SNewL
Next
If Si(O) > 0 Then PushI Dy, Array(Gpno, JnCrLf(O))
FmtNewOneG = DrszFF("Gpno Lines", Dy)
End Function

Function FnyzRen(Fny$(), RenFF$) As String()
Dim D As Dictionary: Set D = DiczRenFF(RenFF)
Dim F: For Each F In Fny
    If D.Exists(F) Then
        PushI FnyzRen, D(F)
    Else
        PushI FnyzRen, F
    End If
Next
End Function

Function HasReczDrs(A As Drs) As Boolean
HasReczDrs = Si(A.Dy) > 0
End Function

Function HasReczDy(Dy()) As Boolean
HasReczDy = Si(Dy) > 0
End Function

Function IntozDrsC(Into, A As Drs, C)
Dim O, Ix%, Dy(), Dr
Ix = IxzAy(A.Fny, C): If Ix = -1 Then Stop
O = Into
Erase O
Dy = A.Dy
If Si(Dy) = 0 Then IntozDrsC = O: Exit Function
For Each Dr In Dy
    Push O, Dr(Ix)
Next
IntozDrsC = O
End Function

Function IsEmpDrs(A As Drs) As Boolean
If HasReczDrs(A) Then Exit Function
If Si(A.Fny) > 0 Then Exit Function
IsEmpDrs = True
End Function

Function IsEqDrs(A As Drs, b As Drs) As Boolean
Select Case True
Case Not IsEqAy(A.Fny, b.Fny), Not IsEqAy(A.Dy, b.Dy)
Case Else: IsEqDrs = True
End Select
End Function

Function IsNeFF(A As Drs, FF$) As Boolean
IsNeFF = JnSpc(A.Fny) <> FF
End Function

Function IsSamDrEleCnt(A As Drs) As Boolean
IsSamDrEleCnt = IsSamDrEleCntzDy(A.Dy)
End Function

Function IsSamDrEleCntzDy(Dy()) As Boolean
If Si(Dy) = 0 Then IsSamDrEleCntzDy = True: Exit Function
Dim C%: C = Si(Dy(0))
Dim Dr
For Each Dr In Itr(Dy)
    If Si(Dr) <> C Then Exit Function
Next
IsSamDrEleCntzDy = True
End Function

Function IxdzDrs(A As Drs) As Dictionary
Set IxdzDrs = DiKqIx(A.Fny)
End Function

Function IxyWiNegzSupSubAy(SupAy, SubAy) As Long()
If Not IsAySuper(SupAy, SubAy) Then Thw CSub, "SupAy & SubAy error", "SupAy SubAy", SupAy, SubAy
Dim J%
For J = 0 To UB(SupAy)
    PushI IxyWiNegzSupSubAy, IxzAy(SubAy, SupAy(J))
Next
End Function

Function IxyzJnA(Fny$(), Jn$) As Long()
IxyzJnA = IxyzSubAy(Fny, FnyAzJn(Jn))
End Function

Function IxyzJnB(Fny$(), Jn$) As Long()
IxyzJnB = IxyzSubAy(Fny, FnyBzJn(Jn))
End Function

Function IxzDyDr&(Dy(), Dr)
Dim IDr, O&: For Each IDr In Itr(Dy)
    If IsEqAy(IDr, Dr) Then IxzDyDr = O: Exit Function
    O = O + 1
Next
IxzDyDr = -1
End Function

Function LasDr(A As Drs)
LasDr = LasEle(A.Dy)
End Function

Function LasRec(A As Drs) As Drs
If Si(A.Dy) = 0 Then Thw CSub, "No LasRec", "Drs.Fny", A.Fny
LasRec = Drs(A.Fny, Av((LasEle(A.Dy))))
End Function

Function LNewO(Dy()) As Drs
LNewO = DrszFF("L NewL OldL", Dy)
End Function

Function NColzDrs%(A As Drs)
NColzDrs = Max(Si(A.Fny), NColzDy(A.Dy))
End Function

Function NoReczDrs(A As Drs) As Boolean
NoReczDrs = NoReczDy(A.Dy)
End Function

Function NoReczDy(Dy()) As Boolean
NoReczDy = Si(Dy) = 0
End Function

Function NReczDrs&(A As Drs)
NReczDrs = Si(A.Dy)
End Function

Function NRowzColEv&(A As Drs, ColNm$, EqVal)
NRowzColEv = NRowzInDyoColEv(A.Dy, IxzAy(A.Fny, ColNm), EqVal)
End Function

Sub PushDrs(O As Drss, M As Drs)
With O
    ReDim Preserve .Ay(.N)
    .Ay(.N) = M
    .N = .N + 1
End With
End Sub

Function ReOrdCol(A As Drs, BySubFF$) As Drs
Dim SubFny$(): SubFny = TermAy(BySubFF)
Dim OFny$(): OFny = ReSeqAy(A.Fny, SubFny)
Dim IAy&(): IAy = Ixy(A.Fny, OFny)
Dim ODy(): ODy = SelCol(A.Dy, IAy)
ReOrdCol = Drs(OFny, ODy)
End Function

Function RLvlGpIx(Dy()) As Long()

End Function

Function SelCol(Dy(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    PushI SelCol, AwIxy(Dr, Ixy)
Next
End Function

Function SelDr(Dr(), IxyWiNeg&()) As Variant()
Dim Ix, U%: U = UB(IxyWiNeg)
For Each Ix In IxyWiNeg
    If IsBet(Ix, 0, U) Then
        PushI SelDr, Dr(Ix)
    Else
        PushI SelDr, Empty
    End If
Next
End Function

Function SelDy(Dy(), SelIxy&()) As Variant()
'Ret : SubSet-of-col of @Dy indicated by @SelIxy
Dim Dr: For Each Dr In Itr(Dy)
    PushI SelDy, AwIxy(Dr, SelIxy)
Next
End Function

Function SqzFny(Fny$()) As Variant()
Dim O()
ReDim O(1 To 2, 1 To Si(Fny))
Dim J%: For J = 0 To Si(Fny)
    SetSqr O, Fny
Next
SqzFny = O
End Function

Function SqzDrs(A As Drs) As Variant()
If NoReczDrs(A) Then
    SqzDrs = SqzFny(A.Fny)
    Exit Function
End If
Dim NC&, NR&, Dy(), Fny$()
    Fny = A.Fny
    Dy = A.Dy
    NC = Max(NColzDy(Dy), Si(Fny))
    NR = Si(Dy)
Dim O()
    ReDim O(1 To 1 + NR, 1 To NC)
    SetSqr O, Fny       '<== Set O, R=1
    Dim R&: For R = 1 To NR
        SetSqr O, Dy(R - 1), R + 1 '<== Set O, Fm
    Next
SqzDrs = O
End Function

Function SyzDrsC(A As Drs, ColNm$) As String()
SyzDrsC = IntozDrsC(EmpSy, A, ColNm)
End Function

Function SyzDyC(Dy(), C) As String()
SyzDyC = IntozDyC(EmpSy, Dy, C)
End Function

Sub UpdDbqCol_IfBlnk_ByPrv(D As Database, Q)
With Rs(D, Q)
    Dim L
    If Not .EOF Then L = .Fields(0).Value
    .MoveNext
    While Not .EOF
        If Trim(Nz(.Fields(0).Value, "")) = "" Then
            .Edit
            .Fields(0).Value = L
            .Update
        Else
            L = .Fields(0).Value
        End If
        .MoveNext
    Wend
    .Close
End With
End Sub

Function FstRecFvzDrsC(A As Drs, C)
If Si(A.Dy) = 0 Then Thw CSub, "No Rec", "Drs.Fny", A.Fny
FstRecFvzDrsC = A.Dy(0)(IxzAy(A.Fny, C))
End Function


Sub Z_DiKqCntzDrs()
Dim Drs As Drs, Dic As Dictionary
'Drs = Vbe_Mth12Drs(CVbe)
Set Dic = DiKqCntzDrs(Drs, "Nm")
BrwDic Dic
End Sub

Sub Z_GpCol()
Dim Col():            Col = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
Dim RLvlGpIx&(): RLvlGpIx = LngAy(1, 1, 1, 3, 3, 2, 2, 3, 0, 0)
Dim G():                G = ColGp(Col, RLvlGpIx)
Stop
End Sub

Sub Z_GpDicDKG()
Dim Act As Dictionary, Dy(), Dr1, Dr2, Dr3
Dr1 = Array("A", , 1)
Dr2 = Array("A", , 2)
Dr3 = Array("B", , 3)
Dy = Array(Dr1, Dr2, Dr3)
Set Act = GRxyzCyDic(Dy, IntAy(0), 2)
Ass Act.Count = 2
Ass IsEqAy(Act("A"), Array(1, 2))
Ass IsEqAy(Act("B"), Array(3))
Stop
End Sub

Sub Z_SelDistCnt()
'BrwDrs SelDistCnt(DoPubFun, "Mdn")
End Sub

Function AddColzGpno(D As Drs, NumColn$, GpnoColn$, Optional RunFmNum% = 1) As Drs
'Fm D : ..@NumColn..  ! must has a @NumColn which is a Num.  And assume they are sorted else thw
'Ret  : ..@GpnoColn  ! a drs with @GpnoColn added at end, which is a Gpno running from @RunFmNum
'                      if the conseq dr having @NumColn is in seg, given them a Gpno.
'                      Thw &IncIfJmp if @NumColn is not in ascending order.
Dim Gpno&: Gpno = RunFmNum
Dim Dy()
    If NoReczDrs(D) Then GoTo X
    Dim Ix%: Ix = IxzAy(D.Fny, NumColn)
    Dim CurNum&
    Dim Dr: Dr = D.Dy(0)
    Dim LasNum&: LasNum = Dr(Ix)
    For Each Dr In Itr(D.Dy)
        CurNum = Dr(Ix)
        Gpno = IncIfJmp(Gpno, LasNum, CurNum)
        PushI Dr, Gpno
        PushI Dy, Dr
        LasNum = CurNum
    Next
X:
AddColzGpno = AddColzFFDy(D, GpnoColn, Dy)
End Function

Function IncIfJmp(N&, LasNum, CurNum)
'Ret : Increased @N if LasNum has jumped else no chg @N
'      @N        if LasNum = CurNum or LasNum - 1 = CurNm
'      @N+1      If LasNum - 1 > CurNum
'      Otherwise Thw
Dim Dif&: Dif = CurNum - LasNum
Select Case Dif
Case 0, 1: IncIfJmp = N
Case Is > 1: IncIfJmp = N + 1
Case Else
    Thw CSub, "No in seq.  CurNum should > LasNum", "LasNum CurNum", LasNum, CurNum
End Select
End Function
