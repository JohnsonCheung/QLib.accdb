Attribute VB_Name = "MxErzLof"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxErzLof."
Public Const LofT1nn$ = _
                            "Lo Ali Bdr Tot Wdt Fmt Lvl Cor Fml Lbl Tit Bet" ' Fmt. i.tm s.pace s.eparated string
Public Const LofT1nnzSng$ = "                               Fml Lbl Tit Bet" ' Sng.sigle field per line
Public Const LofT1nnzMul$ = "Lo Ali Bdr Tot Wdt Fmt Lvl Cor                " ' Mul.tiple field per line
'
'Lo  Nm  Er    [Lo Nm] has error
'Lo  Nm  Mis   [Lo Nm] line is missed
'Lo  Nm  Dup   [Lo Nm] is Dup
'Lo  Fny Mis   [Lo Fny] is missed
'Lo  Fny Dup   [Lo Fny] is missed
'Ali Val NLis  [Ali Val] is not in @AliVal
'Ali Fld NLis  [Ali Fld] is not in @LoFny
'Bdr Val NLis  [Bdr Val] is not in @BdrVal
'Tot Val NLis  [Tot Val] is not in @TotVal
'Wdt Val NNum  [Wdt Val] is not number
'Wdt Val Mis   [Wdt Val] is missed
'Wdt Val NBet  [Wdt Val] is not between 3 to 100
'Lvl Val NNum  [Lvl Val] is not a number
'Lvl Val NBet  [Lvl Val] is not between 2 and 8
'Lvl Fld NLis  [Lvl Fld] is not in @LoFny
'Lvl Fld Dup
'

Function LofT1Ny() As String()
LofT1Ny = TermAy(LofT1nn)
End Function

Function XELoFldDup(DoLo As Drs) As String()

End Function

Function XLonMis(LoLy$()) As Boolean

End Function
Function XELonMis(DoLo As Drs) As String()
'If IsLoMisNm Then XELonMis = Sy("No Lon")
End Function

Function XELonEr(DoLo As Drs) As String()
'XELonEr = M_Lo_ErNm(LnoAy)
End Function

Function XELonDup(DoLo As Drs) As String()
'Dim Lnoss: For Each Lnoss In Itr(LnossAy)
'    PushI XELonDup, FmtQQ(C_Lo_ErNm, Lnoss)
'Next
End Function

Function XELoFldMis(DoLo As Drs) As String()

End Function
Function XDoLo(LTD As Drs) As Drs

End Function
Sub Z_EoLof()
Dim Lof$(), Fny$()
GoSub T0
T0:
    Lof = SampLof
    Fny = SyzSS("A B C D E F G")
    Ept = Sy()
    GoTo Tst
Tst:
    Act = EoLof(Lof, Fny)
    C
    Return
End Sub
Function EoLof(Lof$(), Fny$()) As String()
':Lof: :Fmtr #ListObj-Fmtr# !
':Fmtr: :Ly #Formatter#
Dim A$(), B$(), C$(), D$(), E$()
Dim LTD As Drs: LTD = TDoLTD(Lof).D
Dim DoLo As Drs: DoLo = XDoLo(LTD)
                A = XELonMis(DoLo)
                B = XELonEr(DoLo)
                C = XELonDup(DoLo)
                D = XELoFldMis(DoLo)
                E = XELoFldDup(DoLo)
Dim ELo$():   ELo = Sy(A, B, C, D, E)

'                A = Sy(FmtQQ(MAli_MustLRCenter))
Dim EAli$(): EAli = Sy(A)

                A = XEVal_NBet("Wdt", 10, 200)
Dim EWdt$(): EWdt = Sy(A)


Dim EFmt$(): EFmt = Sy(A, B, C)

                B = XEVal_NBet("Lvl", 2, 9)
Dim ELvl$(): ELvl = Sy(A, B, C)


Dim ECor$(): ECor = Sy(A, B, C)

Dim EFml$(): EFml = Sy(A, B, C)


Dim ELbl$(): ELbl = Sy(A, B, C)

Dim ETit$(): ETit = Sy(A, B, C)

Dim EBet$(): EBet = Sy(A)

Dim EBdr$(): EBdr = Sy(A, B)

Dim ETot$(): ETot = Sy(A, B, C)

EoLof = Sy(ELo, EAli, EBdr, ETot, EWdt, EFmt, ELvl, ECor, EFml, ELbl, ETit, EBet)
End Function

Function WAny_Tot() As Boolean
Dim Lc As ListColumn
'For Each Lc In A_Lo.ListColumns
    'If LcFmtSpecLy_WAny_Tot(Lc, FmtSpecLy) Then WAny_Tot = True: Exit Function
'Next
End Function
Function EoBdr1(X$) As String()
'Return FldAy from Bdr & X
'Dim FldssAy$(): FldssAy = SSSyzAy(AwRmvT1(Bdr, X))
End Function
Function B_EBdr_Dup() As String()
'EoBdrDup = WMsgzDup(DupT1(Bdr), Bdr)
End Function
Function EoBdrExcessFld() As String()
Dim LFny$(), RfNy$(), CFny$()
LFny = EoBdr1("Left")
RfNy = EoBdr1("Right")
CFny = EoBdr1("Center")
'PushIAy EoBdrExcessFld, FmtQQ(M_Dup, AyMinus(CFny, LFny), "Center", "Left")
'PushIAy EoBdrExcessFld, FmtQQ(M_Dup, AyMinus(CFny, RfNy), "Center", "Right")
'PushIAy EoBdrExcessFld, FmtQQ(M_Dup, AyMinus(LFny, RfNy), "Left", "Right")
End Function
Function EoBdrExcessLin() As String()
Dim L
'For Each L In Itr(SyeT1Sy(Bdr, "Left Right Center"))
'    PushI EoBdrExcessLin, FmtQQ(M_Bdr_ExcessLin, L)
'Next
End Function
Function EoBdrFld() As String()
Dim Fny$(): Fny = Sy(EoBdr1("Left"), EoBdr1("Right"), EoBdr1("Center"))
EoBdrFld = WMsgzFny(Fny, "Bdr")
End Function
Function EoBet() As String()
EoBet = Sy(EoBetDup, EoBetFny, EoBetTermCnt)
End Function
Function EoBetDup() As String()
'EoBetDup = WMsgzDup(DupT1(Bet), Bet)
End Function
Function EoBetFny() As String()
'C$ is the col-c of Bet-line.  It should have 2 item and in Fny
'Return Eo of M_Bet_* if any
End Function
Function EoBetTermCnt() As String()
Dim L$, I
'For Each L In Itr(Bet)
    L = I
    If Si(SyzSS(L)) <> 3 Then
        PushI EoBetTermCnt, WMsgzBetTermCnt(L, 3)
    End If
'Next
End Function
Function EoCor() As String()
Dim L$()
'L = Cor
EoCor = Sy(EoCorDup(L), EoCorFld(L), EoCorVal(L))
'Cor = L
End Function
Function EoCorDup(IO$()) As String()

End Function
Function EoCorFld(IO$()) As String()

End Function
Function EoCorVal(IO$()) As String()
Dim Msg$(), Eo$(), L$, I
For Each I In IO
    L = I
    PushI Msg, EoCorVal1(L)
Next
'IO = AwNoEo(IO, Msg, Eo)
End Function
Function EoCorVal1$(L$)
Dim Cor$
Cor = T1(L)
End Function
Function B_EFld() As String()

End Function
Function EoFldss() As String()

End Function

Function EoFldSngzDup(Fny$(), Lof$()) As String() 'It is for [SngFldLin] only.  That means T2 of LofLin is field name.  Return error msg for any FldNm is dup.
Dim T1$, I
For Each I In SyzSS(LofT1nnzSng) 'It is for [SngFldLin] only
    T1 = I
    PushIAy EoFldSngzDup, EoFldSngzDup__WithinT1(T1)
Next
End Function

Function EoFldSngzDup__WithinT1(T1) As String() 'Within T1 any fld is dup?
Dim DupFld$, I

'For Each I In Itr(DupAmT2zLnxs())
    DupFld = I
'    PushIAy EoFldSngzDup__WithinT1, EoFldSngzDup__DupFld_is_fnd(DupFld, Lnxs, T1)
'Next
End Function

Function XErFml(Fny$()) As String()
XErFml = XErFml__InsideFmlHasInvalidFld(Fny)
End Function

Function XErFml__InsideFmlHasInvalidFld(Fny$()) As String()
Dim J&, Fld$, Fml$, O$(), S$, T1
'Dim Lnxs As Lnxs: Lnxs = WLnxszT1("Fml")
'For J = 0 To Lnxs.N - 1
    'With Lnxs.Ay(J)
'        AsgTTRst .Lin, S, Fld, Fml
        If FstChr(Fml) <> "=" Then
            'PushI O, WMsg_Fml_FstChr(.Lno)
        Else
            Dim EoFny$(): 'EoFny = EoFmlFld(Fld, Fml, Fny)
'            PushIAy O, EoFml__InsideFmlHasInvalidFld1(EoFny, .Lno, Fld, Fml)
        End If
    'End With
'Next
XErFml__InsideFmlHasInvalidFld = O
End Function

Function EoFmlFld(Fld$, Fml$, Fny$()) As String()
'Ret :urn Subset-Fny (quote by []) in [Fml] which is error. _
It is error if any-FmlFny not in [Fny] or =[Fld]
Dim Ny$(): Ny = NyzMacro(Fml, OpnBkt:="[")
If HasEle(Ny, Fld) Then 'PushI EoFmlFld, Fld
'PushIAy EoFmlFld, AyMinus(Fml, Fny)
End If
End Function

Function EoFmt() As String()

End Function
Function EoLbl() As String()

End Function
Function B_EoMisFnyzFmti(Fmti) As String()
End Function

Function B_ETot_Cnt_Must_1_Fld() As String()

End Function

Function B_ETot_Must_Sum_Cnt_Avg() As String()
Dim J%
Dim TotKw$(): TotKw = SyzSS("Avg Cnt Sum")
'For J = 0 To TotT1.N - 1
'    With TotT1.Ay(J)
'    If Not HasEle(TotKw, .Lin) Then PushI B_ETot_Must_Sum_Cnt_Avg, MTot_Must_Sum_Cnt_Avg
'    End With
'Next
End Function
Function EoTot_1() '(Cnt$(), CntLxAy%(), Sum$(), SumLxAy%(), Avg$(), AvgLxAy%()) As Eo
'Dim O As New Eo
'Dim J%, C$, Ix%, Msg$
'For J = 0 To UB(Cnt)
'    C = Cnt(J)
'    Ix = Ixy(Sum, C)
'    If Ix >= 0 Then
'        Msg = FmtQQ(M_Fld_IsCnt_FndInSum, CntLxAy(J), Cnt(J), SumLxAy(Ix))
'        O.PushMsg Msg
'    Else
'        Ix = Ixy(Avg, C)
'        If Ix >= 0 Then
'            Msg = FmtQQ(M_Fld_IsCnt_FndInAvg, CntLxAy(J), Cnt(J), AvgLxAy(Ix))
'            O.PushMsg Msg
'        End If
'    End If
'Next
'Set1Lc EoTot_1 = O
End Function
Function B_ELo_() As String()
'W-Eoror-of-LofLinVal:W means working-value. _
which is using the some Module-Lvl-variables and it is private. _
Val here means the LofValFld of LofLin
'E_BLo = Sy(EoVzNotNum, EoVzNotInLis, EoVzFml, EoVzNBet)
End Function
Function EoVzFml() As String()

End Function
Function XEVal_NBet(T1, FmNumVal, ToNumval) As String()
'Dim Lnx(): Lnx = A_T1ToLnxsDic(T1)
End Function
Function EoVzNotInLis() As String()

End Function
Function EoVzNotNum() As String()
Dim T
For Each T In SyzSS("Wdt Lvl")
Next
End Function
Function EoWdt() As String()
End Function

Function WMsgzBetTermCnt$(L, NTerm%)

End Function

Function WMsgzDupNy(DupNy$(), LnoStrAy$()) As String()
Dim N, J&
For Each N In Itr(DupNy)
'    PushIAy WMsgzDupNy, FmtQQ(M_Dup, N, LnoStrAy(J))
    J = J + 1
Next
End Function
Function WMsgzFny(Fny$(), Lin_Ty$) As String()
'Return Msg if given-Fny has some field not in A.Fny
Dim EoFny$(): EoFny = AyMinus(Fny, Fny)
If Si(EoFny) = 0 Then Exit Function
'PushI WMsgzFny, FmtQQ(M_Fny, EoFny, Lin_Ty)
End Function
Sub Z_EoBet()
Dim Fny$()
'---------------
Fny = SyzSS("A B")
'Eoase Bet
'    PushI Bet, "A B C"
'    PushI Bet, "A B C"
Ept = EmpSy
'    PushIAy Ept, WMsgzDup(Sy("A"), Bet)
GoSub Tst
Exit Sub
'---------------
Tst:
    Act = EoBet
    C
    Return
End Sub
Sub Z_EoFldSngzDup()
Dim Lof$(), Fny$(), Act$(), Ept$()
GoSub T1
Exit Sub
T1:
    Lof = SplitVBar("Fml AA sdlkfsdflk|Fml AA skldf|Fml BB sdklfjdlf|Fml BB sdlfkjsdf|Fml BB sdklfjsdf|Fml CC sdfsdf")
    GoTo Tst
Tst:
    Act = EoFldSngzDup(Fny, Lof)
End Sub

Function Y_Lof() As String()
Y_Lof = SampLof
End Function

Function Y_LoFny() As String()
Y_LoFny = FoSampLo
End Function

Property Get FoSampLo() As String()
FoSampLo = SyzSS("A B C D E F")
End Property
Sub XXXX()

#If False Then
Sum Bet A B C
Lo Nm BC
Lo Fld  B C D E F G
Ali Left  B
Ali Right D E
Ali Center F
Wdt 10  B X
Wdt 20 D C C
Wdt 3000 E F G C
Fmt #,##  B C
Fmt #,##.## D E
Lvl 2  C
Bdr Left
Bdr Right G
Bdr Center F
Tot Sum  B
Tot Cnt C
Tot Avg D
Tit A bc | sdf
Tit B bc | sdkf | sdfdf
Cor 12345  B
Fml F  A + B
Fml C  B * 2
Lbl A lksd flks dfj
Lbl B lsdkf lksdf klsdj f

#End If
End Sub

Property Get SampLof() As String()
Dim A As New Bfr
With A
.Var "Sum Bet  A B C"
.Var "Lo Nm BC"
.Var "Lo Fld  B C D E F G"
.Var "Ali Left  B"
.Var "Ali Right D E"
.Var "Ali Center F"
.Var "Wdt 10  B X"
.Var "Wdt 20 D C C"
.Var "Wdt 3000 E F G C"
.Var "Fmt #,##  B C"
.Var "Fmt #,##.## D E"
.Var "Lvl 2  C"
.Var "Bdr Left "
.Var "Bdr Right G"
.Var "Bdr Center F"
.Var "Tot Sum  B"
.Var "Tot Cnt C"
.Var "Tot Avg D"
.Var "Tit A bc | sdf"
.Var "Tit B bc | sdkf | sdfdf"
.Var "Cor 12345  B"
.Var "Fml F  A + B"
.Var "Fml C  B * 2"
.Var "Lbl A lksd flks dfj"
.Var "Lbl B lsdkf lksdf klsdj f"
End With
SampLof = AlignLyz2T(A.Ly)
End Property

Function Lnoss$(Ixy() As Long)
Lnoss = JnSpc(AmIncEleBy1(Ixy))
End Function

