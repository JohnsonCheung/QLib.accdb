Attribute VB_Name = "MxEoLof"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxEoLof."
Const MsgoLon_LinMis$ = "No line [Lo Nm xxx]"
Const MsgoLon_NmEr$ = "No line [Lo Nm xxx]"
Const MsgoLofBet_FmToEq$ = "Lno#(?)"
Const MsgoLofBet_SumNotBet$ = "Lno#(?)"

':FunPfx:Eo: :FunPfx #Error-Of# ! It returns String or Sy.  All variables are given to determine the condition of the error and the message of error to be return.
'                               ! It will call MoXXX to find the message.  & MoXXX will use Const Msg_XXX_.. to build the message
':FunPfx:Mo: :FunPfx #Msg-Of#   ! It takes variables to bld the er message return string or Sy.
':CnstPfx:Msg_: :CnstPfx
':CnstPfx:FFo:  :CnstPfx        ! It is public constant
Public Const VdtLofTotss$ = "Sum Avg Cnt"
Public Const VdtLofT1ss$ = _
                            "Lo Ali Bdr Tot Wdt Fmt Lvl Cor Fml Lbl Tit Bet" ' Fmt. i.tm s.pace s.eparated string
Public Const VdtLofT1sszFV$ = "                               Fml Lbl Tit Bet" ' Sng.sigle field per line
Public Const VdtLofT1sszVFF$ = "Lo Ali Bdr Tot Wdt Fmt Lvl Cor                " ' Mul.tiple field per line
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
Function MoLofBet_FmToEq$(L&, FmToFld$): MoLofBet_FmToEq = FmtQQ(MsgoLofBet_FmToEq, LnoStr(L), FmToFld): End Function
Function MoLofBet_SumNotBet(L&, FmFld$, ToFld$, SumFld$): MoLofBet_SumNotBet = FmtQQ(MsgoLofBet_SumNotBet, LnoStr(L), FmFld, ToFld, SumFld): End Function
Function MoLofLon_NmEr(L&, Nm$): MoLofLon_NmEr = FmtQQ(MsgoLon_NmEr, L, Nm): End Function

Function EoLofLon_LinMis$(Dyo_L_Lon())
If Si(Dyo_L_Lon) = 0 Then EoLofLon_LinMis = MsgoLon_LinMis
End Function

Function EoLofLon_NmEr(Dyo_L_Lon()) As String()
Dim Dr: For Each Dr In Itr(Dyo_L_Lon)
    Dim Nm$: Nm = Dr(1)
    If Not IsNm(Nm) Then
        Dim L&: L = Dr(0)
        PushI EoLofLon_NmEr, MoLofLon_NmEr(L, Nm)
    End If
Next
End Function

Function EoLofLon_LinDup(Dyo_L_Lon()) As String()
If Si(Dyo_L_Lon) <= 1 Then Exit Function
'Dim Lnoss: For Each Lnoss In Itr(LnossAy)
'    PushI EoLonDup, FmtQQ(C_Lo_ErNm, Lnoss)
'Next
End Function

Function EoLofLoFld_LinMis(Dyo_L_LoFny()) As String()
If Si(Dyo_L_LoFny) = 0 Then
End If
End Function

Function EoLofLoFld_FldMis$(Dyo_L_LoFny())
If Si(Dyo_L_LoFny) = 1 Then
    Dim Fny$(): Fny = Dyo_L_LoFny(0)(1)
    If Si(Fny) = 0 Then
    End If
End If
End Function

Function EoLofLoFld_FldDupLin(Dyo_L_Fny()) As String()
If Si(Dyo_L_Fny) > 1 Then
    Stop
End If
End Function

Sub Z_EoLof()
Dim Lof$(), Fny$()
GoSub ZZ
Exit Sub
ZZ:
    Brw EoLof(SampLof, SampLofFny)
    Return
T0:
    Fny = SyzSS("A B C D E F G")
    Ept = Sy()
    GoTo Tst
Tst:
    Act = EoLof(Lof, Fny)
    C
    Return
End Sub
Function EoLof_LoFny(L_LoFny As Drs) As String()
Dim Dy(), A$(), B$, C$()
Dy = L_LoFny.Dy
 A = EoLofLoFld_LinMis(Dy)
 B = EoLofLoFld_FldMis(Dy)
 C = EoLofLoFld_FldDupLin(Dy)
EoLof_LoFny = SyzApNB(A, B, C)
End Function

Function EoLof_Lon(L_Lon As Drs) As String()
Dim Dy(), A$(), B$, C$()
Dy = L_Lon.Dy
 A = EoLofLon_NmEr(Dy)
 B = EoLofLon_LinMis(Dy)
 C = EoLofLon_LinDup(Dy)
EoLof_Lon = SyzApNB(A, B, C)
End Function

Function EoLof_Ali(L_Ali_FldLikAy As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Ali_FldLikAy
Dim A$(), B$()
A = EoColx_NotIn(L_Ali_FldLikAy, "Ali", "Ali", VdtLofAlivv)
B = EoColFldLikAy_3Er(Drs, Fny)
EoLof_Ali = SyzApNB(A, B)
End Function

Function EoLof_Fmt(L_Fmt_FldLikAy As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Fmt_FldLikAy
Dim A$(), B$(), C$(), D$(), E$()
EoLof_Fmt = SyzApNB(A)
End Function

Function EoLof_Lvl(L_Lvl_FldLikAy As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Lvl_FldLikAy
Dim A$(), B$(), C$(), D$(), E$()
A = EoColx_NumNotBet(Drs, "Lvl", 2, 8)
B = EoColx_NotNum(Drs, "Lvl")
C = EoColFldLikAy_3Er(Drs, Fny)
EoLof_Lvl = SyzApNB(A, B, C)
End Function

Function EoLof_Cor(L_Cor_FldLikAy As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Cor_FldLikAy
Dim A$(), B$(), C$(), D$(), E$()
A = EoColx_NumNotBet(Drs, "Cor", 2, 8)
B = EoColx_NotNum(Drs, "Cor")
C = EoColFldLikAy_3Er(Drs, Fny)
EoLof_Cor = SyzApNB(A, B, C)
End Function

Function EoLof_Wdt(L_Wdt_FldLikAy As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Wdt_FldLikAy
Dim A$(), B$(), C$()
A = EoColx_NumNotBet(Drs, "Wdt", 5, 100)
B = EoColx_NotNum(Drs, "Wdt")
C = EoColFldLikAy_3Er(Drs, Fny)
EoLof_Wdt = SyzApNB(A, B, C)
End Function

Function EoLof_Lbl(L_F_Lbl As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_F_Lbl
Dim A$(), B$()
Dim FF$: FF = JnSpc(Fny)
A = EoColx_NotIn(Drs, "F", Valn:="Fld", VdtValss:=FF)
B = EoColx_Dup(Drs, "F", "Fld")
EoLof_Lbl = SyzApNB(A, B)
End Function

Function EoLof_Tot(L_Tot_FldLikAy As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Tot_FldLikAy
Dim A$(), B$(), C$()
A = EoColFldLikAy_3Er(L_Tot_FldLikAy, Fny)
B = EoColx_NotIn(L_Tot_FldLikAy, "Tot", "Tot", VdtLofT1ss)
End Function

Function EoLof_Bet(L_Fm_To_Sum As Drs, Fny$()) As String()
Dim A$(), B$(), C$(), FF$, D As Drs
FF = JnSpc(Fny)
D = L_Fm_To_Sum
A = EoColx_NotIn(D, "Fm", "FmFld", FF)
B = EoColx_NotIn(D, "To", "FmFld", FF)
C = EoColx_NotIn(D, "Sum", "SumFld", FF)
Dim IxL%, IxFm%, IxTo%, IxSum%: AsgIx L_Fm_To_Sum, "L Fm To Sum", IxL, IxFm, IxTo, IxSum
Dim L&, PosFm%, PosTo%, PosSum%, FmFld$, ToFld$, SumFld$
Dim Dr: For Each Dr In Itr(L_Fm_To_Sum.Dy)
    L = Dr(IxL)
    FmFld = Dr(IxFm)
    ToFld = Dr(IxTo)
    SumFld = Dr(IxSum)
    PosFm = IxzAy(Fny, FmFld)
    PosTo = IxzAy(Fny, ToFld)
    PosSum = IxzAy(Fny, SumFld)
    If PosFm = PosTo Then PushI EoLof_Bet, MoLofBet_FmToEq(L, FmFld)
    If IsBet(PosSum, PosFm, PosTo) Then PushI EoLof_Bet, MoLofBet_SumNotBet(L, FmFld, ToFld, SumFld)
Next
End Function

Function EoLof_Fml(L_F_Fml As Drs, Fny$()) As String()
EoLof_Fml = EoColF_3Er(L_F_Fml, Fny)
End Function

Function EoLof(Lof$(), Fny$()) As String()
':Lof:  :Fmtr #ListObj-Fmtr# !
':Fmtr: :Ly   #Formatter#
Dim Dta As LofDta: Dta = LofDta(Lof)
With Dta
Dim ELon$():     ELon = EoLof_Lon(.L_Lon)
Dim ELoFld$(): ELoFld = EoLof_Lon(.L_Lon)
Dim F$():       F = .Fny
Dim EAli$(): EAli = EoLof_Ali(.L_Ali_FldLikAy, F)
Dim EWdt$(): EWdt = EoLof_Wdt(.L_Wdt_FldLikAy, F)
Dim EFmt$(): EFmt = EoLof_Fmt(.L_Fmt_FldLikAy, F)
Dim ELvl$(): ELvl = EoLof_Lvl(.L_Lvl_FldLikAy, F)
Dim ECor$(): ECor = EoLof_Cor(.L_Cor_FldLikAy, F)
Dim ETot$(): ETot = EoLof_Tot(.L_Tot_FldLikAy, F)
Dim EBdr$(): EBdr = EoLof_Bdr(.L_Bdr_FldLikAy, F)
Dim EFml$(): EFml = EoLof_Fml(.L_F_Fml, F)
Dim ELbl$(): ELbl = EoLof_Lbl(.L_F_Lbl, F)
Dim ETit$(): ETit = EoLof_Tit(.L_F_Tit, F)
Dim EBet$(): EBet = EoLof_Bet(.L_Fm_To_Sum, F)
End With
Dim O$(): O = SyzApNB(ELon, ELoFld, EAli, EBdr, ETot, EWdt, EFmt, ELvl, ECor, EFml, ELbl, ETit, EBet)
If Si(O) > 0 Then
    EoLof = SyzAp(AySrtQ(O), "-----------", AddIxPfx(Lof, EiBeg1))
End If
End Function

Function EoLof_Tit(L_F As Drs, Fny$()) As String()
EoLof_Tit = EoColF_3Er(L_F, Fny)
End Function

Function EoLof_Bdr(L_Bdr_FldLikAy As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Bdr_FldLikAy
Dim A$(), B$(), C$()
A = EoColFldLikAy_3Er(Drs, Fny)
B = EoColx_NotIn(Drs, "Bdr", "Bdr", VdtLofBdrvv)
C = EoColx_Dup(Drs, "Bdr", "Bdr")
EoLof_Bdr = SyzApNB(A, B, C)
End Function

Function EoLofBet_FldCannotBetFmTo() As String()
'C$ is the col-c of Bet-line.  It should have 2 item and in Fny
'Return Eo of M_Bet_* if any
End Function

Property Get SampLofFny() As String()
SampLofFny = SyzSS("A B C D E F")
End Property

Property Get SampLof() As String()
Dim A As New Bfr
With A
.Var "Ali Center F"
.Var "Ali Left B"
.Var "Ali Right D E"
.Var "Bdr Center F"
.Var "Bdr Left"
.Var "Bdr Right G"
.Var "Cor 12345 B"
.Var "Fml C B * 2"
.Var "Fml F A + B"
.Var "Fmt #,## B C"
.Var "Fmt #,##.## D E"
.Var "Lbl A lksd flks dfj"
.Var "Lbl B lsdkf lksdf klsdj f"
.Var "Lbl A lksd flks dfj"
.Var "Lo Fld B C D E F G"
.Var "Lo Nm BC"
.Var "Lvl 2 C"
.Var "Sum A B X"
.Var "Tit A bc | sdf"
.Var "Tit B bc | sdkf | sdfdf"
.Var "Tot Avg D"
.Var "Tot Cnt C"
.Var "Tot Sum B"
.Var "Wdt 10 B X"
.Var "Wdt 20 D C C"
.Var "Wdt 3000 E F G C"
End With
SampLof = AlignLyz2T(A.Ly)
End Property

