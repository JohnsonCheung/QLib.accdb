Attribute VB_Name = "MxSampSchm"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxSampSchm."
Public Const SampSchmVbl1$ = _
"Ele" & _
"| Crt Dte;Req;Dft=Now|" & _
"| Tim Dte" & _
"| Lng Lng" & _
"| Mem Mem" & _
"| Dte Dte" & _
"| Nm  Txt;Req;Sz=50" & _
"|FldEle" & _
"| Nm  * *Nm          " & _
"| Tim * *Tim         " & _
"| Dte * *Dte         " & _
"| Crt * CrtTim       " & _
"| Lng * Si           " & _
"| Mem * Lines *Ft *Fx"

Public Const SampSchmVbl3$ = _
"|Tbl" & _
"| LoFmt   *Id Lon" & _
"| LoFmtWdt LoFmtId Wdt | Fldss" & _
"| LoFmtLvl LoFmtId Lvl | Fldss" & _
"| LoFmtBet LoFmtId Fld | FmFld ToFld" & _
"| LoFmtTot LoFmtId TotCalc | Fldss" & _
"|Fld Mem Fldss" & _
"| Nm  Fld FmFld ToFld" & _
"| Lng TotCalc" & _
"|Ele" & _
"| Lvl B Req [VdtRul = >=2 and <=8] Dft=2"
Sub EdtSampSchm4(): EdtSamp 4: End Sub
Sub EdtSampSchm1(): EdtSamp 1: End Sub
Sub EdtSampSchm2(): EdtSamp 2: End Sub
Sub EdtSampSchm3(): EdtSamp 3: End Sub
Property Get SampSchm1() As String(): SampSchm1 = Samp(1): End Property
Property Get SampSchm2() As String(): SampSchm2 = Samp(2): End Property
Property Get SampSchm3() As String(): SampSchm3 = Samp(3): End Property
Property Get SampSchm4() As String(): SampSchm4 = Samp(4): End Property

Sub Z_EtLinEr_zLnx()
Dim T(): ReDim T(1): T(0) = 999
GoSub Cas0
Stop
GoSub Cas1
GoSub Cas2
GoSub Cas3
GoSub Cas4
GoSub Cas5
GoSub Cas6
Exit Sub
Dim EptEr$(), ActEr$()
Cas0:
    T(1) = "Tbl 1"
    Ept = Sy("--- #1000[Tbl 1] FldNm[1] is not a name")
    GoTo Tst
Cas1:
    T(1) = "A"
    Push EptEr, "should have a |"
    Ept = Sy("")
    GoTo Tst
Cas2:
    T(1) = "A | B B"
    Ept = Sy("")
    Push EptEr, "dup fields[B]"
    GoTo Tst
Cas3:
    T(1) = "A | B B D C C"
    Ept = Sy("")
    Push EptEr, "dup fields[B C]"
    GoTo Tst
Cas4:
    T(1) = "A | * B D C"
    Ept = Sy("")
    With Ept
        .T = "A"
        .Fny = SyzSS("A B D C")
    End With
    GoTo Tst
Cas5:
    T(1) = "A | * B | D C"
    Ept = Sy("")
    With Ept
        .T = "A"
        .Fny = SyzSS("A B D C")
        .Sk = SyzSS("B")
    End With
    GoTo Tst
Cas6:
    T(1) = "A |"
    Ept = Sy("")
    Push EptEr, "should have fields after |"
    GoTo Tst
Tst:
'    Act = EoT_LinEoDr(T)
    Return
End Sub

Private Function SampFn$(N%)
SampFn = "SampSchm" & N & ".txt"
End Function
Private Sub EdtSamp(N%)
EdtRes SampFn(N)
End Sub
Private Function Samp(N%) As String()
Samp = Res(SampFn(N))
End Function
Private Function Sampl$(N%)
Sampl = Resl(SampFn(N))
End Function

Function SampSchmlzN$(N_1_To_4%)
If IsBet(N_1_To_4, 1, 4) Then
    SampSchmlzN = Sampl(N_1_To_4)
End If
End Function

