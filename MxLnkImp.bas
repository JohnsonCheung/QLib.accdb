Attribute VB_Name = "MxLnkImp"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxLnkImp."
Public Const StdEleLines$ = _
"E Crt Dte;Req;Dft=Now" & vbCrLf & _
"E Tim Dte" & vbCrLf & _
"E Lng Lng" & vbCrLf & _
"E Mem Mem" & vbCrLf & _
"E Dte Dte" & vbCrLf & _
"E Nm  Txt;Req;Sz=50"
Public Const StdETFLines$ = _
"ETF Nm  * *Nm          " & vbCrLf & _
"ETF Tim * *Tim         " & vbCrLf & _
"ETF Dte * *Dte         " & vbCrLf & _
"ETF Crt * CrtTim       " & vbCrLf & _
"ETF Lng * Si           " & vbCrLf & _
"ETF Mem * Lines *Ft *Fx"

Property Get LnkSpecTp$()
Const A_1$ = "E Mem | Mem Req AlwZLen" & _
vbCrLf & "E Txt | Txt Req" & _
vbCrLf & "E Crt | Dte Req Dft=Now" & _
vbCrLf & "E Dte | Dte" & _
vbCrLf & "F Amt * | *Amt" & _
vbCrLf & "F Crt * | CrtDte" & _
vbCrLf & "F Dte * | *Dte" & _
vbCrLf & "F Txt * | Fun * Txt" & _
vbCrLf & "F Mem * | Lines" & _
vbCrLf & "T Sess | * CrtDte" & _
vbCrLf & "T Msg  | * Fun *Txt | CrtDte" & _
vbCrLf & "T Lg   | * Sess Msg CrtDte" & _
vbCrLf & "T LgV  | * Lg Lines" & _
vbCrLf & "D . Fun | Function name that call the log" & _
vbCrLf & "D . Fun | Function name that call the log" & _
vbCrLf & "D . Msg | it will a new record when Lg-function is first time using the Fun+MsgTxt" & _
vbCrLf & "D . Msg | ..."

LnkSpecTp = A_1
End Property

Sub Z_LnkImp()
Dim LnkImpSrc$(), Db As Database
GoSub T0
Exit Sub
T0:
    LnkImpSrc = Y_LnkImpSrc
    Set Db = TmpDb
    GoTo Tst
Tst:
    LnkImp LnkImpSrc, Db
    Return
End Sub

Sub LnkImp(LnkImpSrc$(), Db As Database)
'ThwIf_Er EoLnk(InpFilSrc, LnkImpSrc), CSub
Dim Ip          As TDoLTDH:                   Ip = TDoLTDH(LnkImpSrc)
Dim FbTblLy$():                        FbTblLy = IndentedLy(LnkImpSrc, "FbTbl")
Dim Dic_Fbt_Fbn As Dictionary: Set Dic_Fbt_Fbn = DiczVkkLy(FbTblLy)
Dim FbTny$():                            FbTny = SyzDicKey(Dic_Fbt_Fbn)

Dim FxTblLy$(): FxTblLy = IndentedLy(LnkImpSrc, "FxTbl")
Dim DFx As Drs:     DFx = WDFx(FxTblLy)                  ' T Fxn Wsn Stru
Dim FxTny$():     FxTny = StrCol(DFx, "T")

Dim Dic_Fn_Ffn As Dictionary: Set Dic_Fn_Ffn = Dic(IndentedLy(LnkImpSrc, "Inp"))

'== Lnk=================================================================================================================
Dim D1   As Drs:   D1 = WLnkFx(DFx, Dic_Fn_Ffn)         ' T S Cn
Dim D2   As Drs:   D2 = WLnkFb(Dic_Fbt_Fbn, Dic_Fn_Ffn)
Dim D    As Drs:    D = AddDrs(D1, D2)
Dim OLnk:  LnkTblzDrs Db, D               ' <======
            
'== Imp=================================================================================================================
Dim Wh         As Dictionary:         Set Wh = Dic(IndentedLy(LnkImpSrc, "Tbl.Where"))
Dim Dic_T_Stru As Dictionary: Set Dic_T_Stru = WDic_T_Stru(FbTny, DFx)

Dim DStru As Drs:  DStru = WDStru(Ip)                     ' Stru F Ty E
Dim ImpSqy$():    ImpSqy = WImpSqy(Dic_T_Stru, DStru, Wh)
Dim OImp:                  RunSqy Db, ImpSqy
Dim ODmpNRec:              DmpNRec Db
End Sub

Function WDStru(Ip As TDoLTDH) As Drs
'Fm Ip : L T1 Dta IsHdr}
'Ret WDStru: Stru F Ty E
Dim A As Drs, Dr, Dy(), B As Drs, T1$, Dta$
A = DwEqSel(Ip.D, "IsHdr", False, "T1 Dta")
B = DwPfx(A, "T1", "Stru.") 'T1 Dta
For Each Dr In Itr(B.Dy)
    T1 = Dr(0)
    Dta = Dr(1)
    PushI Dy, XDrOfStru(T1, Dta)
Next
WDStru = DrszFF("Stru F Ty E", Dy)
End Function

Function XDrOfStru(T1$, Dta$) As Variant()
Dim F$, Ty$, E$, Stru$
Stru = RmvPfx(T1, "Stru.")
F = ShfT1(Dta)
Ty = ShfT1(Dta)
E = RmvSqBkt(Dta)
XDrOfStru = Array(Stru, F, Ty, E)
End Function

Function WDic_T_Stru(FbTny$(), DFx As Drs) As Dictionary
Dim Dr, IxT%, IxStru%, T
Set WDic_T_Stru = New Dictionary
For Each T In Itr(FbTny)
    WDic_T_Stru.Add T, T
Next
AsgIx DFx, "T Stru", IxT, IxStru
For Each Dr In Itr(DFx.Dy)
    WDic_T_Stru.Add Dr(IxT), Dr(IxStru)
Next
End Function

Function WImpSqy(Dic_T_Stru As Dictionary, DStru As Drs, Dic_T_Bexp As Dictionary) As String()
Dim I, Fny$(), Ix As Dictionary, Ey$(), T$, Into$, LnkColLy$(), Bexp$, A As Drs, Stru$
For Each I In Dic_T_Stru.Keys
    Stru = Dic_T_Stru(I)
       T = ">" & I
    Into = "#I" & I
       A = DwEqSel(DStru, "Stru", Stru, "F Ty E")
     Fny = StrCol(A, "F")
      Ey = RmvSqBktzSy(StrCol(A, "E"))
   Bexp = VzDicIf(Dic_T_Bexp, I)
    PushI WImpSqy, SqlSel_Fny_Extny_Into_T_OB(Fny, Ey, Into, T, Bexp)
Next
End Function

Function WDFx(FxTblLy$()) As Drs
'Ret DFx : T Fxn Ws Stru
Dim Lin, L$, A$, T$, Fxn$, Ws$, Stru$, Dy()
For Each Lin In Itr(FxTblLy)
    L = Lin
    T = ShfT1(L)
    A = ShfT1(L)
    Fxn = BefDotOrAll(A)
    Ws = AftDot(A)
    If Fxn = "" Then Fxn = T
    If Ws = "" Then Ws = "Sheet1"
    Stru = StrDft(L, T)
    PushI Dy, Array(T, Fxn, Ws, Stru)
Next
WDFx = DrszFF("T Fxn Ws Stru", Dy)
End Function

Function WLnkFb(Dic_Fbt_Fbn As Dictionary, Dic_Fbn_Fb As Dictionary) As Drs
'Ret: *LnkFb::Drs{T S Cn)
Dim Fbn$, A$, S$, Fbt, T$, Cn$, Fb$, Dy()
For Each Fbt In Dic_Fbt_Fbn.Keys
    Fbn = Dic_Fbt_Fbn(Fbt)
    If Not Dic_Fbn_Fb.Exists(Fbn) Then
        Thw CSub, "Dic_Fbn_Fb does not contains Fbn", "Fbn Dic_Fbn_Fb", Fbn, Dic_Fbn_Fb
    End If
    Fb = Dic_Fbn_Fb(Fbn)
    Cn = DaoCnStrzFb(Fb)
    T = ">" & Fbt
    S = Fbt
    PushI Dy, Array(T, S, Cn)
Next
WLnkFb = DrszFF("T S Cn", Dy)
End Function

Function WLnkFx(DFx As Drs, Dic_Fxn_Fx As Dictionary) As Drs
'Fm : @DFx :: Drs{T Fxn Ws Stru}
'Ret: *LnkFx::Drs{T S Cn}
Dim Dy(), Dr, S$, Fx$, Ws$, Cn$, T$, Fxn$, IxT%, IxWs%, IxFxn%
AsgIx DFx, "T Ws Fxn", IxT, IxWs, IxFxn
For Each Dr In Itr(DFx.Dy)
    T = Dr(IxT)
    Ws = Dr(IxWs)
    Fxn = Dr(IxFxn)
    If Not Dic_Fxn_Fx.Exists(Fxn) Then Thw CSub, "Dic_Fxn_Fx does not have Key", "Fxn-Key Dic_Fxn_Fx", T, Dic_Fxn_Fx
    Fx = Dic_Fxn_Fx(Fxn)
    If IsNeedQte(Ws) Then
        S = "'" & Ws & "$'"
    Else
        S = Ws & "$"
    End If
    Cn = DaoCnStrzFx(Fx)
    T = ">" & T
    PushI Dy, Array(T, S, Cn)
Next
WLnkFx = DrszFF("T S Cn", Dy)
End Function
