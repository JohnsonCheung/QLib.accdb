Attribute VB_Name = "MxDaoLg"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoLg."
Private XSchm$()
Private X_W As Database
Private X_L As Database
Private X_Sess&
Private X_Msg&
Private X_Lg&
Private O$() ' Used by EntAyR

Sub CurLgLis(Optional Sep$ = " ", Optional Top% = 50)
D CurLgLy(Sep, Top)
End Sub

Function CurLgLy(Optional Sep$ = " ", Optional Top% = 50) As String()
CurLgLy = RsMapJn(CurLgRs(Top), Sep)
End Function

Function CurLgRs(Optional Top% = 50) As DAO.Recordset
Set CurLgRs = L.OpenRecordset(FmtQQ("Select Top ? x.*,Fun,MsgTxt from Lg x left join Msg a on x.Msg=a.Msg order by Sess desc,Lg", Top))
End Function

Sub CurSessLis(Optional Sep$ = " ", Optional Top% = 50)
D CurSessLy(Sep, Top)
End Sub

Function CurSessLy(Optional Sep$, Optional Top% = 50) As String()
CurSessLy = RsMapJn(CurSessRs(Top), Sep)
End Function

Function CurSessRs(Optional Top% = 50) As DAO.Recordset
Set CurSessRs = L.OpenRecordset(FmtQQ("Select Top ? * from sess order by Sess desc", Top))
End Function

Function CvSess&(A&)
If A > 0 Then CvSess = A: Exit Function
'CvSess = FvzQ(L, "select Max(Sess) from Sess")
End Function

Sub DmpFei(A As Fei)
'Debug.Print A.ToStr
End Sub

Sub EnsMsg(Fun$, MsgTxt$)
With L.TableDefs("Msg").OpenRecordset
    .Index = "Msg"
    .Seek "=", Fun, MsgTxt
    If .NoMatch Then
        .AddNew
        !Fun = Fun
        !MsgTxt = MsgTxt
        X_Msg = !Msg
        .Update
    Else
        X_Msg = !Msg
    End If
End With
End Sub

Sub EnsSess()
If X_Sess > 0 Then Exit Sub
With L.TableDefs("Sess").OpenRecordset
    .AddNew
    X_Sess = !Sess
    .Update
    .Close
End With
End Sub

Property Get L() As Database
Const CSub$ = CMod & "L"
On Error GoTo X
If IsNothing(X_L) Then
    LgOpn
End If
Set L = X_L
Exit Property
X:
Dim E$, ErNo%
ErNo = Err.Number
E = Err.Description
If ErNo = 3024 Then
    'LgSchmImp
    LgCrtv1
    LgOpn
    Set L = X_L
    Exit Property
End If
'Inf CSub, "Cannot open LgDb", "Er ErNo", E, ErNo
Stop
End Property

Sub Lg(Fun$, MsgTxt$, ParamArray Ap())
EnsSess
EnsMsg Fun, MsgTxt
WrtLg Fun, MsgTxt
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
If Si(Av) = 0 Then Exit Sub
Dim J%, V
With L.TableDefs("LgV").OpenRecordset
    For Each V In Av
        .AddNew
        !Lines = LineszV(V)
        .Update
    Next
    .Close
End With
End Sub

Sub LgAsg(A&, OSess&, OTimStr_Dte$, OFun$, OMsgTxt$)
Dim Q$
Q = FmtQQ("select Fun,MsgTxt,Sess,x.CrtTim from Lg x inner join Msg a on x.Msg=a.Msg where Lg=?", A)
Dim D As Date
AsgRs L.OpenRecordset(Q), OFun, OMsgTxt, OSess, D
'OTimStr_Dte = TimStr(D)
End Sub

Sub LgBeg()
Lg ".", "Beg"
End Sub

Sub LgBrw()
BrwFt LgFt
End Sub

Sub LgCls()
On Error GoTo Er
X_L.Close
Er:
Set X_L = Nothing
End Sub

Sub LgCrt()
CrtFb LgFb
Dim D As Database, T As DAO.TableDef
Set D = Db(LgFb)
'
Set T = New DAO.TableDef
T.Name = "Sess"
AddFdzId T
AddFdzTimstmp T, "Dte"
D.TableDefs.Append T
'
Set T = New DAO.TableDef
T.Name = "Msg"
AddFdzId T
AddFdzTxt T, "Fun MsgTxt"
AddFdzTimstmp T, "Dte"
D.TableDefs.Append T
'
Set T = New DAO.TableDef
T.Name = "Lg"
AddFdzId T
AddFdzLng T, "Sess Msg"
AddFdzTimstmp T, "Dte"
D.TableDefs.Append T
'
Set T = New DAO.TableDef
T.Name = "LgV"
AddFdzId T
AddFdzLng T, "Lg Val"
D.TableDefs.Append T

'CrtPkDTT Db, "Sess Msg Lg LgV"
'CrtSkD Db, "Msg", "Fun MsgTxt"
End Sub

Sub LgCrtv1()
Dim Fb$
Fb = LgFb
If HasFfn(Fb) Then Exit Sub
'DbCrtSchm CrtFb(Fb), LgSchmLines
End Sub

Property Get LgDb() As Database
Set LgDb = L
End Property

Sub LgEnd()
Lg ".", "End"
End Sub

Property Get LgFb$()
LgFb = LgPth & LgFn
End Property

Property Get LgFn$()
LgFn = "Lg.accdb"
End Property

Property Get LgFt$()
Stop '
End Property

Sub LgKill()
LgCls
If HasFfn(LgFb) Then Kill LgFb: Exit Sub
Debug.Print "LgFb-[" & LgFb & "] not Has"
End Sub

Function LgLinesAy(A&) As Variant()
Dim Q$
Q = FmtQQ("Select Lines from LgV where Lg = ? order by LgV", A)
'LgLinesAy = RsAy(L.OpenRecordset(Q))
End Function

Sub LgLis(Optional Sep$ = " ", Optional Top% = 50)
CurLgLis Sep, Top
End Sub

Function LgLy(A&) As String()
Dim Fun$, MsgTxt$, TimStr$, Sess&, Sfx$
LgAsg A, Sess, TimStr, Fun, MsgTxt
Sfx = FmtQQ(" @? Sess(?) Lg(?)", TimStr, Sess, A)
'LgLy = LyzFunMsg(Fun & Sfx, MsgTxt, LgLinesAy(A))
Stop '
End Function

Sub LgOpn()
Set X_L = Db(LgFb)
End Sub

Property Get LgPth$()
Static Y$
'If Y = "" Then Y = PgmPth & "Log\": EnsPth Y
LgPth = Y
End Property

Property Get LgSchm() As String()
If Si(XSchm) = 0 Then
X "E Mem | Mem Req AlwZLen"
X "E Txt | Txt Req"
X "E Crt | Dte Req Dft=Now"
X "E Dte | Dte"
X "E Amt | Cur"
X "F Amt * | *Amt"
X "F Crt * | CrtDte"
X "F Dte * | *Dte"
X "F Txt * | Fun * Txt"
X "F Mem * | Lines"
X "T Sess | * CrtDte"
X "T Msg  | * Fun *Txt | CrtDte"
X "T Lg   | * Sess Msg CrtDte"
X "T LgV  | * Lg Lines"
X "D . Fun | Function name that call the log"
X "D . Fun | Function name that call the log"
X "D . Msg | it will a new record when Lg-function is first time using the Fun+MsgTxt"
X "D . Msg | ..."
End If
LgSchm = XSchm
End Property

Sub SessBrw(Optional A&)
BrwAy SessLy(CvSess(A))
End Sub

Function SessLgAy(A&) As Long()
Dim Q$
Q = FmtQQ("select Lg from Lg where Sess=? order by Lg", A)
'SessLgAy = LngAyDbq(L, Q)
End Function

Sub SessLis(Optional Sep$ = " ", Optional Top% = 50)
CurSessLis Sep, Top
End Sub

Function SessLy(Optional A&) As String()
Dim LgAy&()
LgAy = SessLgAy(A)
'SessLy = AyOfAyAy(MapAy(LgAy, "LgLy"))
End Function

Function SessNLg%(A&)
'SessNLg = FvzQ(L, "Select Count(*) from Lg where Sess=" & A)
End Function

Sub WrtLg(Fun$, MsgTxt$)
With L.TableDefs("Lg").OpenRecordset
    .AddNew
    !Sess = X_Sess
    !Msg = X_Msg
    X_Lg = !Lg
    .Update
End With
End Sub

Private Sub X(A$)
PushI XSchm, A
End Sub


Sub Z_Lg()
LgKill
Debug.Assert Dir(LgFb) = ""
LgBeg
Debug.Assert Dir(LgFb) = LgFn
End Sub
