Attribute VB_Name = "MxDb"
Option Compare Text
Option Explicit
Const CNs$ = "sdfsdf"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDb."
Public Const C_Des$ = "Description"

Sub BrwDb(D As Database)
BrwFb D.Name
End Sub

Sub BrwDbzLasTmp()
BrwDb LasTmpDb
End Sub

Sub ClsDb(D As Database)
On Error Resume Next
D.Close
End Sub

Sub CrtQry(D As Database, Qn$, S)
Dim Q As New DAO.QueryDef: Q.Name = Qn: Q.Sql = S
D.QueryDefs.Append Q
End Sub

Sub CrtTbl(D As Database, T, FldDclAy)
D.Execute FmtQQ("Create Table [?] (?)", T, JnComma(FldDclAy))
End Sub

Sub CrtTblTmp(D As Database)
D.TableDefs.Append TmpTd
End Sub

Sub CrtTblzTmp(D As Database)
DrpT D, "#Tmp"
D.TableDefs.Append TmpTd
End Sub

Function DbPth$(D As Database)
DbPth = Pth(D.Name)
End Function

Function DbzReOpn(D As Database) As Database
Set DbzReOpn = Db(D.Name)
End Function

Sub DmpNRec(D As Database)
Dmp FmtNRec(D)
End Sub

Sub DrpDbIfTmp(D As Database)
If IsDbTmp(D) Then
    Dim N$
    N = D.Name
    D.Close
    DltFfn N
End If
End Sub

Sub DrpT(D As Database, T)
If HasTbl(D, T) Then D.Execute "Drop Table [" & T & "]"
End Sub

Sub DrpTmpTbl(D As Database)
DrpTny D, TmpTny(D)
End Sub

Sub DrpTny(D As Database, Tny)
Dim T
For Each T In Itr(Tny)
    DrpT D, T
Next
End Sub

Sub DrpTT(D As Database, TT$)
':TT: :SS
DrpTny D, SyzSS(TT)
End Sub

Sub DrpzAp(D As Database, ParamArray TblAp())
Dim Av(): Av = TblAp
DrpTny D, Av
End Sub

Function DszDb(D As Database, Optional DsNm$) As Ds
Dim Nm$
If DsNm = "" Then
    Nm = D.Name
Else
    Nm = DsNm
End If
DszDb = DszTny(D, Tny(D), Nm)
End Function

Function DszTny(D As Database, Tny$(), Optional DsNm$) As Ds
Dim T
For Each T In Tny
    AddDt DszTny, DtzT(D, CStr(T))
Next
End Function

Sub EnsTmpTbl(D As Database)
If HasTbl(D, "#Tmp") Then Exit Sub
D.Execute "Create Table [#Tmp] (AA Int, BB Text 10)"
End Sub

Function FFzT$(D As Database, T)
FFzT = Termss(Fny(D, T))
End Function

Function FldDesDic(D As Database) As Dictionary
Dim T$, I, J, F$, Des$
Set FldDesDic = New Dictionary
For Each I In Tni(D)
    T = I
    For Each J In Fny(D, T)
        F = J
        Des = FldDes(D, T, F)
        If Des <> "" Then FldDesDic.Add T & "." & F, D
    Next
Next
End Function

Function FmtNRec(D As Database) As String()
Dim T$(): T = Tny(D)
With New Bfr
.X "Fb   " & D.Name
.X "NTbl " & Si(T)
Dim I, J%
For Each I In Itr(T)
    J = J + 1
    .X AlignR(J, 3) & " " & AlignR(NReczT(D, I), 7) & " " & I
Next
FmtNRec = .Ly
End With
End Function

Function HasFF(D As Database, T, FF$) As Boolean
HasFF = FFzT(D, T) = FF
End Function

Function HasQry(D As Database, Q) As Boolean
HasQry = HasReczQ(D, FmtQQ("Select * from MSysObjects where Name='?' and Type=5", Q))
End Function

Function HasReczQ(D As Database, Q) As Boolean
HasReczQ = HasRec(Rs(D, Q))
End Function

Function HasReczT(D As Database, T) As Boolean
HasReczT = HasRec(RszT(D, T))
End Function

Function HasTbl(D As Database, T) As Boolean
HasTbl = HasItn(DbzReOpn(D).TableDefs, T)
End Function

Function HasTblByMSys(D As Database, T) As Boolean
HasTblByMSys = HasRec(Rs(D, FmtQQ("Select Name from MSysObjects where Type in (1,6) and Name='?'", T)))
End Function

Function IsDbOk(D As Database) As Boolean
On Error GoTo X
IsDbOk = D.Name = D.Name
Exit Function
X:
End Function

Function IsDbTmp(D As Database) As Boolean
IsDbTmp = PthzDb(D) = TmpDbPth
End Function

Function JnStrDiczTwoColSql(TwoColSql) As Dictionary _
'Return a dictionary of Ay with Fst-Fld as Key and Snd-Fld as Sy
'Set JnStrDiczTwoColSql = JnStrDicTwoFldRs(RszSql(TwoColSql))
End Function

Function OupTny(D As Database) As String()
OupTny = AwPfx(Tny(D), "@")
End Function

Function PthzDb$(D As Database)
PthzDb = Pth(D.Name)
End Function

Function Qny(D As Database) As String()
Qny = SyzQ(D, "Select Name from MSysObjects where Type=5 and Left(Name,4)<>'MSYS' and Left(Name,4)<>'~sq_'")
End Function

Sub RenTbl(D As Database, T, ToNm$)
D.TableDefs(T).Name = ToNm
End Sub

Sub RenTblzFmPfx(D As Database, FmPfx$, ToPfx$)
Dim T As TableDef
For Each T In D.TableDefs
    If HasPfx(T.Name, FmPfx) Then
        T.Name = RplPfx(T.Name, FmPfx, ToPfx)
    End If
Next
End Sub

Sub RenTTzAddPfx(D As Database, TT$, Pfx$)
Dim T
For Each T In Ny(TT)
    RenTblzAddPfx D, CStr(T), Pfx
Next
End Sub
Sub RunSql(D As Database, Q)
Const CSub$ = CMod & "Rq"
On Error GoTo X
D.Execute Q
Exit Sub
X:
    CrtQry D, TmpNm, Q
    Dim N$: N = D.Name
    Dim E$: E = Err.Description
    Thw CSub, "Running Sql error", "Er Sql Db", E, Q, N
End Sub
Sub Rq(D As Database, Q)
RunSql D, Q
End Sub

Sub Rqq(D As Database, QQ$, ParamArray Ap())
Dim Av(): If UBound(Ap) > 0 Then Av = Ap
RqqAv D, QQ, Av
End Sub

Sub RqqAv(D As Database, QQ$, Av())
'Ret : Run the %Sql by building from &FmtQQ(@QQ,@Av) in @D
Rq D, FmtQQAv(QQ, Av)
End Sub

Function Rs(D As Database, Q) As DAO.Recordset
Const CSub$ = CMod & "Rs"
On Error GoTo X
Set Rs = D.OpenRecordset(Q)
Exit Function
X: Thw CSub, "Error in opening Rs", "Er Sql Db", Err.Description, Q, D.Name
End Function

Function RszQ(D As Database, Q) As DAO.Recordset
Set RszQ = Rs(D, Q)
End Function

Function RszQQ(D As Database, QQ$, ParamArray Ap()) As DAO.Recordset
Dim Av(): If UBound(Ap) > 0 Then Av = Ap
Set RszQQ = Rs(D, FmtQQAv(QQ, Av))
End Function

Function RszQry(D As Database, QryNm$) As DAO.Recordset
Set RszQry = D.QueryDefs(QryNm).OpenRecordset
End Function
Sub SetTblDeszDic(D As Database, Dic As Dictionary)

End Sub
Sub SetFldDeszDic(D As Database, Dic As Dictionary)
Dim T$, F$, Des$, TDotF$, I, J
For Each I In Dic.Keys
    TDotF = I
    Des = Dic(TDotF)
    If HasDot(TDotF) Then
        AsgBrkDot TDotF, T, F
        SetFldDes D, T, F, Des
    Else
        For Each J In Tny(D)
            T = J
            If HasFld(D, T, F) Then
                SetFldDes D, T, F, Des
            End If
        Next
    End If
Next
End Sub

Sub SetTblAttDes(D As Database, Des$)
SetTdDes D, "Att", Des
End Sub

Function SrcTny(D As Database) As String()
Dim T: For Each T In Tni(D)
    PushNB SrcTny, D.TableDefs(T).SourceTableName
Next
End Function


Function TblAttDes$(D As Database)
TblAttDes = TdDes(D, "Att")
End Function


Function TdStrAy(D As Database, TT$) As String()
Dim T$, I
For Each I In ItrzTT(TT)
    T = I
    PushI TdStrAy, TdStrzT(D, T)
Next
End Function

Function TmpFbAy() As String()
TmpFbAy = FfnAy(TmpDbPth, "*.accdb")
End Function

Function TmpTny(D As Database) As String()
TmpTny = AwPfx(Tny(D), "#")
End Function

Function Tni(D As Database)
Asg Itr(Tny(D)), Tni
End Function

Function TnizInp(D As Database)
Asg Itr(TnyzInp(D)), TnizInp
End Function

Function Tntt$(D As Database)
Tntt = Termss(Tny(D))
End Function

Function Tny(D As Database) As String()
Set D = DAO.DBEngine.OpenDatabase(D.Name)
Dim T As TableDef
For Each T In D.TableDefs
    If Not IsTdSys(T) Then
        If Not IsTdHid(T) Then
            PushI Tny, T.Name
        End If
    End If
Next
End Function

Function Tny1(D As Database) As String()
Dim T As TableDef, O$()
Dim X As DAO.TableDefAttributeEnum
X = DAO.TableDefAttributeEnum.dbHiddenObject Or DAO.TableDefAttributeEnum.dbSystemObject
For Each T In D.TableDefs
    Select Case True
    Case T.Attributes And X
    Case Else
        PushI Tny1, T.Name
    End Select
Next
End Function

Function TnyzADO(D As Database) As String()
TnyzADO = TnyzFb(D.Name)
End Function

Function TnyzInp(D As Database) As String()
TnyzInp = AwLik(Tny(D), ">*")
End Function

Function TnyzMSysObj(D As Database) As String()
TnyzMSysObj = SyzQ(D, "Select Name from MSysObjects where Type in (1,6) and Name not Like 'MSys*' and Name not Like 'f_*_Data'")
End Function


Sub Z_BrwT()
Dim D As Database
Stop
DrpTT D, "#A #B"
Rq D, "Select Distinct Sku,BchNo,CLng(Rate) as RateRnd into [#A] from [#T]"
BrwTT D, "#A #T #B"
End Sub

Sub Z_DszDb()
Dim D As Database, Tny0, Act As Ds, Ept As Ds
Stop
ZZ1:
    Set D = Db(SampFbzDutyDta)
    Act = DszDb(D)
    BrwDs Act
    Exit Sub
ZZ2:
    Tny0 = "Permit PermitD"
    'Set Act = Ds( Tny0)
    Stop
End Sub

Sub Z_Qny()
'DmpAy Qny(Db(SampFbzDutyDta))
End Sub
