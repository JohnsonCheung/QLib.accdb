Attribute VB_Name = "MxTblInf"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxTblInf."
Const Skn$ = "SecondaryKey"

Sub AddFd(D As Database, T, Fd As DAO.Fields)
D.TableDefs(T).Fields.Append Fd
End Sub

Sub AddFdzTF(D As Database, T, F$, Ty As DataTypeEnum, Optional Si%, Optional Precious%)
If HasFld(D, T, F) Then Exit Sub
Dim S$, SqlTy$
SqlTy = SqlTyzDao(Ty, Si, Precious)
S = FmtQQ("Alter Table [?] Add Column [?] ?", T, F, Ty)
D.Execute S
End Sub

Sub AddFdzExpr(D As Database, T, F$, Expr$, Ty As DAO.DataTypeEnum)
D.TableDefs(T).Fields.Append Fd(F, Ty, Expr:=Expr)
End Sub

Sub AsgColApzDrsFF(D As Drs, FF$, ParamArray OColAp())
Dim F, J%
For Each F In TermAy(FF)
    OColAp(J) = ColzDrs(D, CStr(F))
    J = J + 1
Next
End Sub

Sub BrwTblzByDt(D As Database, T)
BrwDt DtzT(D, T)
End Sub

Function CnStrzLnkTbl$(D As Database, T)
CnStrzLnkTbl = D.TableDefs(T).Connect
End Function

Function CnStrzT$(D As Database, T)
On Error Resume Next
CnStrzT = D.TableDefs(T).Connect
End Function

Function ColSetzT(D As Database, T, F$) As Aset
Set ColSetzT = ColSetzRs(ColzRs(D, T, F))
End Function

Function ColzRs(D As Database, T, F$) As DAO.Recordset
Set ColzRs = Rs(D, SqlSel_F_T(F, T))
End Function

Function CommaSpcSqAv$(Av())

End Function

Sub CrtPk(D As Database, T)
D.Execute PkSql(T)
End Sub

Sub CrtSk(D As Database, T, Skff$)
D.Execute SkSqlzFF(T, Skff)
End Sub

Sub CrtEmpTblzDrs(D As Database, T, Drs As Drs)
CrtTblzShtTyscfBql D, T, ShtTyscfBqlzDrs(Drs)
End Sub

Sub CrtTblzCloneEmp(D As Database, T, FmTbl$)
Rq D, SqlSel_Into_T_WhFalse(T, FmTbl)
End Sub

Sub CrtTblzDrs(D As Database, T, Drs As Drs)
CrtTblzDrs D, T, Drs
InsTblzDy D, T, Drs.Dy
End Sub

Sub CrtTblzDrszAllStr(D As Database, T, Drs As Drs)
CrtAllStrEmpTblzDrs D, T, Drs
InsTblzDy D, T, Drs.Dy
End Sub

Sub CrtAllStrEmpTblzDrs(D As Database, T, Drs As Drs)
Dim C&, F, O$(), Dy()
Dy = Drs.Dy
For Each F In Drs.Fny
    If IsMemCol(ColzDy(Dy, C)) Then
        PushI O, "M:" & F
    Else
        PushI O, F
    End If
    C = C + 1
Next
CrtTblzShtTyscfBql D, T, Jn(O, "`")
End Sub

Sub CrtTblzDup(D As Database, T, FmTbl, KK$)
Dim Ky$(), K$, Jn$, Tmp$, J%
Ky = SyzSS(KK)
Tmp = "##" & TmpNm
K = JnComma(Ky)
For J = 0 To UB(Ky)
    Ky(J) = FmtQQ("x.?=a.?", Ky(J), Ky(J))
Next
Jn = Join(Ky, " and ")
Dim Into$
D.Execute FmtQQ("Select Distinct ?,Count(*) as Cnt into [?] from [?] group by ? having Count(*)>1", K, Tmp, FmTbl, K)
D.Execute FmtQQ("Select x.* into [?] from [?] x inner join [?] a on ?", Into, FmTbl, Tmp, Jn)
DrpT D, Tmp
End Sub

Sub CrtTblzJnFld(D As Database, T, KK$, JnFld$, Optional Sep$ = " ")
Dim Tar$, LisFld$
    Tar = T & "_Jn_" & JnFld
    LisFld = JnFld & "_Jn"
Rq D, SqlSel_Fny_Into_T_OB(Ny(KK), Tar, T)
AddFdzTF D, T, LisFld, dbMemo
Dim KKIdx&(), JnFldIx&
    KKIdx = Ixy(Fny(D, T), Ny(KK))
    JnFldIx = IxzTF(D, T, JnFld)
InsTblzDy D, T, DyJnFldKK(DyoT(D, T), KKIdx, JnFldIx)
End Sub

Function CsvLyzDbt(D As Database, T) As String()
CsvLyzDbt = CsvLyzRs(RszT(D, T))
End Function

Function DaoTyzNumCol$(NumCol)
ThwIf_NotAy NumCol, CSub
Dim O As VbVarType: O = VarType(NumCol(0))
If Not IsNumzVbTy(O) Then Stop
Dim V
For Each V In NumCol
    O = MaxNumVbTy(O, VarType(V))
Next
DaoTyzNumCol = DaoTyzVbTy(O)
End Function

Function DaoTyzTF(D As Database, T, F) As DAO.DataTypeEnum
DaoTyzTF = D.TableDefs(T).Fields(F).Type
End Function

Function DiKqCntzTF(D As Database, T, F$) As Dictionary
Set DiKqCntzTF = DiKqCntzRs(ColzRs(D, T, F$))
End Function

Sub DrpFld(D As Database, T, FF$)
Dim F$, I
For Each I In ItrzTT(FF)
    F = I
    D.Execute SqlDrpCol_T_F(T, F)
Next
End Sub

Function DrszT(D As Database, T) As Drs
DrszT = DrszRs(RszT(D, T))
End Function

Function DtzT(D As Database, T) As Dt
DtzT = Dt(T, Fny(D, T), DyoT(D, T))
End Function

Function DyoT(D As Database, T) As Variant()
DyoT = DyoRs(RszT(D, T))
End Function

Function DyoTFF(D As Database, T, FF$) As Variant()
DyoTFF = DyoQ(D, SqlSel_FF_T(FF, T))
End Function

Function Fds(D As Database, T) As DAO.Fields
Set Fds = D.TableDefs(T).OpenRecordset.Fields
End Function

Function FdzTF(D As Database, T, F$) As DAO.Field2
Set FdzTF = D.TableDefs(T).Fields(F)
End Function

Function FnyT(T) As String()
FnyT = Fny(CurrentDb, T)
End Function

Function Fny(D As Database, T) As String()
Fny = Itn(DbzReOpn(D).TableDefs(T).Fields)
End Function

Function FnyzT(D As Database, T) As String()
FnyzT = Fny(D, T)
End Function

Function FstUniqIdx(D As Database, T) As DAO.Index
Set FstUniqIdx = FstzItrT(D.TableDefs(T).Indexes, "Unique")
End Function

Function HasFld(D As Database, T, F$) As Boolean
HasFld = HasItn(D.TableDefs(T).Fields, F)
End Function

Function HasId(D As Database, T, Id&) As Boolean
If HasPk(D, T) Then
    HasId = HasRec(RszId(D, T, Id))
End If
End Function

Function HasIdx(D As Database, T, IdxNm$) As Boolean
HasIdx = HasItn(D.TableDefs(T).Indexes, IdxNm)
End Function

Function HasPk(D As Database, T) As Boolean
HasPk = HasPkzTd(D.TableDefs(T))
End Function

Function HasPkzTd(A As DAO.TableDef) As Boolean
HasPkzTd = HasTruePrp(A.Indexes, "Primary")
End Function

Function HasSk(D As Database, T) As Boolean
HasSk = Not IsNothing(SkIdx(D, T))
End Function

Function HasStdPk(D As Database, T) As Boolean
HasStdPk = HasStdPkzTd(D.TableDefs(T))
End Function

Function HasStdPkzTd(A As DAO.TableDef) As Boolean
If Not HasPkzTd(A) Then Exit Function
Dim Pk$(): Pk = PkFnyzTd(A): If Si(Pk) <> 1 Then Exit Function
Dim P$: P = A.Name & "Id"
If Pk(0) <> P Then Exit Function
HasStdPkzTd = A.Fields(0).Name <> P
End Function

Function HasStdSkzTd(A As DAO.TableDef) As Boolean
If Not HasItn(A.Indexes, Skn) Then Exit Function
HasStdSkzTd = A.Indexes(Skn).Unique = True
End Function

Function Idx(D As Database, T, IdxNm$) As DAO.Index
Set Idx = IdxzTd(D.TableDefs(T), IdxNm)
End Function

Function IdxzTd(A As DAO.TableDef, IdxNm$) As DAO.Index
Set IdxzTd = FstzItn(A.Indexes, IdxNm$)
End Function


Function IntAyzDbtf(D As Database, T, F$) As Integer()
Dim Q$: Q = FmtQQ("Select [?] from [?]", F, T)
IntAyzDbtf = IntAyzQ(D, Q)
End Function

Function IsHidTbl(D As Database, T) As Boolean
IsHidTbl = (D.TableDefs(T).Attributes And DAO.TableDefAttributeEnum.dbHiddenObject) <> 0
End Function

Function IsLnk(D As Database, T) As Boolean
IsLnk = IsLnkzFb(D, T) Or IsLnkzFx(D, T)
End Function

Function IsLnkzFb(D As Database, T) As Boolean
IsLnkzFb = HasPfx(CnStrzT(D, T), ";Database=")
End Function

Function IsLnkzFx(D As Database, T) As Boolean
IsLnkzFx = HasPfx(CnStrzT(D, T), "Excel")
End Function

Function IsMemCol(Col) As Boolean
Dim I
For Each I In Col
    If IsStr(I) Then
        If Len(I) > 255 Then IsMemCol = True: Exit Function
    End If
Next
End Function

Function IsNumzVbTy(A As VbVarType) As Boolean
Select Case A
Case vbByte, vbInteger, vbLong, vbSingle, vbDecimal, vbDouble, vbCurrency: IsNumzVbTy = True
End Select
End Function

Function IsSysTbl(D As Database, T) As Boolean
IsSysTbl = (D.TableDefs(T).Attributes And DAO.TableDefAttributeEnum.dbSystemObject) <> 0
End Function

Function IxzTF%(D As Database, T, Fld$)
Dim F As DAO.Field, O%
For Each F In D.TableDefs(T).Fields
    If F.Name = Fld Then
        IxzTF = O
        Exit Function
    End If
    O = O + 1
Next
IxzTF = -1
End Function

Function JnQSqCommaSpcAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnQSqCommaSpcAp = JnQSqCommaSpc(SyzAy(Av))
End Function

Sub KillIfTmpDb(D As Database)
If IsDbTmp(D) Then
    Dim Fb$: Fb = D.Name
    ClsDb D
    Kill Fb
End If
End Sub

Function LasUpdTim(D As Database, T) As Date
LasUpdTim = DbtPrp(D, T, "LastUpdated")
End Function

Function Lnkinf(D As Database) As String()
Dim T$, I
For Each I In Tni(D)
    T = I
    PushI Lnkinf, LnkinfzT(D, T)
Next
End Function

Function LnkinfzT$(D As Database, T)
Dim O$, LnkFx$, LnkW$, LnkFb$, LnkT$
Select Case True
Case IsLnkzFx(D, T): LnkinfzT = FmtQQ("LnkFx(?).LnkWs(?).Tbl(?).Db(?)", CnStrzDbt(D, T), SrcTn(D, T), T, D.Name)
Case IsLnkzFb(D, T): LnkinfzT = FmtQQ("LnkFb(?).LnkTbl(?).Tbl(?).Db(?)", CnStrzDbt(D, T), SrcTn(D, T), T, D.Name)
End Select
End Function

Function LofVblzDbt$(D As Database, T)
LofVblzDbt = DbtPrp(D, T, "LofVbl")
End Function

Function MaxNumVbTy(A As VbVarType, B As VbVarType) As VbVarType
If A = B Then MaxNumVbTy = A: Exit Function
If Not IsNumzVbTy(B) Then Thw CSub, "Given B is not NumVbTy", "B-VarType", B
Dim O As VbVarType
Select Case A
Case VbVarType.vbByte:      O = B
Case VbVarType.vbInteger:   O = IIf(B = vbByte, A, B)
Case VbVarType.vbLong:      O = IIf((B = vbByte) Or (B = vbInteger), A, B)
Case VbVarType.vbSingle:    O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong), A, B)
Case VbVarType.vbDecimal:   O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong) Or (B = vbSingle), A, B)
Case VbVarType.vbDouble:    O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong) Or (B = vbSingle) Or (B = vbDecimal), A, B)
Case VbVarType.vbCurrency:  O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong) Or (B = vbSingle) Or (B = vbDecimal) Or (B = vbDouble), A, B)
Case Else:                  Thw CSub, "Given A is not NumVbTy", "A-VarType", A
End Select
MaxNumVbTy = O
End Function

Function NColzT&(D As Database, T)
NColzT = D.TableDefs(T).Fields.Count
End Function

Function NReczFxw&(Fx, Wsn, Optional Bexp$)
NReczFxw = FvzCnq(CnzFx(Fx), SqlSelCnt_T_OB(CattnzWsn(Wsn), Bexp))
End Function

Function NReczT&(D As Database, T, Optional Bexp$)
NReczT = FvzQ(D, SqlSelCnt_T_OB(T, Bexp))
End Function

Function NReczTBexp&(D As Database, T, Bexp$)
NReczTBexp = FvzQ(D, SqlSelCnt_T_OB(T, Bexp))
End Function

Function NxtId&(D As Database, T)
Dim S$: S = FmtQQ("select Max(?Id) from [?]", T, T)
NxtId = FvzQ(D, S) + 1
End Function

Function PkFny(D As Database, T) As String()
PkFny = FnyzIdx(PkIdx(D, T))
End Function

Function PkFnyT(T) As String()
PkFnyT = PkFny(CurrentDb, T)
End Function

Function PkFnyzTd(A As DAO.TableDef) As String()
PkFnyzTd = FnyzIdx(PkizTd(A))
End Function

Function PkIdx(D As Database, T) As DAO.Index
Set PkIdx = PkizTd(D.TableDefs(T))
End Function

Function PkIdxNm$(D As Database, T)
PkIdxNm = ObjNm(PkIdx(D, T))
End Function

Function PkizTd(A As DAO.TableDef) As DAO.Index
Set PkizTd = FstzItn(A.Indexes, Pkn)
End Function

Sub RenFld(D As Database, T, F$, ToFld$)
D.TableDefs(T).Fields(F).Name = ToFld
End Sub

Sub RenTblzAddPfx(D As Database, T, Pfx$)
RenTbl D, T, Pfx & T
End Sub


Function RszId(D As Database, T, Id&) As DAO.Recordset
Set RszId = Rs(D, SqlSel_T_WhId(T, Id))
End Function

Function RszT(D As Database, T) As DAO.Recordset
Set RszT = Rs(D, SqlSel_T(T))
End Function

Function RszTF(D As Database, T, F$) As DAO.Recordset
Set RszTF = D.OpenRecordset(SqlSel_F_T(F, T))
End Function

Function RszTFF(D As Database, T, FF$) As DAO.Recordset
Set RszTFF = RszTFny(D, T, Ny(FF))
End Function

Function RszTFny(D As Database, T, Fny$()) As DAO.Recordset
Set RszTFny = D.OpenRecordset(SqlSel_Fny_T(Fny, T))
End Function

Sub SetLofVblzDbt(D As Database, T, LofVbl$)
DbtPrp(D, T, "LofVbl") = LofVbl
End Sub

Function ShtTyszCol$(Col())
Dim O$, I
I = Itr(Col)
Select Case True
Case IsItrBool(I): O = "B"
Case IsItrDte(I): O = "Dte"
Case IsItrNum(I): O = ShtTyzNumCol(Col)
Case IsItrStr(I): O = IIf(IsMemCol(Col), "M", "")
Case Else: Thw CSub, "Col cannot determine its type: Not [Str* Num* Bool* Dte*:Col]", "Col", Col
End Select
ShtTyszCol = O
End Function

Function ShtTyzNumCol$(Col)
ShtTyzNumCol = ShtDaoTy(DaoTyzNumCol(Col))
End Function

Function ShtTyzNumVbTy$(NumVbTy As VbVarType)
Dim O$
Select Case NumVbTy
Case VbVarType.vbByte:      O = "Byt:"
Case VbVarType.vbCurrency:  O = "C:"
Case VbVarType.vbDecimal:   O = "Dec:"
Case VbVarType.vbDouble:    O = "D:"
Case VbVarType.vbInteger:   O = "I:"
Case VbVarType.vbLong:      O = "L:"
Case VbVarType.vbSingle:    O = "S:"
Case Else: Thw CSub, "NumVbTy is not numeric VbTy", "NumVbTyp", ShtTyzNumVbTy(NumVbTy)
End Select
End Function

Function ShtTyzTF$(D As Database, T, F$)
ShtTyzTF = ShtDaoTy(DaoTyzTF(D, T, F$))
End Function

Function SqzT(D As Database, T, Optional ExlFldNm As Boolean) As Variant()
SqzT = SqzRs(RszT(D, T), ExlFldNm)
End Function

Function SrcFbzT$(D As Database, T)
SrcFbzT = Bet(D.TableDefs(T).Connect, "Database=", ";")
End Function

Function SrcTn$(D As Database, T)
SrcTn = D.TableDefs(T).SourceTableName
End Function


Sub Z_CrtDupKeyTbl()
Dim D As Database: Set D = TmpDb
DrpTT D, "#A #B"
'T = "AA"
CrtTblzDup D, "#A", "#B", "Sku BchNo"
DrpDbIfTmp D
End Sub

Sub Z_CrtTblzDrs()
Dim D As Database
GoSub Z
Exit Sub
Z:
    Set D = TmpDb
    DrpTmpTbl D
    CrtTblzDrs D, "#D", SampDrs
    BrwDb D
    Return
End Sub

Sub Z_PkFny()
Z:
    Dim D As Database
    Set D = Db(SampFbzDutyDta)
    Dim Dr(), Dy(), T, I
    For Each I In Tny(D)
        T = I
        Erase Dr
        Push Dr, T
        PushIAy Dr, PkFny(D, T)
        PushI Dy, Dr
    Next
    BrwDy Dy
    Exit Sub
End Sub

Sub Z_ShtTyscfBqlzDrs()
Dim Drs As Drs
GoSub T0
Exit Sub
T0:
    Drs = SampDrs
    Ept = "A`B:B`Byt:C`I:D`L:E`D:G`S:H`C:I`Dte:J`M:K"
    GoTo Tst
Tst:
    Act = ShtTyscfBqlzDrs(Drs)
    C
    Return
End Sub
