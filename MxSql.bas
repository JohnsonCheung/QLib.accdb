Attribute VB_Name = "MxSql"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxSql."
Type SelIntoPm: Fny() As String: Ey() As String: Into As String: T As String: Bexp As String: End Type
Type SelIntoPms: N As Byte: Ay() As SelIntoPm: End Type

Function AddPfxNLTT$(Sy$())
AddPfxNLTT = Jn(AmAddPfx(Sy, C_NLTT), "")
End Function

Function Bexp_E_InLis$(Expr$, InLisStr$)
If InLisStr = "" Then Exit Function
Bexp_E_InLis = FmtQQ("? in (?)", Expr, InLisStr)
End Function

Function Bexp_F_Ev$(F$, Ev)
Bexp_F_Ev = QteSq(F) & "=" & SqlQte(Ev)
End Function

Function Bexp_Fny_Vy$(Fny$(), Vy())
End Function

Sub FEs_AddAs_Or4Spc_ToExtNm(OE$())
Dim J%, C$
For J = 0 To UB(OE)
    
    If Trim(OE(J)) = "" Then
        C = "    "
    Else
        C = " As "
    End If
    OE(J) = OE(J) & C
Next
End Sub

Sub FEs_AddTab2Spc_ToExtNm(OE$())
OE = AmAddPfx(OE, C_T & "  ")
End Sub

Sub FEs_AlignExtNm(OE$())
OE = AlignAy(OE)
End Sub

Sub FEs_AlignFld(OF$())
OF = AlignAy(OF)
End Sub

Sub FEs_SetExtNm_ToBlnk_IfEqToFld(F$(), OE$())
Dim J%
For J = 0 To UB(OE)
    If OE(J) = F(J) Then OE(J) = ""
Next
End Sub

Sub FEs_SqQteExtNm_IfNB(OE$())
Dim J%
For J = 0 To UB(OE)
    If OE(J) <> "" Then
        OE(J) = QteSq(OE(J))
    End If
Next
End Sub

Property Get IsFmt() As Boolean
Static X As Boolean, Y As Boolean
If Not X Then X = True: Y = Cfg.Sql.FmtSql
IsFmt = Y
End Property

Function FmtExprVblAy(ExprVblAy$(), Optional Pfx$, Optional IdentOpt%, Optional Sep$ = ",") As String()
Ass IsVblAy(ExprVblAy)
Dim Ident%
    If IdentOpt > 0 Then
        Ident = IdentOpt
    Else
        Ident = 0
    End If
    If Ident = 0 Then
        If Pfx <> "" Then
            Ident = Len(Pfx)
        End If
    End If
Dim O$(), P$, S$, U&, J&
U = UB(ExprVblAy)
Dim W%
'    W = VblWdtAy(ExprVblAy)
For J = 0 To U
    If J = 0 Then P = Pfx Else P = ""
    If J = U Then S = "" Else S = Sep
'    Push O, VblAlign(ExprVblAy(J), IdentOpt:=Ident, Pfx:=P, WdtOpt:=W, Sfx:=S)
Next
FmtExprVblAy = O
End Function

Function FnyzPfxN(Pfx$, N%) As String()
Dim J%
For J = 1 To N
    PushI FnyzPfxN, Pfx & J
Next
End Function

Function JnCommaSpcFF$(FF$)
JnCommaSpcFF = JnQSqCommaSpc(TermAy(FF))
End Function

Function NsetzNN(NN$) As Aset
Set NsetzNN = AsetzAy(SyzSS(NN))
End Function





Sub PushIelIntoPm(O As SelIntoPms, M As SelIntoPm)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub


Function QNm$(T)
QNm = QteSq(T)
End Function

Function QV$(V)
QV = SqlQte(V)
End Function

Function SelIntoPm(Fny$(), Ey$(), Into$, T$, Optional Bexp$) As SelIntoPm
With SelIntoPm
    .Fny = Fny
    .Ey = Ey
    .Into = Into
    .T = T
    .Bexp = Bexp
End With
End Function

Function SqlAddCol_T_Fny_FzDiSqlTy$(T, Fny$(), FzDiSqlTy As Dictionary)
Dim O$(), F
For Each F In Fny
    PushI O, F & " " & VzDicK(FzDiSqlTy, F, "FzDiSqlTy", "Fld")
Next
SqlAddCol_T_Fny_FzDiSqlTy = FmtQQ("Alter Table [?] add column ?", T, JnComma(O))
End Function

Function SqlAddColzAy$(T, ColAy$())
SqlAddColzAy = SqlAddColzLis(T, JnCommaSpc(ColAy))
End Function

Function SqlAddColzLis$(T, ColLis$)
SqlAddColzLis = FmtQQ("Alter Table [?] add column ?", T, ColLis)
End Function

Function PkSql$(T)
PkSql = FmtQQ("Create Index PrimaryKey on [?] (?Id) with Primary", T, T)
End Function

Function PkSqy(Tny$()) As String()
Dim T: For Each T In Itr(Tny)
    PushI PkSqy, PkSql(T)
Next
End Function

Function SkSqlzFF$(T, Skff$)
SkSqlzFF = SkSql(T, Ny(Skff))
End Function

Function SkSqy(Tny$(), SkFnyAy()) As String()
Dim J%: For J = 0 To UB(Tny)
    Dim SkFny$(): SkFny = SkFnyAy(J)
    PushI SkSqy, SkSql(Tny(J), SkFny)
Next
End Function

Function SkSql$(T, SkFny$())
SkSql = FmtQQ("Create unique Index SecondaryKey on [?] (?)", T, JnComma(SyzQteSq(SkFny)))
End Function

Function SqlCrtTbl_T_X$(T, X$)
SqlCrtTbl_T_X = FmtQQ("Create Table [?] (?)", T, X)
End Function

Function SqlDlt$(T, Bexp$)
SqlDlt = SqlDlt_T(T) & PWh(Bexp)
End Function

Function SqlDlt_T$(T)
SqlDlt_T = "Delete * from [" & T & "]"
End Function

Function SqlDrpCol_T_F$(T, F$)
SqlDrpCol_T_F = FmtQQ("Alter Table [?] drop column [?]", T, F$)
End Function

Function SqlDrpFld$(T, Fny$())
Dim S$: S = JnCommaSpc(SyzQteSq(Fny))
SqlDrpFld = "Alter Table [" & T & "] drop column " & S
End Function

Function SqlDrpTbl_T$(T)
SqlDrpTbl_T = "Drop Table [" & T & "]"
End Function

Function SqlIns_T_FF_Dr$(T, FF$, Dr)
Dim Fny$(): Fny = SyzSS(FF)
ThwIf_DifSi Fny, Dr, CSub
Dim A$, B$
A = JnComma(SyzQteSqIf(Fny))
B = JnComma(SqlQteVy(Dr))
SqlIns_T_FF_Dr = FmtQQ("Insert Into [?] (?) Values(?)", T, A, B)
End Function

Function SqlIns_T_FF_ValAp$(T, FF$, ParamArray ValAp())
Dim Av(): Av = ValAp
SqlIns_T_FF_ValAp = PIns_T(T) & PBkt_FF(FF) & " Values" & PBkt_Av(Av)
End Function

Function SqlQte$(V)
Dim O$, C$
C = SqlQteChr(V)
If C <> "" Then SqlQte = Qte(CStr(V), C): Exit Function
Select Case True
Case IsBool(V): O = IIf(V, "true", "false")
Case IsEmpty(V), IsNull(V), IsNothing(V): O = "null"
Case Else: O = V
End Select
SqlQte = O
End Function

Function SqlQteChr$(V)
Dim O$
Select Case True
Case IsStr(V): O = "'"
Case IsDate(V): O = "#"
End Select
SqlQteChr = O
End Function

Function SqlQteChrzT$(A As DAO.DataTypeEnum)
Select Case A
Case _
    DAO.DataTypeEnum.dbBigInt, _
    DAO.DataTypeEnum.dbByte, _
    DAO.DataTypeEnum.dbCurrency, _
    DAO.DataTypeEnum.dbDecimal, _
    DAO.DataTypeEnum.dbDouble, _
    DAO.DataTypeEnum.dbFloat, _
    DAO.DataTypeEnum.dbInteger, _
    DAO.DataTypeEnum.dbLong, _
    DAO.DataTypeEnum.dbNumeric, _
    DAO.DataTypeEnum.dbSingle: Exit Function
Case _
    DAO.DataTypeEnum.dbChar, _
    DAO.DataTypeEnum.dbMemo, _
    DAO.DataTypeEnum.dbText: SqlQteChrzT = "'"
Case _
    DAO.DataTypeEnum.dbDate: SqlQteChrzT = "#"
Case Else
    Thw CSub, "Invalid DaoTy", "DaoTy", A
End Select
End Function

Function SqlQteVy(Vy) As String()
Dim V
For Each V In Vy
    PushI SqlQteVy, SqlQte(V)
Next
End Function

Function SqlSel_Dist_Fny_EDict_Into_T_Wh_Gp_Ord$(IsDist As Boolean, Fny$(), EDic As Dictionary, T$, Wh$, Gp$, Ord$)

End Function

Function SqlSel_F$(F$)
SqlSel_F = SqlSel_F_T(F, F)
End Function

Function SqlSel_F_T$(F$, T, Optional Bexp$)
SqlSel_F_T = FmtQQ("Select [?] from [?]?", F, T, PWh(Bexp))
End Function

Function SqlSel_F_T_F_Ev$(F$, T, WhFld$, Ev())
SqlSel_F_T_F_Ev = SqlSel_F_T(F, T, PExpr_F_InAy(WhFld, Ev))
End Function

Function SqlSel_FF_EDic_Into_T_OB$(FF$, EDic As Dictionary, Into, T, Optional Bexp$)
Dim Fny$(): Fny = SyzSS(FF)
Dim ExprAy$(): ExprAy = SyzDicKy(EDic, Fny)
Stop
SqlSel_FF_EDic_Into_T_OB = SqlSel_Fny_Extny_Into_T_OB(Fny, ExprAy, Into, T, Bexp)
End Function

Function SqlSel_FF_ExprDic_T$(FF$, E As Dictionary, T, Optional IsDis As Boolean)
'SelFFExprDicP = "Select" & vbCrLf & FFExprDicAsLines(FF$, ExprDic)
End Function

Function SqlSel_FF_X_Wh$(FF$, X$, Bexpr$)
SqlSel_FF_X_Wh = PSel_FF(FF) & PFm_X(X) & PWh(Bexpr)
End Function

Function SqlSel_FF_T$(FF, T, Optional IsDis As Boolean)
SqlSel_FF_T = PSel_FF(FF, IsDis) & PFm(T)
End Function

Function SqlSel_FF_T_Bexp$(FF$, T, Bexp$)

End Function

Function SqlSel_FF_T_Ord(FF$, T, OrdMinusSfxFF$)
SqlSel_FF_T_Ord = PSel_FF(FF) & PFm(T) & POrd_MinusSfxFF(OrdMinusSfxFF)
End Function

Function SqlSel_FF_T_Ordff$(FF$, T, OrdMinusSfxFF$)
SqlSel_FF_T_Ordff = PSel_FF(FF) & PFm(T) & POrd_MinusSfxFF(OrdMinusSfxFF)
End Function

Function SqlSel_FF_T_WhF_InVy$(FF, T, WhF$, InVy, Optional IsDis As Boolean)
Dim W$
W = PExpr_F_InAy(WhF$, InVy)
SqlSel_FF_T_WhF_InVy = SqlSel_FF_T(FF, T, IsDis)
End Function

Function SqlSel_Fny_Extny_Into_T_OB$(Fny$(), Extny$(), Into, T, Optional Bexp$)
SqlSel_Fny_Extny_Into_T_OB = PSel_Fny_Extny(Fny, Extny) & PInto_T(Into) & PFm(T) & PWh(Bexp)
End Function

Function SqlSel_Fny_Into_T_OB$(Fny$(), Into$, T, Optional Bexp$)

End Function

Function SqlSel_Fny_T(Fny$(), T, Optional Bexp$, Optional IsDis As Boolean)
SqlSel_Fny_T = PSel_Fny(Fny, IsDis) & PFm(T) & PWh(Bexp)
End Function

Function SqlSel_Fny_T_WhFny_EqVy$(Fny$(), T, WhFny$(), EqVy)
SqlSel_Fny_T_WhFny_EqVy = SqlSel_Fny_T(Fny, T, PWh_Fny_EqVy(WhFny, EqVy))
End Function

Function SqlSel_Into_T_WhFalse(Into, T)
SqlSel_Into_T_WhFalse = FmtQQ("Select * Into [?] from [?] where false", Into, T)
End Function

Function SqlSel_T$(T)
SqlSel_T = "Select *" & PFm(T)
End Function

Function SqlSel_T_Wh$(T, Bexp$)
SqlSel_T_Wh = SqlSel_T(T) & PWh(Bexp)
End Function

Function SqlSel_T_F$(T, F)
SqlSel_T_F = FmtQQ("Select [?] from [?]", F, T)
End Function

Function SqlSel_T_F_Wh$(T, F, Bexp$)
SqlSel_T_F_Wh = SqlSel_T_F(T, F) & PWh(Bexp)
End Function

Function SqlSel_T_WhId$(T, Id&)
SqlSel_T_WhId = PSel_T(T) & " " & PWh_T_Id(T, Id)
End Function

Function SqlSel_X_Into_T_OB$(X$, Into$, T$, Optional OBexp$)
SqlSel_X_Into_T_OB$ = PSelzX(X) & PInto_T(Into) & PFm(T) & PWh(OBexp)
End Function

Function SqlSel_X_Into_T_OB_OGp_OOrd_ODis$(X$, Into$, T$, Optional OBexp$, Optional OGp$, Optional OOrd$, Optional ODis As Boolean)
SqlSel_X_Into_T_OB_OGp_OOrd_ODis$ = PSelzX(X, ODis) & PInto_T(Into) & PFm(T) & PWh(OBexp) & PGp(OGp) & POrd(OOrd$)
End Function

Function SqlSel_X_Into_T_T_Jn$(X$, Into$, T1$, T2$, Jn$)

End Function

Function SqlSel_X_T$(X$, T, Optional Bexp$)
SqlSel_X_T = PSelzX(X) & PFm(T) & PWh(Bexp)
End Function

Function SqlSelCnt_T_OB$(T, Optional Bexp$)
SqlSelCnt_T_OB = "Select Count(*)" & PFm(T) & PWh(Bexp)
End Function

Function SqlSelDis_FF_T$(FF$, T)
SqlSelDis_FF_T = SqlSel_FF_T(FF$, T, IsDis:=True)
End Function

Function SqlSelzInto$(Into$, SelX$, Fm$, Optional Gp$, Optional Bexp$)
Dim Dis As Boolean: If Gp <> "" Then Dis = True
SqlSelzInto = PSelzX(SelX, Dis) & PFm(Fm) & PWh(Bexp) & PGp(Gp)
End Function

Function SqlSelzIntoCpy$(Into$, Fm$)
SqlSelzIntoCpy = PSelzX("*") & PInto_T(Into) & PFm(Fm)
End Function

Function SqlSelzIntoFF$(Into$, FF$, Fm$, Optional OBexp$, Optional ODis As Boolean)
SqlSelzIntoFF = PSel_FF(FF, ODis) & PInto_T(Into) & PFm(Fm) & PWh(OBexp)
End Function

Function SqlSelzIntoFmX$(Into$, SelX$, FmX$, Optional Gp$, Optional Bexp$)
Dim Dis As Boolean: If Gp <> "" Then Dis = True
SqlSelzIntoFmX = PSelzX(SelX, Dis) & PInto_T(Into) & PFmzX(FmX) & PWh(Bexp) & PGp(Gp)
End Function

Function SqlUpd_T_FF_EqDr_Whff_Eqvy$(T, FF$, Dr, WhFF$, EqVy)
SqlUpd_T_FF_EqDr_Whff_Eqvy = PUpd(T) & PSet_FF_EqDr(FF, Dr) & PWh_FF_Eqvy(WhFF, EqVy)
End Function

Function SqlUpd_T_Sk_Fny_Dr$(T, Sk$(), Fny$(), Dr)
If Si(Sk) = 0 Then Stop
Dim PUpd$, Set_$, Wh$: GoSub X_PUpd_Set_Wh
'UpdSql = PUpd & Set_ & Wh
Exit Function
X_PUpd_Set_Wh:
    Dim Fny1$(), Dr1(), Skvy(): GoSub X_Fny1_Dr1_SkVy
    PUpd = "Update [" & T & "]"
    Set_ = PSet_Fny_Vy(Fny1, Dr1)
    Wh = PWh_Fny_EqVy(Sk, Skvy)
    Return
X_Ay:
    Dim L$(), R$()
    L = AlignQteSq(Fny)
    R = SqlQteVy(Dr)
    Return
X_Fny1_Dr1_SkVy:
    Dim Ski, J%, Ixy%(), I%
    For Each Ski In Sk
'        I = IxzAy(Fny, Ski)
        If I = -1 Then Stop
        Push Ixy, I
        Push Skvy, Dr(I)    '<====
    Next
    Dim F
    For Each F In Fny
        If Not HasEle(Ixy, J) Then
            Push Fny1, F        '<===
            Push Dr1, Dr(J)     '<===
        End If
        J = J + 1
    Next
    Return
End Function

Function SqlUpdzEy$(T, Fny$(), Ey$(), Optional OBexp$)
SqlUpdzEy = PUpd(T) & PSet_Fny_Ey(Fny, Ey) & PWh(OBexp)
End Function

Function SqlUpdzJn$(T$, FmA$, JnFny$(), SetFny$())
'Fm T     : Table nm to be update.  It will have alias x.
'Fm FmA   : Table nm used to update @T.  It will has alias a.
'Fm JnFny : Fld nm common in @T & @FmA.  It will use to bld the jn clause with alias x and a.
'Fm SetX  : Fny in @T to be updated.  No alias, by the ret sql will put the alias x.  Sam ele as @EqA.
'Ret      : upd sql stmt updating @T from @FmA using @JnFny as jn clause setting @T fld as stated in @SetX eq to @FmA fld as stated in @EqA
Dim U$: U = PUpdzXAJn(T, FmA, JnFny)
Dim S$: S = PSetzXAFny(SetFny)
SqlUpdzJn = U & C_NL & S
End Function

Function SqlUpdzXSet$(TblX$, SetX$)
SqlUpdzXSet = PUpdzX(TblX) & PSetzX(SetX)
End Function

Function SqlzSelIntoPm$(A As SelIntoPm)
With A
'SqlzSelIntoPm = SqlSel_Fny_Extny_Into_T(.Fny, .Extny, .Into, .T, .Bexp)
End With
End Function

Function SqpAEqB_Fny_AliasAB$(Fny$(), Optional AliasAB$ = "x a")
Dim A1$: A1 = BefSpc(AliasAB) ' Alias1
Dim A2$: A2 = BefSpc(AliasAB) ' Alias2
Dim A$(): A = AmAddPfx(Fny, A1 & ".")
Dim B$(): B = AmAddPfx(Fny, A2 & ".")
Dim J$(): J = LyzAyab(A, B, " = ")
SqpAEqB_Fny_AliasAB = JnCommaSpc(J)
End Function

Function SqyCrtPkzTny(Tny$()) As String()
Dim T
For Each T In Itr(Tny)
    PushI SqyCrtPkzTny, PkSql(T)
Next
End Function

Function SqyDlt_T_WhFld_InAset(T, F, S As Aset, Optional SqlWdt% = 3000) As String()
Dim A$
Dim Ey$()
    A = SqlDlt_T(T) & " Where "
    Ey = PFldInX_F_InAset_Wdt(F, S, SqlWdt - Len(A))
Dim E
For Each E In Ey
    PushI SqyDlt_T_WhFld_InAset, A & E & vbCrLf
Next
End Function

Function SqyzSelIntoPms(A As SelIntoPms) As String()
Dim J As Byte
For J = 0 To A.N - 1
    PushI SqyzSelIntoPms, SqlzSelIntoPm(A.Ay(J))
Next
End Function

Sub Z_SqlSel_Fny_Ey_Into_T_OB()
Dim Fny$(), Ey$(), Into$, T$, Bexp$
GoSub Z
Exit Sub
Z:
    Fny = SyzSS("Sku CurRateAc VdtFm VdtTo HKD Per CA_Uom")
    Ey = TermAy("Sku [     Amount] [Valid From] [Valid to] Unit per Uom")
    Into = "#IZHT086"
    T = ">ZHT086"
    Bexp = ""
    Debug.Print SqlSel_Fny_Extny_Into_T_OB(Fny, Ey, Into, T, Bexp)
    Return
End Sub
