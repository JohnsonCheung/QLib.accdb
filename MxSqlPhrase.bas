Attribute VB_Name = "MxSqlPhrase"
Option Explicit
Option Compare Text
Const CLib$ = "QSql."
Const CMod$ = CLib & "MxSqlPhrase."
Const KwBet$ = "between"
Const KwSet$ = "set"
Const KwDis$ = "distinct"
Const KwUpd$ = "update"
Const KwInto$ = "into"
Const KwSel$ = "select"
Const KwFm$ = "from"
Const KwGp$ = "group by"
Const KwWh$ = "where"
Const KwAnd$ = "and"
Const KwOn$ = "on"
Const KwLJn$ = "left join"
Const KwIJn$ = "inner join"
Const KwOr$ = "or"
Const KwOrd$ = "order by"
Const KwLeftJn$ = "left join"
Function PFm$(Fm)
PFm = C_Fm & QteSq(Fm)
End Function

Function PFm_X$(X)
PFm_X = C_Fm & X
End Function

Function PFmzX$(FmX$)
PFmzX = PFm(FmX) & " x"
End Function

Function PGp$(Gp$)
If Gp = "" Then Exit Function
PGp = C_T & KwGp & C_T & Gp
End Function

Function PGp_ExprVblAy$(ExprVblAy$())
PGp_ExprVblAy = "|  Group By " & JnCrLf(FmtExprVblAy(ExprVblAy))
End Function

Function PIns_T$(T)
PIns_T = "Insert into [" & T & "]"
End Function

Function PInto_T$(T)
PInto_T = C_Into & "[" & T & "]"
End Function

Function PAnd_Bexp$(Bexp$)
If Bexp = "" Then Exit Function
'PAnd_Bexp = NxtLin & "and " & NxtLin_Tab & Bexp
End Function

Function PBexp_Fny_EqVy$(Fny$(), EqVy)

End Function

Function PBkt_Av$(Av())
Dim O$(), I
For Each I In Av
    PushI O, SqlQte(I)
Next
PBkt_Av = QteBktJnComma(Av)
End Function

Function PBkt_FF$(FF$)
PBkt_FF = QteBkt(SyzSS(FF))
End Function

Function PDis$(Dis As Boolean)
If Dis Then PDis = " " & KwDis
End Function

Function PExpr_F_InAy$(F, InVy)

End Function

Function PExpr_T_RecId$(T, RecId)
PExpr_T_RecId = FmtQQ("?Id=?", T, RecId)
End Function

Function PFldInX_F_InAset_Wdt(F, S As Aset, Wdt%) As String()
Dim A$
    A = "[F] in ("
Dim I
'For Each I In LyJnQSqlCommaAsetW(S, Wdt - Len(A))
    PushI PFldInX_F_InAset_Wdt, I
'Next
End Function


Function POnzJnXA(JnFny$())
Dim X$(): X = SyzQAy("x.[?]", JnFny)
Dim A$(): A = SyzQAy("a.[?]", JnFny)
Dim J$(): J = LyzAyab(X, A, " = ")
Dim S$: S = JnAnd(J)
POnzJnXA = KwOn & " " & S
End Function

Function POrd$(Ord$)

End Function

Function POrd_MinusSfxFF$(OrdMinusSfxFF$)
If OrdMinusSfxFF = "" Then Exit Function
Dim O$(): O = SyzSS(OrdMinusSfxFF)
Dim I, J%
For Each I In O
    If HasSfx(O(J), "-") Then
        O(J) = RmvSfx(O(J), "-") & " desc"
    End If
    J = J + 1
Next
POrd_MinusSfxFF = C_NLT & "order by " & JnCommaSpc(O)
End Function

Function PSel_F$(F$)
PSel_F = "Select [" & F & "]"
End Function

Function PSel_FF$(FF, Optional Dis As Boolean)
PSel_FF = PSel_Fny(SyzSS(FF), Dis)
End Function

Function PSel_FF_Extny$(FF$, Extny$())
PSel_FF_Extny = PSelzX(PSel_Fny_Extny(Ny(FF), Extny))
End Function

Function PSel_Fny$(Fny$(), Optional Dis As Boolean)
PSel_Fny = KwSel & PDis(Dis) & C_NLTT & JnCommaSpc(Fny)
End Function

Function PSel_Fny_Extny$(Fny$(), Extny$(), Optional IsDis As Boolean)
If Not IsFmt Then PSel_Fny_Extny = PSel_Fny_Extny_NOFMT(Fny, Extny): Exit Function
Dim E$(), F$()
F = Fny
E = Extny
FEs_SetExtNm_ToBlnk_IfEqToFld F, E
FEs_SqQteExtNm_IfNB E
FEs_AlignExtNm E
FEs_AddAs_Or4Spc_ToExtNm E
FEs_AddTab2Spc_ToExtNm E
FEs_AlignFld F
PSel_Fny_Extny = KwSel & PDis(IsDis) & C_NL & Join(LyzAyab(E, F), C_CNL)
End Function

Function PSel_Fny_Extny_NOFMT(Fny$(), Extny$(), Optional IsDis As Boolean)
Dim O$(), J%, E$, F$
For J = 0 To UB(Fny)
    F = Fny(J)
    E = Trim(Extny(J))
    Select Case True
    Case E = "", E = F: PushI O, F
    Case Else: PushI O, QteSq(E) & " As " & F
    End Select
Next
PSel_Fny_Extny_NOFMT = KwSel & PDis(IsDis) & " " & JnCommaSpc(O)
End Function

Function PSel_T$(T)
PSel_T = KwSel & C_T & "*" & PFm(T)
End Function

Function PSelzX$(X$, Optional Dis As Boolean)
PSelzX = KwSel & PDis(Dis) & X
End Function

Function PSet_FF_EqDr$(FF$, EqDr)

End Function

Function PSet_FF_ExprAy$(FF, Ey$())
Const CSub$ = CMod & "PSet_FF_Ey"
Dim Fny$(): Fny = SyzSS(FF)
Ass IsVblAy(Ey)
If Si(Fny) <> Si(Ey) Then Thw CSub, "[FF-Sz} <> [Si-Ey], where [FF],[Ey]", Si(Fny), Si(Ey), FF, Ey
Dim AFny$()
    AFny = AlignAy(Fny)
    AFny = AmAddSfx(AFny, " = ")
Dim W%
    'W = VblWdtAy(Ey)
Dim Ident%
    W = WdtzAy(AFny)
Dim Ay$()
    Dim J%, U%, S$
    U = UB(AFny)
    For J = 0 To U
        If J = U Then
            S = ""
        Else
            S = ","
        End If
        'Push Ay, VblAlign(Ey(J), Pfx:=AFny(J), IdentOpt:=Ident, WdtOpt:=W, Sfx:=S)
    Next
Dim Vbl$
    Dim Ay1$()
    Dim P$
    For J = 0 To U
        If J = 0 Then P = "|  Set" Else P = ""
'        Push Ay1, VblAlign(Ay(J), Pfx:=P, IdentOpt:=6)
    Next
    Vbl = JnVBar(Ay1)
PSet_FF_ExprAy = Vbl
End Function

Function PSet_FF_Ey$(FF$, Ey$())
PSet_FF_Ey = PSet_Fny_Ey(SyzSS(FF), Ey)
End Function

Function PSet_Fny_Evy$(Fny$(), EqVy)

End Function

Function PSet_Fny_Ey$(Fny$(), Ey$())
Dim J$(): J = LyzAyab(SyzQteSq(Fny), Ey, " = ")
Dim J1$(): J1 = AmAddPfx(J, C_TT)
Dim S$: S = Jn(J, "," & C_NL)
PSet_Fny_Ey = C_NLT & KwSet & C_NL & S
End Function

Function PSet_Fny_Vy$(Fny$(), Vy())
Dim F$(): F = SyzQteSq(Fny)
Dim V$(): V = SqlQteVy(Vy)
PSet_Fny_Vy = JnComma(LyzAyab(F, V, "="))
End Function

Function PSet_Fny_Vy1$(Fny$(), Vy())
Dim A$: GoSub X_A
PSet_Fny_Vy1 = "  Set " & A
Exit Function
X_A:
    Dim L$(): L = SyzQteSq(Fny)
    Dim R$(): R = SqlQteVy(Vy)
    Dim J%, O$()
    For J = 0 To UB(L)
        Push O, L(J) & " = " & R(J)
    Next
    A = JnComma(O)
    Return
End Function

Function PSetzX$(SetX$)
PSetzX = C_T & KwSet & C_NL & SetX
End Function

Function PSetzXA(FnyX$(), FnyA$())
Dim X$(): X = AmAddPfxS(FnyX, "x.[", "]")
Dim A$(): A = AmAddPfxS(FnyA, "a.[", "]")
Dim J$(): J = LyzAyab(X, A, " = ")
          J = AmAddPfx(J, C_TT)
Dim S$:   S = Jn(J, "," & C_NL)
PSetzXA = PSetzX(S)
End Function

Function PSetzXAFny(Fny$())
PSetzXAFny = PSetzXA(Fny, Fny)
End Function

Function PTblzXAJn$(TblX$, TblA$, JnFny$())
PTblzXAJn = C_TT & "[" & TblX & "] x" & C_NLTT & KwIJn & " [" & TblA & "] a " & POnzJnXA(JnFny)
End Function

Function PUpd$(T)
PUpd = KwUpd & C_T & QNm(T)
End Function

Function PUpdzX$(TblX$)
PUpdzX = KwUpd & C_NL & TblX
End Function

Function PUpdzXAJn$(TblX$, TblA$, JnFny$())
Dim X$: X = PTblzXAJn(TblX, TblA, JnFny)
PUpdzXAJn = PUpdzX(X)
End Function

Function PWh$(Bexp$)
If Bexp = "" Then Exit Function
PWh = C_Wh & Bexp
End Function

Function PWh_F_Eqv(F$, EqVal) ' Ssk is single-Sk-value
PWh_F_Eqv = C_Wh & QNm(F) & "=" & QV(EqVal)
End Function

Function PWh_F_InVy$(F$, InVy)
PWh_F_InVy = C_Wh & PExpr_F_InAy(F, InVy)
End Function

Function PWh_FF_Eqvy$(FF$, EqVy)

End Function

Function PWh_Fny_EqVy$(Fny$(), EqVy)
PWh_Fny_EqVy = C_Wh & PBexp_Fny_EqVy(Fny, EqVy)
End Function

Function PWh_T_EqK$(T, K&)
PWh_T_EqK = PWh_F_Eqv(T & "Id", K)
End Function

Function PWh_T_Id$(T, Id)
PWh_T_Id = PWh(FmtQQ("[?]Id=?", T, Id))
End Function

Function PWhBet_F_Fm_To$(F$, FmV, ToV)
PWhBet_F_Fm_To = C_Wh & QNm(F) & " " & KwBet & QV(FmV) & " " & KwAnd & " " & QV(ToV)
End Function

Sub Z_PGp_ExprVblAy()
Dim ExprVblAy$()
    Push ExprVblAy, "1lskdf|sdlkfjsdfkl sldkjf sldkfj|lskdjf|lskdjfdf"
    Push ExprVblAy, "2dfkl sldkjf sldkdjf|lskdjfdf"
    Push ExprVblAy, "3sldkfjsdf"
DmpAy SplitVBar(PGp_ExprVblAy(ExprVblAy))
End Sub

Sub Z_PSel()
Dim Fny$(), ExprVblAy$()
ExprVblAy = Sy("F1-Expr", "F2-Expr   AA|BB    X|DD       Y", "F3-Expr  x")
Fny = SplitSpc("F1 F2 F3xxxxx")
'Debug.Print LineszVbl(PSelFFFldLvs(Fny, ExprVblAy))
End Sub

Sub Z_PSel_Fny_Extny()
Dim Fny$()
Dim Extny$()
GoSub Z
Exit Sub
Z:
    Fny = SyzSS("Sku CurRateAc VdtFm VdtTo HKD Per CA_Uom")
    Extny = TermAy("Sku [     Amount] [Valid From] [Valid to] Unit per Uom")
    Debug.Print PSel_Fny_Extny(Fny, Extny)
    Return
End Sub

Sub Z_PSet_Fny_VyFmt()
Dim Fny$(), Vy()
Ept = LineszVbl("|  Set|" & _
"    [A xx] = 1                     ,|" & _
"    B      = '2'                   ,|" & _
"    C      = #2018-12-01 12:34:56# ")
Fny = TermAy("[A xx] B C"): Vy = Array(1, "2", #12/1/2018 12:34:56 PM#): GoSub Tst
Exit Sub
Tst:
    Act = PSet_Fny_Vy(Fny, Vy)
    C
    Return
End Sub

Sub Z_PSetFFEqvy()
Dim Fny$(), ExprVblAy$()
Fny = SyzSS("a b c d")
Push ExprVblAy, "1sdfkl|lskdfj|skldfjskldfjs dflkjsdf| sdf"
Push ExprVblAy, "2sdfkl|lskdfjdf| sdf"
Push ExprVblAy, "3sdfkl|fjskldfjs dflkjsdf| sdf"
Push ExprVblAy, "4sf| sdf"
    Act = PSet_Fny_Evy(Fny, ExprVblAy)
'Debug.Print LineszVbl(Act)
End Sub

Sub Z_PWh_F_InVy()
Dim F$, Vy()
F = "A"
Vy = Array(1, "2", #2/1/2017#)
Ept = " where A=1 and B='2' and C=#2017-2-1#"
GoSub Tst
Exit Sub
Tst:
    Act = PWh_F_InVy(F, Vy)
    C
    Return
End Sub

Sub Z_PWhFldInVy_StrPAy()

End Sub



Property Get C_And$()
If IsFmt Then
    C_And = C_NLT & KwAnd & C_T
Else
    C_And = " " & KwAnd & " "
End If
End Property

Property Get C_CNL$()
C_CNL = "," & vbCrLf  'Comma-NewLin-Tab
End Property

Property Get C_CNLT$()
C_CNLT = "," & vbCrLf & C_T  'Comma-NewLin-Tab
End Property

Property Get C_Comma$()
If IsFmt Then
    C_Comma = "," & vbCrLf
Else
    C_Comma = ", "
End If
End Property

Property Get C_CommaSpc$()
If IsFmt Then
    C_CommaSpc = C_CNLT
Else
    C_CommaSpc = ", "
End If
End Property

Property Get C_Fm$()
C_Fm = C_NLT & KwFm & C_T
End Property

Property Get C_Into$()
C_Into = C_NLT & KwInto & C_T
End Property

Property Get C_NL$() ' New Line
If IsFmt Then
    C_NL = vbCrLf
Else
    C_NL = " "
End If
End Property

Property Get C_NLT$() ' New Line Tabe
If IsFmt Then
    C_NLT = C_NL & C_T
Else
    C_NLT = " "
End If
End Property

Property Get C_NLTT$() ' New Line Tabe
If IsFmt Then
    C_NLTT = C_NLT & C_T
Else
    C_NLTT = " "
End If
End Property

Property Get C_T$()
If IsFmt Then
    C_T = "    "
Else
    C_T = " "
End If
End Property

Property Get C_TT$()
C_TT = C_T & C_T
End Property

Property Get C_Wh$()
C_Wh = C_NLT & KwWh & C_T
End Property

