Attribute VB_Name = "MxEoCol"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxEoCol."
Const MsgoCol_Dup$ = "Lno(?) has Dup-?[?]"
Const MsgoCol_NotIn$ = "Lno(?) has ?[?] which is invalid.  Valid-?=[?]"
Const MsgoCol_NotNum$ = "Lno(?) has non-numeric-?[?]"
Const MsgoColx_Blnk$ = "Lno(?) has a blank [?] value"
Const MsgoColAy_Empty = "Lno(?) has a value of no-element-ay of a column-which-is-an-array"
Const MsgoColFldLikAy_NotInFny$ = "Lno(?) has FldLik[?] not in Fny[?]"
Const MsgoColNum_NotBet$ = "Lno(?) has ?[?] not between [?] and [?]"
Function MoCol_Dup(Lnoss$, Valn$, Dup):                           MoCol_Dup = FmtQQ(MsgoCol_Dup, Lnoss, Valn, Dup):                      End Function
Function MoCol_NotIn(L&, V$, Valn$, VdtValss$):                 MoCol_NotIn = FmtQQ(MsgoCol_NotIn, LnoStr(L), Valn, V, Valn, VdtValss):  End Function
Function MoCol_NotNum$(L&, Valn$, V$):                         MoCol_NotNum = FmtQQ(MsgoCol_NotNum, LnoStr(L), Valn, V):                 End Function
Function MoColx_Blnk$(L&, Valn$):                               MoColx_Blnk = FmtQQ(MsgoColx_Blnk, LnoStr(L), Valn):                     End Function
Function MoColAy_Empty(L&):                                   MoColAy_Empty = FmtQQ(MsgoColAy_Empty, LnoStr(L)):                         End Function
Function MoColFldLikAy_NotInFny$(L&, F, FF$):        MoColFldLikAy_NotInFny = FmtQQ(MsgoColFldLikAy_NotInFny, LnoStr(L), F, FF):         End Function
Function MoColNum_NotBet(L&, Valn$, NumV, FmV, ToV):        MoColNum_NotBet = FmtQQ(MsgoColNum_NotBet, LnoStr(L), Valn, NumV, FmV, ToV): End Function

Function EoColF_3Er(Wi_L_Colx As Drs, Fny$()) As String()
EoColF_3Er = EoColx_3Er(Wi_L_Colx, "F", Fny)
End Function

Function EoColx_3Er(Wi_L_Colx As Drs, ColxNm$, Vy$()) As String()
Dim D As Drs: D = Wi_L_Colx
Dim A$(), b$(), C$(), VV$
VV = JnSpc(Vy)
A = EoColx_NotIn(D, "F", "Fld", VV)
b = EoColx_Dup(D, "F", "Fld")
C = EoColx_Blnk(D, ColxNm)
EoColx_3Er = AddSy(A, b)
End Function

Function EoColx_Blnk(Wi_L_Colx As Drs, ColxNm$, Optional Valn0$) As String()
Dim Valn$: Valn = IIf(Valn0 = "", ColxNm, Valn0)
Dim IxL%: IxL = IxzAy(Wi_L_Colx.Fny, ColxNm)
Dim Dr: For Each Dr In Itr(Wi_L_Colx.Dy)
    If IsBlnk(Dr(IxL)) Then
        Dim L&: L = Dr(IxL)
        PushI EoColx_Blnk, MoColx_Blnk(L, Valn)
    End If
Next
End Function

Function EoColFldLikAy_3Er(Wi_L_LikAy As Drs, Fny$()) As String()
Dim D As Drs: D = Wi_L_LikAy
EoColFldLikAy_3Er = SyzAp( _
    EoColAy_Empty(D, "FldLikAy"), _
    EoColx_Dup(D, "FldLikAy", "FldLik"), _
    EoColFldLikAy_NotInFny(D, Fny))
End Function

Function EoColAy_Empty(Wi_L_Ay As Drs, ColAyNm$) As String()
Dim IxL%, IxFny%: AsgIx Wi_L_Ay, "L Fny", IxL, IxFny
Dim Dr: For Each Dr In Itr(Wi_L_Ay.Dy)
    Dim Fny$(): Fny = Dr(IxFny)
    If Si(Fny) = 0 Then
        Dim L&: L = Dr(IxL)
        PushI EoColAy_Empty, MoColAy_Empty(L)
    End If
Next
End Function

Function EoColx_NotIn(Wi_L_Colx As Drs, ColxNm$, Valn$, VdtValss$) As String()
Dim IxL%, IxColx%: AsgIx Wi_L_Colx, "L " & ColxNm, IxL, IxColx
Dim VdtVy$(): VdtVy = SyzSS(VdtValss)
Dim Dr: For Each Dr In Itr(Wi_L_Colx.Dy)
    Dim V$: V = Dr(IxColx)
    Dim L&: L = Dr(IxL)
    If Not HasEle(VdtVy, V) Then
        PushI EoColx_NotIn, MoCol_NotIn(L, V, Valn, VdtValss)
    End If
Next
End Function

Function EoColFldLikAy_NotInFny(Wi_L_LikAy As Drs, InFny$()) As String()
Dim IxFny%, IxL%: AsgIx Wi_L_LikAy, "L Fny", IxL, IxFny
Dim FF$: FF = JnSpc(InFny)
Dim Dr: For Each Dr In Itr(Wi_L_LikAy.Dy)
    Dim Fny$(): Fny = Dr(IxFny)
    Dim F: For Each F In Fny
        If Not HasEle(InFny, F) Then
            Dim L&: L = Dr(IxL)
            PushI EoColFldLikAy_NotInFny, MoColFldLikAy_NotInFny(L, F, FF)
        End If
    Next
Next
End Function

Function EoColx_Dup(Wi_L_Colx As Drs, ColxNm$, Optional Valn0$) As String()
Dim Valn$: Valn = DftStr(Valn0, ColxNm)
Dim Colx():      Colx = Col(Wi_L_Colx, ColxNm)
Dim LnoCol&(): LnoCol = LngCol(Wi_L_Colx, "L")
Dim AllLik$():          'AllLik = CvSy(AyzAyOfAy(FldLikAyCol))
Dim DupAy$():            DupAy = AwDup(AllLik)
Dim DupLik: For Each DupLik In Itr(DupAy)
    Dim Lnoss$: 'Lnoss = Lnoss_FmLnoCol_WhSyCol_HasS(LnoCol, FldLikAyCol, DupLik)
    PushI EoColx_Dup, MoCol_Dup(Lnoss, Valn, DupLik)
Next
If Si(EoColx_Dup) > 0 Then
    Dmp EoColx_Dup
    Stop
End If
End Function

Function EoColx_Dup1(Wi_L_Colx As Drs, ColxNm$, Optional Valn0$) As String()
'@Valn :Nm #Val-Nm-ToBe-Shw-InMsg#
Dim Valn$: Valn = DftStr(ColxNm, Valn0)
Dim U%: U = UB(Wi_L_Colx.Dy)
Dim F$():           F = Wi_L_Colx.Fny
Dim Sy$():         Sy = StrCol(Wi_L_Colx, ColxNm)
Dim LnoCol&(): LnoCol = LngCol(Wi_L_Colx, "L")
Dim DupAy$():   DupAy = AwDup(Sy)
Dim Dup: For Each Dup In Itr(DupAy)
    Dim Lnoss$: Lnoss = Lnoss_FmLnoCol_WhStrCol_HasS(LnoCol, Sy, Dup)
    PushI EoColx_Dup1, MoCol_Dup(Lnoss, Valn, Dup) '<==
Next
End Function

Function EoColx_NumNotBet(Wi_L_Colx As Drs, NumColxNm$, FmV, ToV) As String()
Dim IxNum%, IxL%: AsgIx Wi_L_Colx, JnSpcAp(NumColxNm, "L"), IxNum, IxL
Dim Dr: For Each Dr In Itr(Wi_L_Colx.Dy)
    Dim Num: Num = Val(Dr(IxNum))
    If Not IsBet(Num, FmV, ToV) Then
        Dim L&: L = Dr(IxL)
        PushI EoColx_NumNotBet, MoColNum_NotBet(L, NumColxNm, Num, FmV, ToV)
    End If
Next
End Function

Function EoColx_NotNum(Wi_L_Colx As Drs, ColxNm$) As String()
Dim IxL%, IxColxNm%: AsgIx Wi_L_Colx, "L " & ColxNm, IxL, IxColxNm
Dim Dr: For Each Dr In Wi_L_Colx.Dy
    Dim V$: V = Dr(IxColxNm)
    Dim L&
    If Not IsNumeric(V) Then
        L = Dr(IxL)
        PushI EoColx_NotNum, MoCol_NotNum(L, ColxNm, V)
    End If
Next
End Function
