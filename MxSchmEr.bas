Attribute VB_Name = "MxSchmEr"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxSchmEr."
Const Msg_DF_DesMis$ = "L#({L}) Des.Fld({F}) has no des *DF_DesMis"
Const Msg_DF_FldNotUse$ = "L#({L}) Des.Fld({F}) has not used *DF_FldNotUse"

Const Msg_DT_DesMis$ = "L#({L}) Des.Tb{T}) has no des *DT_DesMsg"
Const Msg_DT_TblNDef$ = "L#({L}) Des.Tbl({T}) is not defined *DT_TblNDef"

Const Msg_DTF_DesMis$ = "L#({L}) Des.TblF{{TF}) has no des *DTF_DesMis"
Const Msg_DTF_FldNDef$ = "L#({L}) Des.TblF Tbl({T}) has Fld({F}) is not defined *DTF_FldNDef"
Const Msg_DTF_TblNDef$ = "L#({L}) Des.TblF.Tbl({T}) is not defined *DTF_EleNoEleStr"

Const Msg_E_EleNoEleStr$ = "L#({L}) Ele({E}) has no EleStr *E_EleNoEleStr"
Const Msg_E_EleNotUse$ = "L#({L}) Ele({E}) is not used *E_EleNotUse"
Const Msg_E_EleStrEr$ = "L#({L}) Ele({E}) has er in EleStr({EleStr}): {Er} *E_EleStrEr"

Const Msg_EF_EleNoFld$ = "L#({L}) Ele({E}) has no Fld *EF_EleNoFld"
Const Msg_EF_EleNotUse$ = "L#({L}) Ele({E}) has all Likf not used.  The line can be deleted. *EF_EleNotUse"
Const Msg_EF_LikfNotUsed$ = "L#({L}) Ele({E}) has LikFld({F}) not used *EF_LikfNotUsed"

Const Msg_Sk_FldNDef$ = "L#({L}) SkTbl({T}) does not has Fld({F}) *Sk_FldNDef"
Const Msg_Sk_FldDup$ = "L#({L}) SkTbl({T}) has dup Fld({F}) *Sk_FldDup"
Const Msg_Sk_TblNDef$ = "L#({L}) SkTbl({T}) is not defined *Sk_TblNDef"
Const Msg_Sk_TblDup$ = "L#({L}) SkTbl({T}) is duplicated *Sk_TblDup"
Const Msg_Sk_TblNoFld$ = "L#({L}) SkTbl({T}) does not has Fld. *Sk_TblNoFld"

Const Msg_T_FldDup$ = "L#({L}) Tbl({T}) HasDupFld({F}) *T_FldDup"
Const Msg_T_FldNoEleNorStd$ = "L#({L}) Tbl({T}) Fld({F}) is not def in EleF nor {StdFldLikss}"
Const Msg_T_FldDupMulLin$ = "L#({L1}) Tbl({T}) HasDupFld({F}) in L#({L1}) *T_FldDupMulLin" 'L# has multiple Lno#.  Tbl may be multiple TblNm, Fld is single field
Const Msg_T_LinMis$ = "No Tbl Lin"

Const FFoLF$ = "L F"
Const FFoLEF$ = "L E F"
Const FFoLT$ = "L T"
Const FFoLTFL$ = "L1 T F L2"
Const FFo_L_TF$ = "L TF"
Const FFoLE$ = "L E"
Const FFoLTF$ = "L T F"
Const FFo_L_E_EleStr_Er = "L E EleStr Er"

Private Function Er_DesMis(L%(), T_or_F_or_TF$(), D$(), FF$, Msg$) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If D(J) = "" Then
        PushI Dy, Array(L(J), T_or_F_or_TF(J))
    End If
Next
Er_DesMis = MsgzDrs(Msg, DrszFF(FF, Dy))
End Function

Private Function Er_DF_DesMis(L%(), F$(), D$()) As String()
Er_DF_DesMis = Er_DesMis(L, F, D, FFoLF, Msg_DF_DesMis)
End Function

Private Function Er_DF_FldNotUse(L%(), F$(), AllFny$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If Not HasEle(AllFny, F(J)) Then
        PushI Dy, Array(L(J), F(J))
    End If
Next
Er_DF_FldNotUse = MsgzDrs(Msg_DF_FldNotUse, DrszFF(FFoLF, Dy))
End Function

Private Function Er_DT_DesMis(L%(), T$(), D$()) As String()
Er_DT_DesMis = Er_DesMis(L, T, D, FFoLT, Msg_DT_DesMis)
End Function

Private Function Er_DT_TblNDef(L%(), T$(), AllTny$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If Not HasEle(AllTny, T(J)) Then
        PushI Dy, Array(L(J), T(J))
    End If
Next
Er_DT_TblNDef = MsgzDrs(Msg_DT_TblNDef, DrszFF(FFoLT, Dy))
End Function

Private Function Er_DTF_DesMis(L%(), TF$(), D$()) As String()
Er_DTF_DesMis = Er_DesMis(L, TF, D, FFo_L_TF, Msg_DTF_DesMis)
End Function

Private Function Er_DTF_FldNDef(L%(), T$(), F$(), AllFny$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If Not HasEle(AllFny, F(J)) Then
        PushI Dy, Array(L(J), T(J), F(J))
    End If
Next
Er_DTF_FldNDef = MsgzDrs(Msg_DTF_FldNDef, DrszFF(FFoLTF, Dy))
End Function

Private Function Er_DTF_TblNDef(L%(), T$(), AllTny$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If Not HasEle(AllTny, T(J)) Then
        PushI Dy, Array(L(J), T(J))
    End If
Next
Er_DTF_TblNDef = MsgzDrs(Msg_DTF_TblNDef, DrszFF(FFoLT, Dy))
End Function

Private Function Er_E_EleNoEleStr(L%(), E$(), EleStr$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If EleStr(J) = "" Then
        PushI Dy, Array(L(J), E(J))
    End If
Next
Er_E_EleNoEleStr = MsgzDrs(Msg_E_EleNoEleStr, DrszFF(FFo_L_E_EleStr_Er, Dy))
End Function

Private Function Er_E_EleNotUse(L%(), E$(), InUseEny$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If Not HasEle(InUseEny, E(J)) Then
        PushI Dy, Array(L(J), E(J))
    End If
Next
Er_E_EleNotUse = MsgzDrs(Msg_E_EleNotUse, DrszFF(FFoLE, Dy))
End Function

Private Function Er_E_EleStrEr(L%(), E$(), EleStr$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    Dim Er$: Er = ErzEleStr(EleStr(J))
    If Er <> "" Then
        PushI Dy, Array(L(J), E(J), EleStr(J), Er)
    End If
Next
Er_E_EleStrEr = MsgzDrs(Msg_E_EleStrEr, DrszFF(FFo_L_E_EleStr_Er, Dy))
End Function

Private Function Er_EF_EleNoFld(L%(), E$(), E_FldLikAy()) As String()
Dim Dy()
Er_EF_EleNoFld = MsgzDrs(Msg_EF_EleNoFld, DrszFF(FFoLT, Dy))
End Function

Function HasLik(Ay, Lik) As Boolean
Dim V: For Each V In Itr(Ay)
    If V Like Lik Then HasLik = True: Exit Function
Next
End Function

Private Function Er_EF_EleNotUse(L%(), E$(), FldLikAy(), AllFny$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    Dim I%: For I = 0 To UB(FldLikAy)
        If HasLik(AllFny, FldLikAy(I)) Then GoTo X
    Next
    PushI Dy, Array(L(J), E(J))
X:
Next
Er_EF_EleNotUse = MsgzDrs(Msg_EF_EleNotUse, DrszFF(FFoLE, Dy))
End Function

Private Function Er_EF_LikfNotUsed(L%(), E$(), FldLikAy(), AllFny$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    Dim LikAy$(): LikAy = FldLikAy(J)
    Dim I%: For I = 0 To UB(LikAy)
        Dim Lik$: Lik = LikAy(J)
        If Not HasLik(AllFny, Lik) Then
            PushI Dy, Array(L(J), E(J), LikAy(I))
        End If
    Next
Next
Er_EF_LikfNotUsed = MsgzDrs(Msg_EF_LikfNotUsed, DrszFF(FFoLEF, Dy))
End Function

Private Function Er_Sk_FldNDef(L%(), T$(), SkFny(), AllFny$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    Dim Fny$(): Fny = SkFny(J)
    Dim F: For Each F In Itr(Fny)
        If Not HasEle(AllFny, F) Then
            PushI Dy, Array(L(J), T(J), F)
        End If
    Next
Next
Er_Sk_FldNDef = MsgzDrs(Msg_Sk_FldNDef, DrszFF(FFoLTF, Dy))
End Function

Private Function Er_FldDup(L%(), T$(), Fny(), Msg$) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    Dim TFny$(): TFny = Fny(J)
    Dim Dup$(): Dup = AwDup(TFny)
    Dim F: For Each F In Itr(Dup)
        PushI Dy, Array(L(J), T(J), F)
    Next
Next
Er_FldDup = MsgzDrs(Msg, DrszFF(FFoLTF, Dy))
End Function

Private Function Er_Sk_FldDup(L%(), T$(), SkFny()) As String()
Er_Sk_FldDup = Er_FldDup(L, T, SkFny, Msg_Sk_FldDup)
End Function

Private Function Er_Sk_TblNDef(L%(), T$(), AllTny$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If Not HasEle(AllTny, T(J)) Then
        PushI Dy, Array(L(J), T(J))
    End If
Next
Er_Sk_TblNDef = MsgzDrs(Msg_Sk_TblNDef, DrszFF(FFoLT, Dy))
End Function

Private Function Er_Sk_TblDup(L%(), T$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    Dim Dup$(): Dup = AwDup(T)
    Dim D: For Each D In Itr(Dup)
        PushI Dy, Array(L(J), D)
    Next
Next
Er_Sk_TblDup = MsgzDrs(Msg_Sk_TblDup, DrszFF(FFoLT, Dy))
End Function

Private Function Er_Sk_TblNoFld(L%(), T$(), SkFny()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    Dim Fny$(): Fny = SkFny(J)
    If Si(Fny) = 0 Then
        PushI Dy, Array(L(J), T(J))
    End If
Next
Er_Sk_TblNoFld = MsgzDrs(Msg_Sk_TblNoFld, DrszFF(FFoLT, Dy))
End Function

Private Function Er_T_FldDup(L%(), T$(), Fny()) As String()
Er_T_FldDup = Er_FldDup(L, T, Fny, Msg_T_FldDup)
End Function

Private Function Er_T_FldDupMulLin(L%(), T$(), Fny()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    Dim TFny$(): TFny = Fny(J)
    Dim F: For Each F In Itr(TFny)
        Dim I%: For I = 0 To UB(L)
            If I <> J Then
                Dim IFny$(): IFny = Fny(I)
                If HasEle(IFny, F) Then
                    PushI Dy, Array(L(J), T(J), F, I)
                End If
            End If
        Next
    Next
Next
Er_T_FldDupMulLin = MsgzDrs(Msg_T_FldDupMulLin, DrszFF(FFoLTFL, Dy))
End Function

Private Function Er_T_FldNoEleNorStd(L%(), T$(), Fny(), AllFnyWithEle$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    Dim IFny$(): IFny = Fny(J)
    Dim F: For Each F In Itr(IFny)
        If Not HasEle(AllFnyWithEle, F) Then
            If Not IsStdFld(F) Then
                PushI Dy, Array(L(J), T(J), F)
            End If
        End If
    Next
Next
Er_T_FldNoEleNorStd = MsgzDrs(Msg_T_FldNoEleNorStd, DrszFF(FFoLTF, Dy))
End Function

Private Function Er_T_LinMis$(T$())
If Si(T) = 0 Then Er_T_LinMis = Msg_T_LinMis
End Function

Function ErlzSchml$(Schml$)
ErlzSchml = JnCrLf(ErzSchmDta(SchmDta(SplitCrLf(Schml))))
End Function

Function IsItmInLikAy(Itm, LikAy) As Boolean
Dim Lik: For Each Lik In Itr(LikAy)
    If Itm Like Lik Then IsItmInLikAy = True: Exit Function
Next
End Function

Private Function Fnd_AllFnyWithEle(AllFny$(), EF_FldLikAy()) As String()
Dim F: For Each F In Itr(AllFny)
    Dim J%: For J = 0 To UB(EF_FldLikAy)
        Dim LikAy$(): LikAy = EF_FldLikAy(J)
        If IsItmInLikAy(F, LikAy) Then
            PushI Fnd_AllFnyWithEle, F
            GoTo Nxt
        End If
    Next
Nxt:
Next
End Function

Private Function Fnd_InUseEny(AllFnyWithEle$(), E_E$(), EF_FldLikAy()) As String()
'T use F, F use E, all E is in use
Dim F: For Each F In Itr(AllFnyWithEle)
    Dim J%: For J = 0 To UB(EF_FldLikAy)
        Dim LikAy$(): LikAy = EF_FldLikAy(J)
        If IsItmInLikAy(F, LikAy) Then
            PushI Fnd_InUseEny, E_E(J)
            GoTo Nxt
        End If
    Next
Nxt:
Next
End Function

Function ErzSchmDta(A As SchmDta) As String()
With A
Dim X_AllFny$(): X_AllFny = AwDist(AyzAyOfAy(.T_Fny))
Dim X_AllFnyWithEle$(): X_AllFnyWithEle = Fnd_AllFnyWithEle(X_AllFny, .EF_FldLikAy) 'AllFny in T has ele
Dim X_InUseEny$(): X_InUseEny = Fnd_InUseEny(X_AllFnyWithEle, .E_E, .EF_FldLikAy)  'T use F, F use E, all E is in use

Dim DF1$():   DF1 = Er_DF_DesMis(.DF_L, .DF_F, .DF_D)
Dim DF2$():   DF2 = Er_DF_FldNotUse(.DF_L, .DF_F, X_AllFny)
Dim DT1$():   DT1 = Er_DT_DesMis(.DT_L, .DT_T, .DT_D)
Dim DT2$():   DT2 = Er_DT_TblNDef(.DT_L, .DT_T, .T_T)
Dim DTF1$(): DTF1 = Er_DTF_DesMis(.DTF_L, .DTF_TF, .DTF_D)
Dim DTF2$(): DTF2 = Er_DTF_FldNDef(.DTF_L, .DTF_T, .DTF_F, X_AllFny)
Dim DTF3$(): DTF3 = Er_DTF_TblNDef(.DTF_L, .DTF_T, .T_T)
Dim E1$():     E1 = Er_E_EleNoEleStr(.E_L, .E_E, .E_EleStr)
Dim E2$():     E2 = Er_E_EleNotUse(.E_L, .E_E, X_InUseEny)
Dim E3$():     E3 = Er_E_EleStrEr(.E_L, .E_E, .E_EleStr)
Dim EF1$():   EF1 = Er_EF_EleNoFld(.EF_L, .EF_E, .EF_FldLikAy)
Dim EF2$():   EF2 = Er_EF_EleNotUse(.EF_L, .EF_E, .EF_FldLikAy, X_AllFny)
Dim EF3$():   EF3 = Er_EF_LikfNotUsed(.EF_L, .EF_E, .EF_FldLikAy, X_AllFny)
Dim Sk1$():   Sk1 = Er_Sk_FldNDef(.Sk_L, .Sk_T, .Sk_Fny, X_AllFny)
Dim Sk2$():   Sk2 = Er_Sk_FldDup(.Sk_L, .Sk_T, .Sk_Fny)
Dim Sk3$():   Sk3 = Er_Sk_TblNDef(.Sk_L, .Sk_T, .T_T)
Dim Sk4$():   Sk4 = Er_Sk_TblDup(.Sk_L, .Sk_T)
Dim Sk5$():   Sk5 = Er_Sk_TblNoFld(.Sk_L, .Sk_T, .Sk_Fny)
Dim T1$():     T1 = Er_T_FldDup(.T_L, .T_T, .T_Fny)
Dim T2$():     T2 = Er_T_FldDupMulLin(.T_L, .T_T, .T_Fny)
Dim T3$():     T3 = Er_T_FldNoEleNorStd(.T_L, .T_T, .T_Fny, X_AllFnyWithEle)
Dim T4$:       T4 = Er_T_LinMis(.T_T)
End With

Dim D$(): D = SyzAp(DT1, DT2, DF1, DF2, DTF2, DTF2, DTF3)
Dim E$(): E = SyzAp(E1, E2, E3)
Dim EF$(): EF = SyzAp(EF1, EF2, EF3)
Dim Sk$(): Sk = SyzAp(Sk1, Sk2, Sk3, Sk4, Sk5)
Dim T$(): T = SyzAp(T1, T2, T3)
ErzSchmDta = SyzAp(D, E, EF, Sk, T)
End Function

