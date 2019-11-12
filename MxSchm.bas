Attribute VB_Name = "MxSchm"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxSchm."
Private Type O
    Pkq() As String
    Skq() As String
    DiT As Dictionary
    DiTF As Dictionary
    Td() As DAO.TableDef
End Type

Sub CrtSchm(D As Database, Schm$())
Dim A As SchmDta: A = SchmDta(Schm)
Dim Er$(): Er = ErzSchmDta(A): If Si(Er) > 0 Then Thw CSub, "There is error in Schm", "Er Schm", Er, Schm
With Fnd_O(A)
    AppTdAy D, .Td
    RunSqy D, .Skq
    RunSqy D, .Pkq
    SetTblDeszDic D, .DiT
    SetFldDeszDic D, .DiTF
End With
End Sub

Private Function Fnd_O(A As SchmDta) As O
With Fnd_O
    .Skq = SkSqy(A.Sk_T, A.Sk_Fny)
    .Pkq = PkSqy(PkTny(A.T_T, A.T_Fny))
    Set .DiT = DiczSyab(A.DT_T, A.DT_D)
    Set .DiTF = DiTFqDes(A.DF_F, A.DF_D, A.DTF_T, A.DTF_F, A.DTF_D)
    Dim D As Dictionary: Set D = DiFqEleStr(A.T_Fny, A.EF_E, A.EF_FldLikAy, A.E_E, A.E_EleStr)
    .Td = TdAy(Tny:=A.T_T, FnyAy:=A.T_Fny, DiFqEleStr:=D)
End With
End Function
Private Function DiTFqDes(DF_F$(), DF_D$(), DTF_T$(), DTF_F$(), DTF_D$()) As Dictionary

End Function

Private Function PkTny(T_T$(), T_FnyAy()) As String()
Dim J%: For J = 0 To UB(T_T)
    Dim Fny$(): Fny = T_FnyAy(J)
    If T_T(J) & "Id" = Fny(0) Then PushI PkTny, T_T(J)
Next
End Function

Function IxzLikssAy%(Itm, LikssAy$())
Dim I%, Likss: For Each Likss In LikssAy
    If ItmInLikAy(Itm, SyzSS(Likss)) Then IxzLikssAy = I: Exit Function
    I = I + 1
Next
IxzLikssAy = -1
End Function

Function ItmInLikAy(Itm, LikAy$()) As Boolean
Dim Lik: For Each Lik In LikAy
    If Itm Like Lik Then ItmInLikAy = True: Exit Function
Next
End Function

Private Function DiFqE(AllFny$(), EF_E$(), EF_FldLikAy()) As Dictionary
Set DiFqE = New Dictionary
Dim F: For Each F In AllFny
    Stop
    Dim Ix%: 'Ix = IxzLikssAy(F, EF_FldLikAy)
    Dim E$: E = EF_E(Ix)
    DiFqE.Add F, E
Next
End Function

Private Function DiFqEleStr(T_Fny(), EF_E$(), EF_FldLikAy(), E_E$(), E_EleStr$()) As Dictionary
Dim AllFny$(): AllFny = AwDist(AyzAyOfAy(T_Fny))
Dim FqE As Dictionary: Set FqE = DiFqE(AllFny, EF_E, EF_FldLikAy)
Dim EqEs As Dictionary: Set EqEs = DiczAyab(E_E, E_EleStr)
Dim FqEs As Dictionary: Set FqEs = ChnDic(FqE, EqEs)
Set DiFqEleStr = DiwAy(FqEs, AllFny)
End Function

Private Function TdAy(Tny$(), FnyAy(), DiFqEleStr As Dictionary) As DAO.TableDef()
Dim F() As DAO.Field
Dim T, J%: For Each T In Itr(Tny)
    Dim Fny$(): Fny = FnyAy(J)
    F = FdAyzFny(Fny, DiFqEleStr)
        PushObj TdAy, TdzTFdAy(T, F)
    J = J + 1
Next
End Function

Private Function FdAyzFny(Fny$(), DiFqEleStr As Dictionary) As DAO.Field()
Dim F: For Each F In Fny
    PushObj FdAyzFny, FdzEleStr(F, DiFqEleStr(F))
Next
End Function

Private Sub Z_CrtSchm()
Dim D As Database, Schm$()
GoSub T1
Exit Sub

T1:
    Set D = TmpDb
    Schm = SampSchm1
    GoTo Tst
Tst:
    CrtSchm D, Schm
    Return
End Sub

