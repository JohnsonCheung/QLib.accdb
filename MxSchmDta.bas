Attribute VB_Name = "MxSchmDta"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxSchmDta."
Type SchmDta
    T_L() As Integer
    T_T() As String
    T_Fny() As Variant
    
    EF_L() As Integer
    EF_E() As String
    EF_FldLikAy() As Variant ' Each ele is FldLikAy
    
    E_L() As Integer
    E_E() As String
    E_EleStr() As String
    
    DT_L() As Integer ' ! Each Ele is Lno
    DT_T() As String ' ! Each Ele is T Des
    DT_D() As String
    
    DF_L() As Integer
    DF_F() As String
    DF_D() As String
    
    DTF_L() As Integer  'L#
    DTF_TF() As String  'Tbl.Fld
    DTF_T() As String   'Tbl
    DTF_F() As String   'Fld
    DTF_D() As String   'Des
    
    Sk_L() As Integer   'L#
    Sk_T() As String    'SkTbl
    Sk_Fny() As Variant 'SkFny
End Type
Private Const C_T$ = "Tbl"
Private Const C_EF$ = "EleF"
Private Const C_E$ = "Ele"
Private Const C_DT$ = "Des.Tbl"
Private Const C_DF$ = "Des.Fld"
Private Const C_DTF$ = "Des.TblF"
Private Const C_Sk$ = "Sk"

Private Function DoL2c(A As TDoLTD, T1$, SplitDtaCol_ToCC$) As Drs
'Ret : Drs-L-<@SplitDtaCol_ToCC>
Dim D As Drs: D = A.D
Dim D1 As Drs: D1 = DwEQExl(D, "T1", T1)
Dim Dr, Dy(): For Each Dr In Itr(D1.Dy)
    Dim Dta$: Dta = Dr(1)
    PushI Dy, Array(Dr(0), T1zS(Dta), RmvT1(Dta))
Next
DoL2c = DrszFF("L " & SplitDtaCol_ToCC$, Dy)
End Function

Sub Unzip(Dy2(), OAy1, OAy2)
Erase OAy1, OAy2
Dim U&: U = UB(Dy2): If U = -1 Then Exit Sub
ReDim OAy1(U)
ReDim OAy2(U)
Dim Dr, J&:: For Each Dr In Dy2
    OAy1(J) = Dr(0)
    OAy2(J) = Dr(1)
Next
End Sub

Sub Unzip3(Dy3(), OAy1, OAy2, OAy3)
Erase OAy1, OAy2, OAy3
Dim U&: U = UB(Dy3): If U = -1 Then Exit Sub
ReDim OAy1(U)
ReDim OAy2(U)
ReDim OAy3(U)
Dim Dr, J&: For Each Dr In Dy3
    OAy1(J) = Dr(0)
    OAy2(J) = Dr(1)
    OAy3(J) = Dr(2)
    J = J + 1
Next
End Sub

Sub Unzip4(Dy4(), OAy1, OAy2, OAy3, OAy4)
Erase OAy1, OAy2, OAy3, OAy4
Dim U&: U = UB(Dy4): If U = -1 Then Exit Sub
ReDim OAy1(U)
ReDim OAy2(U)
ReDim OAy3(U)
ReDim OAy4(U)
Dim Dr, J&: For Each Dr In Dy4
    OAy1(J) = Dr(0)
    OAy2(J) = Dr(1)
    OAy3(J) = Dr(2)
    OAy4(J) = Dr(3)
Next
End Sub
Function SchmDta(Schm$()) As SchmDta
Dim D As TDoLTD:      D = TDoLTDzInd(Schm)
Dim DoDF As Drs:   DoDF = DoL2c(D, C_DF, "F Des")
Dim DoDT As Drs:   DoDT = DoL2c(D, C_DT, "T Des")
Dim DoDTF As Drs: DoDTF = DoL2c(D, C_DTF, "TblF Des")
Dim DoT As Drs:     DoT = DoL2c(D, C_T, "T FF")
Dim DoEF As Drs:   DoEF = DoL2c(D, C_EF, "E LikFF")
Dim DoE As Drs:     DoE = DoL2c(D, C_E, "E EleStr")
Dim DoSk As Drs:   DoSk = DoL2c(D, C_Sk, "E Skff")
'Insp CSub, "AA", "LTD T EF E DT DF DTF Sk", _
    FmtDrs(D.D), _
    FmtDrs(DoT), _
    FmtDrs(DoEF), _
    FmtDrs(DoE), _
    FmtDrs(DoDT), _
    FmtDrs(DoDF), _
    FmtDrs(DoDTF), _
    FmtDrs(DoSk)

With SchmDta
    Dim FssAy$()
    Dim EF_FldLikss$()
        Unzip3 DoT.Dy, .T_L, .T_T, FssAy:    .T_Fny = T_Fny(.T_T, FssAy)
       Unzip3 DoDF.Dy, .DF_L, .DF_F, .DF_D
       Unzip3 DoDT.Dy, .DT_L, .DT_T, .DT_D
      Unzip3 DoDTF.Dy, .DTF_L, .DTF_TF, .DTF_D: AsgBrk1DotAy .DTF_TF, .DTF_T, .DTF_F
        Unzip3 DoE.Dy, .E_L, .E_E, .E_EleStr
       Unzip3 DoEF.Dy, .EF_L, .EF_E, EF_FldLikss:  .EF_FldLikAy = AmSyzzSS(EF_FldLikss)
       Unzip3 DoSk.Dy, .Sk_L, .Sk_T, FssAy:        .Sk_Fny = AmSyzzSS(FssAy)
End With
End Function

Sub AsgBrk1DotAy(DotAy$(), OBefDot$(), OAftDot$())
Dim U&: U = UB(DotAy)
ReDim OBefDot(U)
ReDim OAftDot(U)
Dim DotNm, J&: For Each DotNm In Itr(DotAy)
    With Brk1Dot(DotNm)
        OBefDot(J) = .S1
        OAftDot(J) = .S2
    End With
    J = J + 1
Next
End Sub

Private Function T_Fny(Tny$(), FssAy$()) As Variant()
Dim Fss, J%: For Each Fss In Itr(FssAy)
    PushI T_Fny, SyzSS(Replace(Fss, "*", Tny(J)))
Next
End Function

Sub Z_SchmDta()
Dim Act As SchmDta, Schm$()
GoSub T1
Exit Sub
T1:
    Act = SchmDta(SampSchm1)
    Stop
    Return
End Sub

