Attribute VB_Name = "MxLofDta"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxLofDta."
Type LofDta
    L_Lon As Drs ' L
    L_LoFny As Drs ' L
    Fny() As String
    L_Ali_FldLikAy As Drs
    L_Wdt_FldLikAy As Drs
    L_Bdr_FldLikAy As Drs
    L_Lvl_FldLikAy As Drs
    L_Cor_FldLikAy As Drs
    L_Tot_FldLikAy As Drs
    L_Fmt_FldLikAy As Drs
    L_F_Tit As Drs
    L_F_Fml As Drs
    L_F_Lbl As Drs
    L_Fm_To_Sum As Drs
    DoErLTD As Drs
End Type

Private Sub Z_LofDta()
Dim A As LofDta: A = LofDta(SampLof)
Stop
End Sub

Private Function Do_L_Lon(DoLTTD As Drs) As Drs
'@DoLTTD ! Select T1.T2="Lo Nm" and Set *Lon = *Dta
'Ret :Drs-L-Lon
Dim D As Drs: D = DwCC2EqExl(DoLTTD, "T1 T2", "Lo", "Nm")
Do_L_Lon = DrszFF("L Lon", D.Dy)
End Function

Private Function Do_L_LoFny(DoLTTD As Drs) As Drs
'@DoLTTD ! Select T1="Lo" and FstTerm(*Dta)='Fld' and Set *LoFny = SyzSS(Rst(*Dta))
'Ret :Drs-L-LoFny
Dim D As Drs: D = DwCC2EqExl(DoLTTD, "T1 T2", "Lo", "Fld")
Dim Dy(), Dr: For Each Dr In Itr(D.Dy)
    PushI Dr, SyzSS(Pop(Dr))
    PushI Dy, Dr
Next
Do_L_LoFny = DrszFF("L LoFny", Dy)
End Function

Private Function Do_L_Fm_To_Sum(DoLTTD As Drs) As Drs
'@DoLTTD ! *T1.T2 = 'Sum Bet' *Dta = *Fm *To *Sum
'Ret  :Drs-L-Fm-To-Sum
Dim Dy()
    Dim Dr: For Each Dr In Itr(DwCC2EqExl(DoLTTD, "T1 T2", "Sum", "Bet").Dy)
        Dim L&: L = Dr(0)
        Dim Dta$: Dta = Dr(1)
        Dim FldFm$
        Dim FldTo$
        Dim FldSum$: AsgTTRst Dta, FldFm, FldTo, FldSum
        PushI Dy, Array(L, FldFm, FldTo, FldSum)
    Next
Do_L_Fm_To_Sum = DrszFF("L Fm To Sum", Dy)
End Function

Function Do_L_F_XXX(DoLTTD As Drs, T_Val$) As Drs
Dim D As Drs: D = DwEqExl(DoLTTD, "T1", T_Val)
Dim Dy()
    Dim L&, XXX$, F$, Dta$
    Dim Dr: For Each Dr In Itr(D.Dy)
          L = Dr(0)
          F = Dr(1)
        XXX = Dr(2)
              PushI Dy, Array(L, F, XXX)
    Next
Do_L_F_XXX = DrszFF("L F " & T_Val, Dy)
End Function

Function FnyzFnyAy(FnyAy()) As String()
Dim Fny: Fny = AyzAyOfAy(FnyAy)
FnyzFnyAy = CvSy(AwDist(Fny))
End Function

Sub AA()
Z_LofDta
End Sub

Function LofDta(Lof$()) As LofDta
Dim D As Drs: D = DoLTTD(Lof)
With LofDta
    .L_Lon = Do_L_Lon(D)
    .L_LoFny = Do_L_LoFny(D)
    .Fny = FnyzFnyAy(Col(.L_LoFny, "LoFny"))
    .L_Ali_FldLikAy = Do_L_XXX_FldLikAy(D, "Ali")
    .L_Wdt_FldLikAy = Do_L_XXX_FldLikAy(D, "Wdt")
    .L_Bdr_FldLikAy = Do_L_XXX_FldLikAy(D, "Bdr")
    .L_Lvl_FldLikAy = Do_L_XXX_FldLikAy(D, "Lvl")
    .L_Cor_FldLikAy = Do_L_XXX_FldLikAy(D, "Cor")
    .L_Tot_FldLikAy = Do_L_XXX_FldLikAy(D, "Tot")
    .L_Fmt_FldLikAy = Do_L_XXX_FldLikAy(D, "Fmt")
    .L_F_Tit = Do_L_F_XXX(D, "Tit")
    .L_F_Fml = Do_L_F_XXX(D, "Fml")
    .L_F_Lbl = Do_L_F_XXX(D, "Lbl")
    .L_Fm_To_Sum = Do_L_Fm_To_Sum(D)
    .DoErLTD = DoErLTD(D, VdtLofT1ss)
End With
End Function

Function Do_L_XXX_FldLikAy(DoLTTD As Drs, ColXXXNm$) As Drs
Dim D As Drs: D = DwEqExl(DoLTTD, "T1", ColXXXNm)
Dim Dy()
    Dim L&, XXX$, Fny$()
    Dim Dr: For Each Dr In Itr(D.Dy)
          L = Dr(0)
        XXX = Dr(1)
        Fny = SyzSS(Dr(2))
              PushI Dy, Array(L, XXX, Fny)
    Next
Do_L_XXX_FldLikAy = DrszFF("L " & ColXXXNm & " FldLikAy", Dy)
End Function
