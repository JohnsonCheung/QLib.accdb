Attribute VB_Name = "MxLpm"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxLpm."
Function Lpm(PmStr$, LpmSpec$) As Dictionary
Dim O As New Dictionary
Set Lpm.DicK_ToSy = O
Dim Ay$(): Ay = SyzSS(PmStr)
    Dim I, S$
    Dim CurPmNm$
    For Each I In Itr(Ay)
        S = I
        If FstChr(S) = "-" Then
            CurPmNm = RmvFstChr(S)
            LpmPushPmNm O, CurPmNm
        Else
            LpmPushPm O, CurPmNm, S
        End If
    Next
ThwIf_LpmEr O, LpmSpec
End Function
Sub ThwIf_LpmEr(Lpm As Dictionary, LpmSpec$)

End Sub
Sub LpmPushPmNm(Lpm As Dictionary, PmNm$)
If Lpm.Exists(PmNm) Then Exit Sub
Lpm.Add PmNm, Sy()
End Sub

Function WhNmzLpm(Lpm As Dictionary) As WhNm
'WhNmzLpm = WhNm(LpmPatn(Lpm, NmPfx), LpmLikAy(Lpm, NmPfx), LpmExlLikAy(Lpm, NmPfx))
End Function

Function LpmHasSw(Lpm As Dictionary, SwNm) As Boolean
LpmHasSw = HasEle(LpmSwNy(Lpm), SwNm)
End Function
Function LpmSwNy(Lpm As Dictionary) As String()
Dim K, V
For Each K In Lpm.Keys
    V = Lpm(K)
    If IsSy(Lpm(K)) Then
        PushI LpmSwNy, K
    End If
Next
End Function

Sub LpmPushPm(Lpm As Dictionary, Nm$, Optional V$)
Dim J%, S$()
If Lpm.Exists(Nm) Then
    If V = "" Then Exit Sub
    S = Lpm(Nm)
    PushI S, V
    Lpm(Nm) = S
Else
    Lpm.Add Nm, Sy(V)
End If
End Sub
Sub DmpLpm(Lpm As Dictionary)
D FmtLpm(Lpm)
End Sub
Function FmtLpm(Lpm As Dictionary) As String()
Dim PmNm, O$()
For Each PmNm In Lpm.Keys
    PushI O, FmtzNmSy(PmNm, CvSy(Lpm(PmNm)))
Next
FmtLpm = AlignLyzSepss(O, "ValCnt Val(")
End Function

Function FmtzNmSy$(PmNm, Sy$())
Select Case Si(Sy)
Case 0:    FmtzNmSy = FmtQQ("PmSw(?)", PmNm)
Case Else: FmtzNmSy = FmtQQ("Pm(?) ValCnt(?) Val(?)", PmNm, Si(Sy), JnSpc(Sy))
End Select
End Function

Function LpmLikAy(Lpm As Dictionary, NmPfx) As String()
LpmLikAy = SyzLpm(Lpm, NmPfx & "LikAy")
End Function

Function LpmExlLikAy(Lpm As Dictionary, NmPfx) As String()
LpmExlLikAy = SyzLpm(Lpm, NmPfx & "ExlLikAy")
End Function

Function SyzLpm(Lpm As Dictionary, PmNm) As String()
'If Lpm.Exists(NmPfx & PmNm) Then
'    SyzLpm = Lpm(PmNm)
'End If
End Function
Function StrVzLpm$(Lpm As Dictionary, PmNm)
Dim Vy$()
    Vy = SyzLpm(Lpm, PmNm)
Select Case Si(Vy)
Case 0
Case 1: StrVzLpm = Vy(0)
'Case Else: Thw CSub, FmtQQ("Parameter [-?] should have one value", PmNm), "Pm PmValSz VzPm-Sy", Fmt, Si(Vy), Vy
End Select
End Function
Function LpmPatn$(Lpm As Dictionary, NmPfx)
LpmPatn = StrVzLpm(Lpm, NmPfx & "Patn")
End Function
