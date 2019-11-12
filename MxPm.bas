Attribute VB_Name = "MxPm"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxPm."
Function FmtAyPm(Ay, PmStr$) As String() 'PmStr [FF..] [AlignNCol:FF..] ..
Dim T1, D As Dictionary
Set D = T1ToAlignNColDic(PmStr)
For Each T1 In D
    PushIAy FmtAyPm, FmtAyPmzT1(Ay, T1, D(T1))
Next
End Function

Function FmtAyPmzT1(Ay, T1, AlignNCol) As String()
FmtAyPmzT1 = AlignLyzNTerm(AwT1(Ay, T1), CInt(AlignNCol))
End Function

Function T1ToAlignNColDic(PmStr$) As Dictionary
Set T1ToAlignNColDic = New Dictionary
    Dim Ay$(), F, D As Dictionary
    Ay = TermAy(PmStr)
    Set D = T1ToAlignNColDiczNoSrt(Ay)
    For Each F In TermAy(Ay(0))
        If D.Exists(F) Then
            T1ToAlignNColDic.Add F, D(F)
        Else
            T1ToAlignNColDic.Add F, 1
        End If
    Next
End Function

Function T1ToAlignNColDiczNoSrt(PmLy$()) As Dictionary
Dim J%, W%, F
Set T1ToAlignNColDiczNoSrt = New Dictionary
For J = 2 To UB(PmLy)
    With Brk(PmLy(J), ":")
        W = .S1
        For Each F In TermAy(.S2)
            T1ToAlignNColDiczNoSrt.Add F, W
        Next
    End With
Next
End Function
