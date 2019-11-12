Attribute VB_Name = "MxThw"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxThw."
Type CfgInf
    ShwInf As Boolean
    ShwTim As Boolean
End Type
Type CfgSql
    FmtSql As Boolean
End Type
Type Cfg
    Inf As CfgInf
    Sql As CfgSql
End Type

Public Property Get Cfg() As Cfg
Static X As Boolean, Y As Cfg
If Not X Then
    X = True
    Y.Sql.FmtSql = True
    Y.Inf.ShwInf = True
    Y.Inf.ShwTim = True
End If
Cfg = Y
End Property

Sub InspzAv(Fun$, Msg$, NN$, Av())

End Sub

Sub Insp(Fun$, Msg$, ParamArray Nap())
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
Dim F$: If Fun <> "" Then F = " (@" & Fun & ")"
Dim A$(): A = BoxzS("Insp: " & Msg & F)
BrwAy Sy(A, LyzNav(Nav))
End Sub

Sub Z_Thw()
Thw "SF", "AF"
End Sub

Sub InfLin(Fun$, Msg$, ParamArray Nap())
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
D LinzFunMsgNav(Fun, Msg, Nav)
End Sub

Sub WarnLin(Fun$, Msg$, ParamArray Nap())
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
Debug.Print LinzFunMsgNav(Fun, Msg, Nav)
End Sub

Sub Warn(Fun$, Msg$, ParamArray Nap())
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
D LyzFunMsgNav(Fun, Msg, Nav)
End Sub
Sub PrpEr(CSub)
Debug.Print "PrpEr: "; CSub; ": "; Err.Description
End Sub
Sub Thw(Fun$, Msg$, ParamArray Nap())
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
Dim A$(): A = BoxzS("Program error")
BrwAy AddSy(LyzFunMsg(Fun, Msg), LyzNav(Nav))
Halt
End Sub

Sub ThwNav(Fun$, Msg$, Nav())
BrwAy LyzFunMsgNav(Fun, Msg, Nav)
Halt
End Sub

Sub Ass(A As Boolean)
Debug.Assert A
End Sub

Sub ThwNever(Fun$, Optional Msg$ = "Program should not reach here")
Thw Fun, Msg
End Sub

Sub Halt(Optional Fun$)
Err.Raise -1, Fun, "Please check messages opened in notepad"
End Sub

Sub Done()
MsgBox "Done"
End Sub

Sub ThwLoopingTooMuch(Fun$)
Thw Fun, "Looping too much"
End Sub

Sub ThwPmEr(VzPm, Fun$, Optional MsgWhyPmEr$ = "Invalid value")
Thw Fun, "Parameter error: " & MsgWhyPmEr, "Pm-Type Pm-Val", TypeName(VzPm), FmtV(VzPm)
End Sub

Sub D(Optional V)
Dim A$(): A = FmtV(V)
DmpAy A
End Sub

Sub Dmp(A)
D A
End Sub

Sub DmpTy(A)
Debug.Print TypeName(A)
End Sub

Sub DmpAyWithIx(Ay)
Dim J&
For J = 0 To UB(Ay)
    Debug.Print J; ": "; Ay(J)
Next
End Sub

Sub DmpAy(Ay)
Dim J&
For J = 0 To UB(Ay)
    Debug.Print Ay(J)
Next
End Sub


Sub StopEr(Er$())
If Si(Er) = 0 Then Exit Sub
BrwAy Er
Stop
End Sub

Sub ThwImpossible(Fun$)
Thw Fun, "Impossible to reach here"
End Sub

Sub Inf(Fun$, Msg$, ParamArray Nap())
If Not Cfg.Inf.ShwInf Then Exit Sub
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
D LyzFunMsgNav(Fun, Msg, Nav)
End Sub


