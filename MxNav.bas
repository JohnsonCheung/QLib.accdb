Attribute VB_Name = "MxNav"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxNav."

Function LyzNav(Nav()) As String()
If Si(Nav) = 0 Then Exit Function
Dim Ny$(), Av()
AsgNyAv Nav, Ny, Av
LyzNav = LyzNyAv(Ny, Av)
End Function

Sub AsgNyAv(Nav(), ONy$(), OAv())
If Si(Nav) = 0 Then
    Erase ONy
    Erase OAv
    Exit Sub
End If
Dim TT$: TT = Nav(0)
ONy = TermAy(TT)
OAv = AeFstEle(Nav)
End Sub

Sub Z_LyzNav()
Dim Nav(): Nav = Array("aa bb", 1, 2)
D LyzNav(Nav)
End Sub

Function LinzFunMsgNav$(Fun$, Msg$, Nav())
LinzFunMsgNav = LinzFunMsg(Fun, Msg) & " " & LinzNav(Nav)
End Function

Function LyzFunMsgNav(Fun$, Msg$, Nav()) As String()
Dim A$(): A = LyzFunMsg(Fun, Msg)
Dim B$(): B = IndentSy(LyzNav(Nav))
LyzFunMsgNav = AddAy(A, B)
End Function

Function LyzMsgNav(Msg$, Nav()) As String()
LyzMsgNav = AddAy(LyzMsg(Msg), IndentSy(LyzNav(Nav)))
End Function

Function LinzLyzMsgNav$(Msg$, Nav())
LinzLyzMsgNav = EnsSfxDot(Msg) & " | " & LinzNav(Nav)
End Function

Function LinzNav$(Nav())
Dim Ny$(), Av()
AsgNyAv Nav, Ny, Av
LinzNav = LinzNyAv(Ny, Av)
End Function

