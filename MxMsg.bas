Attribute VB_Name = "MxMsg"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxMsg."
#Const Doc = False
#If Doc Then
ksdf
sdfs sdf
#End If
Function AddNNAv(Nav(), NN$, Av()) As Variant()
Dim O(): O = Nav
If Si(O) = 0 Then
    PushI O, NN
Else
    O(0) = O(0) & " " & NN
End If
PushAy O, Av
AddNNAv = O
End Function

Function AddNmV(Nav(), Nm$, V) As Variant()
AddNmV = AddNNAv(Nav, Nm, Av(V))
End Function

Sub DmpNNAp(NN$, ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
D LyzNyAv(Ny(NN), Av)
End Sub

Function LyzNyAv(Ny$(), Av(), Optional Sep$ = ": ") As String()
Dim J%, O$(), N$()
ResiMax Ny, Av
N = AlignAy(Ny)
For J = 0 To UB(Ny)
    PushIAy LyzNyAv, LyzNv(N(J), Av(J), Sep)
Next
End Function


Function LinzNyAv$(Ny$(), Av())
Dim J%, U1%, U2%, N$, V$, O$()
U1 = UB(Ny)
U2 = UB(Av)
For J = 0 To Max(U1, U2)
    If J <= U1 Then N = QteSq(Ny(J)) Else N = "[?]"
    If J <= U2 Then V = Cell(Av(J)) Else V = "?"
    PushI O, N & " " & V
Next
LinzNyAv = JnVbarSpc(O)
End Function


Function SclzNyAv$(Ny$(), Av())
SclzNyAv = JnSemi(LyzNyAv(Ny, Av))
End Function

Function LyzFunMsg(Fun$, Msg$) As String()
Dim O$(), MsgL1$, MsgRst$
AsgBrk1Dot Msg, MsgL1, MsgRst
PushI LyzFunMsg, EnsSfxDot(MsgL1) & IIf(Fun = "", "", "  @" & Fun)
PushIAy LyzFunMsg, IndentSy(WrpLy(SplitCrLf(MsgRst)))
End Function


Function LyzFunMsgNap(Fun$, Msg$, ParamArray Nap()) As String()
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
If Fun = "" And Msg = "" And Si(Nav) = 0 Then Exit Function
LyzFunMsgNap = LyzFunMsgNav(Fun, Msg, Nav)
End Function

Function LyzFunMsgObjPP(Fun$, Msg$, Obj As Object, PP$) As String()
LyzFunMsgObjPP = AddAy(LyzFunMsg(Fun, Msg), LyzObjPP(Obj, PP))
End Function

Function LyzFunMsgNyAv(Fun$, Msg$, Ny$(), Av()) As String()
LyzFunMsgNyAv = AddAy(LyzFunMsg(Fun, Msg), IndentSy(LyzNyAv(Ny, Av)))
End Function

Function LyzFunMsgNNAv(Fun$, Msg$, NN$, Av()) As String()
LyzFunMsgNNAv = LyzFunMsgNyAv(Fun, Msg, SyzSS(NN), Av)
End Function

Function LyzNv(Nm$, V, Optional Sep$ = ": ") As String()
Dim Ly$(): Ly = FmtV(V)
Dim J%, S$
If Si(Ly) = 0 Then
    PushI LyzNv, Nm & Sep
Else
    PushI LyzNv, Nm & Sep & Ly(0)
End If
S = Space(Len(Nm) + Len(Sep))
For J = 1 To UB(Ly)
    PushI LyzNv, S & Ly(J)
Next
End Function

Function LinzNv$(Nm$, V)
LinzNv = Nm & "=[" & Cell(V) & "]"
End Function

Function LyzMsgNap(Msg$, ParamArray Nap()) As String()
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
LyzMsgNap = LyzMsgNav(Msg, Nav)
End Function

Function LyzNmDrs(Nm$, A As Drs, Optional MaxColWdt% = 100) As String()
LyzNmDrs = LyzNmLy(Nm, FmtDrszNoRdu(A, MaxColWdt), EiNoIx)
End Function

Function LyzNmLy(Nm$, Ly$(), Optional B As EmIxCol = EiBeg1) As String()
Dim L$(), J&, S$
If Si(Ly) = 0 Then
    PushI LyzNmLy, Nm & "(No Lin)"
    Exit Function
End If
L = AddIxPfx(Ly, B)
'Brw L:Stop
S = Space(Len(Nm))
PushI LyzNmLy, Nm & L(0)
For J = 1 To UB(L)
    PushI LyzNmLy, S & L(J)
Next
End Function

Function LyzMsg(Msg$) As String()
LyzMsg = LyzFunMsg("", Msg)
End Function

Function LyzNNAp(NN$, ParamArray Ap()) As String()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
LyzNNAp = LyzNNAv(NN, Av)
End Function



Function LyzNNAv(NN$, Av()) As String()
LyzNNAv = LyzNyAv(Ny(NN), Av)
End Function


Function LinzFunMsg$(Fun$, Msg$)
Dim F$: F = IIf(Fun = "", "", " | @" & Fun)
Dim A$: A = Msg & F
If Cfg.Inf.ShwTim Then
    LinzFunMsg = NowStr & " | " & A
Else
    LinzFunMsg = A
End If
End Function
