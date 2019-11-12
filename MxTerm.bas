Attribute VB_Name = "MxTerm"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxTerm."

Sub AsgTTRst(Lin, OT1, OT2, ORst$)
AsgAp T2Rst(Lin), OT1, OT2, ORst
End Sub

Sub Asg3TRst(Lin, OT1, OT2, OT3, ORst$)
AsgAp T3Rst(Lin), OT1, OT2, OT3, ORst
End Sub

Sub Asg4T(Lin, O1$, O2$, O3$, O4$)
AsgAp Fst4Term(Lin), O1, O2, O3, O4
End Sub

Sub Asg4TRst(Lin, O1$, O2$, O3$, O4$, ORst$)
AsgAp T4Rst(Lin), O1, O2, O3, O4, ORst
End Sub

Sub AsgTRst(Lin, OT1, ORst)
AsgAp SyzTRst(Lin), OT1, ORst
End Sub

Sub AsgTT(Lin, O1, O2)
AsgAp T2Rst(Lin), O1, O2
End Sub

Sub AsgAmT1RstAy(Ly$(), OAmT1$(), ORstAy$())
Erase OAmT1, ORstAy
Dim L: For Each L In Itr(Ly)
    PushI OAmT1, T1(L)
    PushI ORstAy, RmvT1(L)
Next
End Sub

Sub AsgT1Sy(LinOf_T1_SS, OT1, Osy$())
Dim Rst$
AsgTRst LinOf_T1_SS, OT1, Rst
Osy = SyzSS(Rst)
End Sub

Function Fst2Term(Lin) As String()
Fst2Term = FstNTerm(Lin, 2)
End Function

Function Fst3Term(Lin) As String()
Fst3Term = FstNTerm(Lin, 3)
End Function

Function Fst4Term(Lin) As String()
Fst4Term = FstNTerm(Lin, 4)
End Function

Function FstNTerm(Lin, N%) As String()
Dim J%, L$
L = Lin
For J = 1 To N
    PushI FstNTerm, ShfT1(L)
Next
End Function


Function SyzTRst(Lin) As String()
SyzTRst = NTermRst(Lin, 1)
End Function

Function T2Rst(Lin) As String()
T2Rst = NTermRst(Lin, 2)
End Function

Function T3Rst(Lin) As String()
T3Rst = NTermRst(Lin, 3)
End Function

Function T4Rst(Lin) As String()
T4Rst = NTermRst(Lin, 4)
End Function

Function NTermRst(Lin, N%) As String()
':NTermRst: :Sy ! With @N+1 ele from @L
Dim L$, J%
L = Lin
For J = 1 To N
    PushI NTermRst, ShfT1(L)
Next
PushI NTermRst, L
End Function

Sub Z_NTermRst()
Dim Lin
Lin = "  [ksldfj ]":  Ept = "ksldfj ": GoSub Tst
Lin = "  [ ksldfj ]": Ept = " ksldf ": GoSub Tst
Lin = "  [ksldfj]":  Ept = "ksldf": GoSub Tst
Exit Sub
Tst:
    Act = T1(Lin)
    C
    Return
End Sub

Function SrcT1AsetP() As Aset
Set SrcT1AsetP = T1Aset(SrczP(CPj))
End Function

Function T1Aset(Ly$()) As Aset
Dim O As New Aset, L
For Each L In Itr(Ly)
    O.PushItm T1(L)
Next
Set T1Aset = O
End Function
Function T1zS$(S)
T1zS = T1(S)
End Function

Function T1$(S)
T1 = ShfT1(CStr(S))
End Function

Function T2zS$(S)
T2zS = T2(S)
End Function

Function T2$(S)
T2 = TermN(S, 2)
End Function

Function T3$(S)
T3 = TermN(S, 3)
End Function

Function TermN$(S, N%)
Dim L$, J%
L = LTrim(S)
For J = 1 To N - 1
    L = RmvT1(L)
Next
TermN = T1(L)
End Function

Sub Z_TermN()
Dim N%, A$
N = 1: A = "a b c": Ept = "a": GoSub Tst
N = 2: A = "a b c": Ept = "b": GoSub Tst
N = 3: A = "a b c": Ept = "c": GoSub Tst
Exit Sub
Tst:
    Act = TermN(A, N)
    C
    Return
End Sub

':Term: :S ! No-spc-str or Sq-quoted-str
':Termss: :SS
':NN: :SS ! spc-sep-str of :Nm
Function RmvT1XzA$(Lin, TermAy$())
Dim T$: T = T1(Lin)
If HasEle(TermAy, T) Then
    RmvT1XzA = RmvT1x(Lin, T)
Else
    RmvT1XzA = Lin
End If
End Function

Function RmvT1x$(Lin, T1x$)
Dim T$: T = T1(Lin)
If T = T1x Then
    RmvT1x = RmvT1(Lin)
Else
    RmvT1x = LTrim(Lin)
End If
End Function

Function RplT1$(L, T1$, By$)
If HasT1(L, T1) Then
    RplT1 = By & Mid(L, Len(T1) + 1)
Else
    RplT1 = L
End If
End Function

Function Termss(TermAy)
Termss = TLin(TermAy)
End Function

Function TLin(TermAy)
TLin = JnTerm(TermAy)
End Function

Function TLinzAp$(ParamArray TermAp())
Dim Av(): Av = TermAp
TLinzAp = JnTerm(Av)
End Function

Function JnTerm$(TermAy)
JnTerm = JnSpc(SyzQteSqIf(RmvBlnkzAy(TermAy)))
End Function

Function CommaTerm$(Lin, Ix)
Dim Ay$(): Ay = Split(Lin, ",")
If Not IsBet(Ix, 0, UB(Ay)) Then Exit Function
CommaTerm = Ay(Ix)
End Function

Function LinzTermAy$(TermAy)
LinzTermAy = TLin(TermAy)
End Function

Function TermAset(S) As Aset
Set TermAset = AsetzAy(TermAy(S))
End Function

Function TermItr(S)
Asg Itr(TermAy(S)), TermItr
End Function

Function TermLin$(Dr())
Dim O$()
Dim V: For Each V In Itr(Dr)
    If HasSpc(V) Then
        PushI O, QteSq(V)
    Else
        PushI O, V
    End If
Next
TermLin = JnSpc(O)
End Function
Function TermAy(TermLin) As String()
Dim L$, J%
L = TermLin
While L <> ""
    J = J + 1: If J > 5000 Then Stop
    PushNB TermAy, ShfT1(L)
Wend
End Function

Function ShfTerm(OLin$, X$) As Boolean
If T1(OLin) = X Then
    ShfTerm = True
    OLin = RmvT1(OLin)
End If
End Function

Function ShfT1$(OLin)
Dim O$: O = LTrim(OLin)
If FstChr(O) = "[" Then
    Dim P%
    P = InStr(O, "]")
    If P = 0 Then
        Thw CSub, "OLin has fstchr [, but no ]", "OLin", OLin
    End If
    ShfT1 = Mid(O, 2, P - 2)
    OLin = LTrim(Mid(O, P - 1))
    Exit Function
End If
P = InStr(O, " ")
If P = 0 Then
    ShfT1 = O
    OLin = ""
Else
    ShfT1 = Left(O, P - 1)
    OLin = LTrim(Mid(O, P + 1))
End If
End Function

Function ShfTermDot$(OLin$)
With Brk2Dot(OLin, NoTrim:=True)
    ShfTermDot = .S1
    OLin = .S2
End With
End Function

Sub Z_ShfT1()
Dim O$, OEpt$
O = " S   DFKDF SLDF  "
OEpt = "DFKDF SLDF  "
Ept = "S"
GoSub Tst
'
O = " AA BB "
Ept = "AA"
OEpt = "BB "
GoSub Tst
'
Exit Sub
Tst:
    Act = ShfT1(O)
    C
    Ass O = OEpt
    Return
End Sub
