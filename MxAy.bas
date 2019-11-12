Attribute VB_Name = "MxAy"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAy."
':Stmt$ = "Stmt is a string between StmtBrkPatn"
':FunPfx-Am: :Fun-Pfx #Ay-Ele-XXX# ! Given @Ay will return same number of ele after doing some mapping
Public Const StmtBrkPatn$ = "(\.  |\r\n|\r)"
':Sy: :String-Array #String-Array#
':SS: :Lin #Spc-Separated# ! It will be separated in :Sy
':IntSeg: :IntAy #Int-Sequence# ! Each next element is always 1 more than previous one
':LngSeg: :LngAy #Int-Sequence# ! Each next element is always 1 more than previous one
Sub AsgAp(Ay, ParamArray OAp())
Dim J%, OAv()
OAv = OAp
For J = 0 To Min(UB(Ay), UB(OAv))
    OAp(J) = Ay(J)
Next
End Sub

Sub AsgT1AyRstAy(Ly$(), OT1Sy$(), ORestSy$())
OT1Sy = AmT1(Ly)
ORestSy = AmzRmvT1(Ly)
End Sub

Function BrwAy(Ay, Optional Fnn$, Optional OupTy As EmOupTy = EmOupTy.EiOtBrw)
If OupTy = EiOtDmp Then
    Dmp Ay
    Exit Function
End If
Dim T$: T = TmpFt("BrwAy", Fnn)
WrtAy Ay, T
BrwFt T, OupTy = EiOtVc
BrwAy = Ay
End Function

Function ChkDup(Ay, QMsg$) As String()
Dim Dup
Dup = AwDup(Ay)
If Si(Dup) = 0 Then Exit Function
PushI ChkDup, FmtQQ(QMsg, JnSpc(Dup))
End Function

Function LyzVbl(Vbl) As String()
LyzVbl = SplitVBar(Vbl)
End Function
Function DupAmT1(Ly$(), Optional C As VbCompareMethod = vbTextCompare) As String()
Dim A$(): A = AmT1(Ly)
DupAmT1 = AwDup(A, C)
End Function

Function ChkAyEmp(Ay, Msg$) As String()
If Si(Ay) = 0 Then ChkAyEmp = Sy(Msg)
End Function

Function ChkEqAy(Ay1, Ay2, Optional N1$ = "Exp", Optional N2$ = "Act") As String()
Dim U&: U = UB(Ay1)
Dim O$()
    If U <> UB(Ay2) Then Push O, FmtQQ("Array [?] and [?] has different Si: [?] [?]", N1, N2, Si(Ay1), Si(Ay2)): GoTo X
If Si(Ay1) = 0 Then Exit Function
Dim O1$()
    Dim A2: A2 = Ay2
    Dim J&, ReachLimit As Boolean
    Dim Cnt%
    For J = 0 To U
        If Ay1(J) <> Ay2(J) Then
            Push O1, FmtQQ("[?]-th Ele is diff: ?[?]<>?[?]", N1, N2, Ay1(J), Ay2(J))
            Cnt = Cnt + 1
        End If
        If Cnt > 10 Then
            ReachLimit = True
            Exit For
        End If
    Next
'If IsEmp(O1) Then Exit Function
Dim O2$()
    Push O2, FmtQQ("Array [?] and [?] both having size[?] have differnt element(s):", N1, N2, Si(Ay1))
    If ReachLimit Then
        Push O2, FmtQQ("At least [?] differences:", Si(O1))
    End If
PushAy O, O2
PushAy O, O1
X:
Push O, FmtQQ("Ay-[?]:", N1)
PushIAy O, QteSy(SyzAy(Ay1), "[]")
Push O, FmtQQ("Ay-[?]:", N2)
PushIAy O, QteSy(SyzAy(Ay2), "[]")
ChkEqAy = O
End Function

Function AyOfAyAy(AyOfAy)
If Si(AyOfAy) = 0 Then Exit Function
Dim O
O = ResiU(AyOfAy(0))
Dim X
For Each X In AyOfAy
    PushAy O, X
Next
AyOfAyAy = O
End Function

Sub Z_AyFlat()
Dim AyOfAy()
AyOfAy = Array(SyzSS("a b c d"), SyzSS("a b c"))
Ept = SyzSS("a b c d a b c")
GoSub Tst
Exit Sub
Tst:
    Act = AyFlat(AyOfAy)
    C
    Return
End Sub

Function AyFlat(AyOfAy)
AyFlat = AyOfAyAy(AyOfAy)
End Function

Function AyItmCnt%(Ay, M)
If Si(Ay) = 0 Then Exit Function
Dim O%, X
For Each X In Itr(Ay)
    If X = M Then O = O + 1
Next
AyItmCnt = O
End Function

Function NxtNm$(Ny$(), Optional MaxN% = 0)
Stop
End Function

Function NxtFn$(Fn$, FnAy$(), Optional MaxN% = 999)
If Not HasEle(FnAy, Fn) Then NxtFn = Fn: Exit Function
NxtFn = AyMax(AwLik(FnAy, Fn & "(???)"))
End Function

Function ItrzLines(Lines$)
Asg Itr(SplitCrLf(Lines$)), ItrzLines
End Function

Function NItr&(Itr)
Dim O&, V
For Each V In Itr
    O = O + 1
Next
NItr = O
End Function

Function ItrzAy(Ay)
ItrzAy = Itr(Ay)
End Function

Function ItwNm(Itr, Nm)
Dim O: For Each O In Itr
    If O.Name = Nm Then Asg O, ItwNm
Next
End Function

Function Itr(Ay)
If Si(Ay) = 0 Then Set Itr = New Collection Else Itr = Ay
End Function

Function RSyzTrim(Ay) As String()
If Si(Ay) = 0 Then Exit Function
Dim O$(), I
For Each I In Ay
    Push O, RTrim(I)
Next
RSyzTrim = O
End Function

Function ResiN(Ay, N&)
'Ret : empty ay of si @N of sam base ele as @Ay @@
ResiN = ResiU(Ay, N - 1)
End Function
Function IFin(V, Ay)
'Ret @V if in @Ay else Empty
If HasEle(Ay, V) Then IFin = V
End Function
Function ResiMax(OAy1, OAy2)
'Ret : resi the min si of ay to sam si as the other @@
Dim U1&, U2&: U1 = UB(OAy1): U2 = UB(OAy2)
Select Case True
Case U1 > U2: OAy2 = ResiU(OAy2, U1)
Case U2 > U1: OAy1 = ResiU(OAy1, U2)
End Select
ResiMax = OAy1
End Function

Function ResiU(Ay, Optional U& = -1)
'Ret : new ay redim preserve @Ay to @U
Dim O: O = Ay
If U < 0 Then Erase O: ResiU = O: Exit Function
ReDim Preserve O(U)
ResiU = O
End Function

Function AyRev(Ay) 'Return reversed Ay
Dim O: O = Ay
Dim J&, U&
U = UB(O)
For J = 0 To U
    Asg Ay(U - J), O(J)
Next
AyRev = O
End Function

Function AyRevI(Ay)
Dim O: O = Ay
Dim J&, U&
U = UB(O)
For J = 0 To U
    O(J) = Ay(U - J)
Next
AyRevI = O
End Function

Function AyRevOy(Oy() As Object)
Dim O: O = Oy
Dim J&, U&
U = UB(O)
For J = 0 To U
    Set O(J) = Oy(U - J)
Next
AyRevOy = O
End Function

Function AmRplzMid(Ay, B As Fei, ByAy)
With AyabczAyFei(Ay, B)
AmRplzMid = AddAyAp(.A, ByAy, .C)
End With
End Function

Function SampLinzAy$(Ay)
Dim S$, U&
U = UB(Ay)
If U >= 0 Then
    Select Case True
    Case IsPrim(Ay(0)): S = "[" & Ay(0) & "]"
    Case IsObject(Ay(0)), IsArray(Ay(0)): S = "[*Ty:" & TypeName(Ay(0)) & "]"
    Case Else: Stop
    End Select
End If
SampLinzAy = "*Ay:[" & U & "]" & S
End Function

Function SeqDiKqCnt(Ay) As Dictionary 'The return dic of key=AyEle pointing to 2-Ele-LngAp with Ele-0 as Seq#(0..) and Ele- as Cnt
Dim S&, O As New Dictionary, L&(), X
For Each X In Itr(Ay)
    If O.Exists(X) Then
        L = O(X)
        L(1) = L(1) + 1
        O(X) = L
    Else
        ReDim L(1)
        L(0) = S
        L(1) = 1
        O.Add X, L
    End If
Next
Set SeqDiKqCnt = O
End Function
Function StrColzSq(Sq(), Optional C = 1) As String()
If Si(Sq) = 0 Then Exit Function
Dim R&
For R = 1 To UBound(Sq, 1)
    PushI StrColzSq, Sq(R, C)
Next
End Function
Function SqHzAp(ParamArray Ap()) As Variant()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
SqHzAp = SqH(Av)
End Function

Function SqH(Ay) As Variant()
Dim N&: N = Si(Ay)
If N = 0 Then Exit Function
Dim J&, V
Dim O()
ReDim O(1 To 1, 1 To N)
For Each V In Ay
    J = J + 1
    O(1, J) = V
Next
SqH = O
End Function

Function SqV(Ay) As Variant()
Dim N&: N = Si(Ay)
If N = 0 Then Exit Function
Dim J&, V
Dim O()
ReDim O(1 To N, 1 To 1)
For Each V In Ay
    J = J + 1
    O(J, 1) = V
Next
SqV = O
End Function


Function IndentSy(Sy$(), Optional Indent% = 4) As String()
Dim I, S$
S = Space(Indent)
For Each I In Itr(Sy)
    PushI IndentSy, S & I
Next
End Function


Function AyMinzGT0(Ay)
If Si(Ay) = 0 Then Exit Function
Dim O: O = Ay(0)
Dim V: For Each V In Ay
    If V > 0 Then
        If O = 0 Then
            O = V
        Else
            If V < O Then O = V
        End If
    End If
Next
AyMinzGT0 = O
End Function

Function AySum#(NumAy)
Dim O#, V: For Each V In Itr(NumAy)
    O = O + V
Next
AySum = O
End Function

Function AyMin(Ay)
Dim N&: N = Si(Ay)
If N = 0 Then Exit Function
Dim O: O = Ay(0)
Dim J&: For J = 1 To N - 1
    If Ay(J) < O Then O = Ay(J)
Next
AyMin = O
End Function

Function AyMax(Ay)
If Si(Ay) = 0 Then Exit Function
Dim O, I
For Each I In Ay
    If I > O Then O = I
Next
AyMax = O
End Function

Function WdtzAy%(Ay)
Dim O%, V
For Each V In Itr(Ay)
    O = Max(O, Len(V))
Next
WdtzAy = O
End Function

Sub WrtAy(Ay, Ft$, Optional OvrWrt As Boolean)
WrtStr JnCrLf(Ay), Ft, OvrWrt
End Sub

Function AyLTrim(Ay) As String()
Dim L
For Each L In Itr(Ay)
    PushI AyLTrim, LTrim(L)
Next
End Function
Function SyEnsSfxDot(Ay) As String()
SyEnsSfxDot = SyEnsSfx(Sy, ".")
End Function
Function SyEnsSfx(Sy$(), Sfx$) As String()
Dim I: For Each I In Itr(Sy)
    PushI SyEnsSfx, EnsSfx(I, Sfx)
Next
End Function
Function StmtLy(StmtLin) As String()
StmtLy = SyEnsSfxDot(AyLTrim(Split(StmtLin, ". ")))
End Function
Function AyZip(A1, A2) As Variant()
Dim U1&: U1 = UB(A1)
Dim U2&: U2 = UB(A2)
Dim U&: U = Max(U1, U2)
Dim O()
    Dim J&
    O = ResiU(O, U)
    For J = 0 To U
        If U1 >= J Then
            If U2 >= J Then
                O(J) = Array(A1(J), A2(J))
            Else
                O(J) = Array(A1(J), Empty)
            End If
        Else
            If U2 >= J Then
                O(J) = Array(, A2(J))
            Else
                Stop
            End If
        End If
    Next
AyZip = O
End Function

Function AyZip_Ap(Ay, ParamArray Ap()) As Variant()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
Dim UCol%
    UCol = UB(Av)

Dim URow1&
    URow1 = UB(Ay)

Dim URow&
Dim URowAy&()
    Dim J%, IURow%
    URow = URow1
    For J = 0 To UB(Av)
        IURow = UB(Av(J))
        Push URowAy, IURow
        If IURow > URow Then URow = IURow
    Next

Dim ODy()
    Dim Dr()
    ODy = ResiU(ODy, URow)
    Dim I%
    For J = 0 To URow
        Erase Dr
        If URow1 >= J Then
            Push Dr, Ay(J)
        Else
            Push Dr, Empty
        End If
        For I = 0 To UB(Av)
            If URowAy(I) >= J Then
                Push Dr, Av(I)(J)
            Else
                Push Dr, Empty
            End If
        Next
        ODy(J) = Dr
    Next
AyZip_Ap = ODy
End Function

Function ItmAddAy(Itm, Ay)
ItmAddAy = InsEle(Ay, Itm)
End Function

Sub Z_AyabczAyFE()
Dim Ay(): Ay = Array(1, 2, 3, 4)
Dim Act As Ayabc: Act = AyabczAyFE(Ay, 1, 2)
Ass IsEqAy(Act.A, Array(1))
Ass IsEqAy(Act.B, Array(2, 3))
Ass IsEqAy(Act.C, Array(4))
End Sub

Sub Z_AyAsgAp()
Dim O%, Ay$
'AyAsgAp Array(234, "abc"), O, Ay
Ass O = 234
Ass Ay = "abc"
End Sub

Sub Z_ChkEqAy()
DmpAy ChkEqAy(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Sub Z_MaxEle()
Dim Ay()
Dim Act
Act = MaxEle(Ay)
Stop
End Sub

Sub Z_AyMinus()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyMinus(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
ThwIf_NE Exp, Act
'
Act = AyMinusAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
ThwIf_NE Exp, Act
End Sub

Sub Z_SyzAy()
Dim Act$(): Act = SyzAy(Array(1, 2, 3))
Ass Si(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Sub Z_AmTrim()
DmpAy AmTrim(Sy(1, 2, 3, "  a"))
End Sub


Sub Z_ChkDup()
Dim Ay
Ay = Array("1", "1", "2")
Ept = Sy("This item[1] is duplicated")
GoSub Tst
Exit Sub
Tst:
    Act = ChkDup(Ay, "This item[?] is duplicated")
    C
    Return
End Sub

Sub Z_ChkEqAy5()
DmpAy ChkEqAy(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Sub Z_AyabczAyFei()
Dim Ay(): Ay = Array(1, 2, 3, 4)
Dim M As Fei: M = Fei(1, 2)
Dim Act As Ayabc: Act = AyabczAyFei(Ay, M)
Ass IsEqAy(Act.A, Array(1))
Ass IsEqAy(Act.B, Array(2, 3))
Ass IsEqAy(Act.C, Array(4))
End Sub

Sub Z_HasDupEle()
Ass HasDupEle(Array(1, 2, 3, 4)) = False
Ass HasDupEle(Array(1, 2, 3, 4, 4)) = True
End Sub

Sub Z_AyInsAy()
Dim Act, Exp, Ay(), B(), At&
Ay = Array(1, 2, 3, 4)
B = Array("X", "Z")
At = 1
Exp = Array(1, "X", "Z", 2, 3, 4)

Act = InsAy(Ay, B, At)
Ass IsEqAy(Act, Exp)
End Sub

Sub Z_AyMinus6()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyMinus(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
ThwIf_AyabNE Exp, Act
'
Act = AyMinusAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
ThwIf_AyabNE Exp, Act
End Sub

Sub Z_SyzAy2()
Dim Act$(): Act = SyzAy(Array(1, 2, 3))
Ass Si(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Sub Z_AmTrim2()
DmpAy AmTrim(Sy(1, 2, 3, "  a"))
End Sub

Sub Z_KKCMiDy()
Dim Dy(), Act As KKCntMulItmColDy, KKColIx%(), IxzAy%
PushI Dy, Array()
PushI Dy, Array()
PushI Dy, Array()
PushI Dy, Array()
PushI Dy, Array()
PushI Dy, Array()
'Ass Si(Act) = 4
'Ass IsEqAy(Act(0), Array("Ay", 3, 1, 2, 3))
'Ass IsEqAy(Act(1), Array("B", 3, 2, 3, 4))
'Ass IsEqAy(Act(2), Array("C", 0))
'Ass IsEqAy(Act(3), Array("D", 1, "X"))
End Sub



Sub Z_AddPfxzSslIn()
Dim Ssl$, Exp$(), Pfx$
Ssl = "B C D"
Pfx = "A"
Exp = SyzSS("AB AC AD")
GoSub Tst
Exit Sub
Tst:
    Dim Act$()
    Act = AddPfxzSslIn(Pfx, Ssl)
    Debug.Assert IsEqAy(Act, Exp)
Return
End Sub

Function AddPfxzSslIn(Pfx$, SsLin) As String()
AddPfxzSslIn = AmAddPfx(SyzSS(SsLin), Pfx)
End Function

Function SpcSepStr$(S)
If S = "" Then SpcSepStr = ".": Exit Function
SpcSepStr = QteSqIf(EscSqBkt(SlashCrLf(EscBackSlash(S))))
End Function

Function RevSS$(SS)
If SS = "." Then Exit Function
RevSS = UnTidleSpc(UnSlashTab(UnSlashCrLf(SS)))
End Function

Function SslzDr$(Dr)
Dim J&, U&, O$()
U = UB(Dr)
If U < 0 Then Exit Function
ReDim O(U)
For J = 0 To U
    O(J) = SpcSepStr(Dr(J))
Next
SslzDr = JnSpc(O)
End Function

Function IsSyDte(Sy$()) As Boolean
Dim S: For Each S In Sy
    If Not IsDteStr(S) Then Exit Function
Next
End Function

Function IsSyDbl(Sy$()) As Boolean
Dim S: For Each S In Sy
    If Not IsDblStr(S) Then Exit Function
Next
End Function

Function DteAyzSy(Sy$()) As Date()
Dim I: For Each I In Sy
    PushI DteAyzSy, I
Next
End Function

Function DblAyzSy(Sy$()) As Double()
Dim I: For Each I In Sy
    PushI DblAyzSy, I
Next
End Function

Function ItrzSS(SS)
ItrzSS = Itr(SyzSS(SS))
End Function

Function SrtSS$(SS$)
SrtSS = JnSpc(AySrt(SyzSS(SS)))
End Function
Function SyzSS(SS) As String()
SyzSS = SplitSpc(RplDblSpc(Trim(SS)))
End Function

Function ItrzTT(TT$)
Asg Itr(TermAy(TT)), ItrzTT
End Function

Function IsEqSy(A$(), B$()) As Boolean
If Not IsEqSi(A, B) Then Exit Function
Dim J&, X
For Each X In Itr(A)
    If X <> B(J) Then Exit Function
    J = J + 1
Next
IsEqSy = True
End Function

Function IsEqDr(A, B) As Boolean
Dim X, J&
For Each X In Itr(A)
    If X <> B(J) Then Exit Function
    J = J + 1
Next
IsEqDr = True
End Function

Function IsEqAy(A, B) As Boolean
If Not IsArray(A) Then Exit Function
If Not IsArray(B) Then Exit Function
If Not IsEqSi(A, B) Then Exit Function
Dim J&, X
For Each X In Itr(A)
    If Not IsEq(X, B(J)) Then Exit Function
    J = J + 1
Next
IsEqAy = True
End Function
Function SqzDy(Dy(), Optional SKipNRow& = 1) As Variant()
Dim O(), NR&, NC&
NR = Si(Dy)
NC = NColzDy(Dy)
ReDim O(1 To NR, 1 To NC)
Dim R&: For R = 1 To NR
    Dim Dr: Dr = Dy(R - 1)
    SetSqr O, Dr, R
Next
SqzDy = O
End Function


Function AyzAyOfAy(AyOfAy())
Dim O: O = AyOfAy(0)
Dim J&: For J = 1 To UB(AyOfAy)
    PushIAy O, AyOfAy(J)
Next
AyzAyOfAy = O
End Function


