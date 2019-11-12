Attribute VB_Name = "MxAyStd"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAyStd."


Function AeAtCnt(Ay, Optional At = 0, Optional Cnt = 1)
If Cnt <= 0 Then Thw CSub, "Cnt cannot <=0", "At Cnt Ay", At, Cnt, Ay
If Si(Ay) = 0 Then AeAtCnt = Ay: Exit Function
If At = 0 Then
    If Si(Ay) = Cnt Then
        AeAtCnt = ResiU(Ay)
        Exit Function
    End If
End If
Dim U&: U = UB(Ay)
If At > U Then Stop
If At < 0 Then Stop
Dim O: O = Ay
Dim J&
For J = At To U - Cnt
    Asg O(J + Cnt), O(J)
Next
ReDim Preserve O(U - Cnt)
AeAtCnt = O
End Function

Function AeBlnk(Ay) As String()
Dim I: For Each I In Itr(Ay)
    If Trim(I) <> "" Then PushI AeBlnk, I
Next
End Function

Function AeBlnkEleAtEnd(A$()) As String()
If Si(A) = 0 Then Exit Function
If LasEle(A) <> "" Then AeBlnkEleAtEnd = A: Exit Function
Dim J%
For J = UB(A) To 0 Step -1
    If Trim(A(J)) <> "" Then
        Dim O$()
        O = A
        ReDim Preserve O(J)
        AeBlnkEleAtEnd = O
        Exit Function
    End If
Next
End Function

Function AeBlnkStr(Ay) As String()
Dim X
For Each X In Itr(Ay)
    If Trim(X) <> "" Then
        PushI AeBlnkStr, X
    End If
Next
End Function

Function AeEle(Ay, Ele) 'Rmv Fst-Ele eq to Ele from Ay
AeEle = AePred(Ay, PredzEq(Ele))
End Function

Function AeEleAt(Ay, Optional At& = 0, Optional Cnt& = 1)
AeEleAt = AeAtCnt(Ay, At, Cnt)
End Function

Function AeEleLik(Ay, Lik$) As String()
If Si(Ay) = 0 Then Exit Function
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) Like Lik Then AeEleLik = AeEleAt(Ay, J): Exit Function
Next
End Function

Function AeEmpEle(Ay)
Dim O: O = ResiU(Ay)
If Si(Ay) > 0 Then
    Dim X
    For Each X In Itr(Ay)
        PushNonEmp O, X
    Next
End If
AeEmpEle = O
End Function

Function AeEmpEleAtEnd(Ay)
Dim LasU&, U&
Dim O: O = Ay
For LasU = UB(Ay) To 0 Step -1
    If Not IsEmp(O(LasU)) Then
        Exit For
    End If
Next
If LasU = -1 Then
    Erase O
Else
    ReDim Preserve O(LasU)
End If
AeEmpEleAtEnd = O
End Function

Function AeFei(Ay, b As Fei)
With b
    AeFei = AeFmTo(Ay, .FmIx, .EIx)
End With
End Function

Function AeFmTo(Ay, FmIx, EIx)
Const CSub$ = CMod & "AeFmTo"
Dim U&
U = UB(Ay)
If 0 > FmIx Or FmIx > U Then Thw CSub, "[FmIx] is out of range", "U FmIx EIx Ay", UB(Ay), FmIx, EIx, Ay
If FmIx > EIx Or EIx > U Then Thw CSub, "[EIx] is out of range", "U FmIx EIx Ay", UB(Ay), FmIx, EIx, Ay
Dim O
    O = Ay
    Dim I&, J&
    I = 0
    For J = EIx + 1 To U
        O(FmIx + I) = O(J)
        I = I + 1
    Next
    Dim Cnt&
    Cnt = EIx - FmIx + 1
    ReDim Preserve O(U - Cnt)
AeFmTo = O
End Function

Function AeFstEle(Ay)
AeFstEle = AeEleAt(Ay)
End Function

Function AeFstLas(Ay)
Dim J&
AeFstLas = Ay
Erase AeFstLas
For J = 1 To UB(Ay) - 1
    PushI AeFstLas, Ay(J)
Next
End Function

Function AeFstNEle(Ay, Optional N& = 1)
Dim O: O = ResiU(Ay)
Dim J&
For J = N To UB(Ay)
    Push O, Ay(J)
Next
AeFstNEle = O
End Function

Function AeIxSet(Ay, IxSet As Aset)
Dim J&, O
O = Ay: Erase O
For J = 0 To UBound(Ay)
    If Not IxSet.Has(J) Then PushI O, Ay(J)
Next
AeIxSet = O
End Function

Function AeIxy(Ay, SrtdIxy)
'Fm SrtdIxy : holds index if Ay to be remove.  It has been sorted else will be stop
Ass IsArray(Ay)
Ass IsSrtdzAy(SrtdIxy)
Dim J&
Dim O: O = Ay
For J = UB(SrtdIxy) To 0 Step -1
    O = AeEleAt(O, CLng(SrtdIxy(J)))
Next
AeIxy = O
End Function

Function AeKss(Ay, ExlKss$) As String()
Stop
'AeKss = AePred(Ay, PredzIsKss(ExlKss))
End Function

Function AeLasEle(Ay)
AeLasEle = AeEleAt(Ay, UB(Ay))
End Function

Function AeLasNEle(Ay, Optional NEle% = 1)
If NEle = 0 Then AeLasNEle = Ay: Exit Function
Dim O: O = Ay
Select Case Si(Ay)
Case Is > NEle:    ReDim Preserve O(UB(Ay) - NEle)
Case NEle: Erase O
Case Else: Stop
End Select
AeLasNEle = O
End Function

Function AeLik(Ay, Lik$) As String()
AeLik = AePred(Ay, PredzLik(Lik))
End Function

Function AeNegative(Ay)
Dim I
AeNegative = ResiU(Ay)
For Each I In Itr(Ay)
    If I >= 0 Then
        PushI AeNegative, I
    End If
Next
End Function

Function AeNEle(Ay, Ele, Cnt%)
If Cnt <= 0 Then Stop
AeNEle = ResiU(Ay)
Dim X, C%
C = Cnt
For Each X In Itr(Ay)
    If C = 0 Then
        PushI AeNEle, X
    Else
        If X <> Ele Then
            Push AeNEle, X
        Else
            C = C - 1
        End If
    End If
Next
X:
End Function

Function AePatn(Ay, Patn$) As String()
Dim I, Re As New RegExp
Re.Pattern = Patn
For Each I In Itr(Ay)
    If Not Re.Test(I) Then PushI AePatn, I
Next
End Function

Function AePfx(Ay, Pfx$) As String()
AePfx = AePred(Ay, PredzPfx(Pfx))
End Function

Function AePred(Ay, P As IPred) As String()
Dim I: For Each I In Itr(Ay)
    If Not P.Pred(I) Then
        PushI AePred, I
    End If
Next
End Function

Function AeRe(Ay, Re As RegExp) As String()
Dim I
For Each I In Itr(Ay)
    If Not Re.Test(I) Then PushI AeRe, I
Next
End Function

Function AeSngQRmk(Ay) As String()
Dim I, S$
For Each I In Itr(Ay)
    S = I
    If Not IsSngQRmk(S) Then PushI AeSngQRmk, S
Next
End Function

Function Am(Ay, Fun$) As Variant()
Dim I: For Each I In Itr(Ay)
    PushI Am, Run(Fun, I)
Next
End Function
Function AmAddPfx(Ay, Pfx) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmAddPfx, Pfx & I
Next
End Function

Function AmAddPfxS(Ay, Pfx, Sfx) As String()
Dim I
For Each I In Itr(Ay)
    PushI AmAddPfxS, Pfx & I & Sfx
Next
End Function

Function AmAddSfx(Ay, Sfx) As String()
Dim I
For Each I In Itr(Ay)
    PushI AmAddSfx, I & Sfx
Next
End Function

Function AmAft(Sy$(), Sep$) As String()
Dim S: For Each S In Itr(Sy)
    Push AmAft, Aft(S, Sep)
Next
End Function

Function AmAftRev(Sy$(), Sep$) As String()
Dim S: For Each S In Itr(Sy)
    Push AmAftRev, AftRev(S, Sep)
Next
End Function

Function AmAlign(Ay) As String()
Dim W%: W = WdtzAy(Ay)
Dim I: For Each I In Itr(Ay)
    PushI AmAlign, Align(I, W)
Next
End Function

Function AmAlignR(Ay) As String()
Dim W%: W = WdtzAy(Ay)
Dim I: For Each I In Itr(Ay)
    PushI AmAlignR, AlignR(I, W)
Next
End Function

Function AmBef(Sy$(), Sep$) As String()
Dim S: For Each S In Itr(Sy)
    Push AmBef, Bef(S, Sep)
Next
End Function

Function AmIncEleBy1(NumAy)
AmIncEleBy1 = AmIncEleByN(NumAy, 1)
End Function

Function AmIncEleByN(NumAy, N)
Dim O: O = NumAy
Dim J&
For J = 0 To UB(O)
    O(J) = O(J) + N
Next
AmIncEleByN = O
End Function

Function AmRmv2Dash(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmv2Dash, Rmv2Dash(I)
Next
End Function

Function AmRmvLasChr(Ay) As String()
'Gen:AyFor RmvLasChr
Dim I
For Each I In Itr(Ay)
    PushI AmRmvLasChr, RmvLasChr(I)
Next
End Function

Function AmRmvPfx(Ay, Pfx$) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmvPfx, RmvPfx(I, Pfx)
Next
End Function

Function AmRmvT1(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmvT1, RmvT1(I)
Next
End Function

Function AmRpl(Ay, Fm$, By$, Optional Cnt& = 1, Optional IgnCas As Boolean) As String()
Dim I
For Each I In Itr(Ay)
    PushI AmRpl, Replace(I, Fm, By, Count:=Cnt, Compare:=CprMth(IgnCas))
Next
End Function

Function AmRTrim(Sy$()) As String()
Dim S: For Each S In Itr(Sy)
    Push AmRTrim, RTrim(S)
Next
End Function

Function AmSyzzSS(Ly$()) As Variant()
Dim L: For Each L In Itr(Ly)
    PushI AmSyzzSS, SyzSS(L)
Next
End Function

Function AmT1(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmT1, T1(I)
Next
End Function

Function AmT2(Ay) As String()
Dim L: For Each L In Itr(Ay)
    PushI AmT2, T2(L)
Next
End Function

Function AmT3(Ay) As String()
Dim L: For Each L In Itr(Ay)
    PushI AmT3, T3(L)
Next
End Function

Function AmTab(Ay, Optional NTab% = 1) As String()
AmTab = AmAddPfx(Ay, TabN(NTab))
End Function

Function AmTrim(Sy$()) As String()
Dim S: For Each S In Itr(Sy)
    Push AmTrim, Trim(S)
Next
End Function

Function AmzRmvT1(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmzRmvT1, RmvT1(I)
Next
End Function

Function AmzRmvTT(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmzRmvTT, RmvTT(I)
Next
End Function

Function AwAftEle(Ay, Ele)
Dim O: O = Ay: Erase O
Dim I, F As Boolean: For Each I In Itr(Ay)
    If F Then
        PushI O, I
    Else
        If I = Ele Then F = True
    End If
   
Next
Thw CSub, "No @Ele in @Ay", "Ele Ay", Ele, Ay
End Function

Function AwBefEle(Ay, Ele)
Dim O: O = Ay: Erase O
Dim I: For Each I In Itr(Ay)
    PushI O, I
    If I = Ele Then AwBefEle = O: Exit For
Next
Thw CSub, "No @Ele in @Ay", "Ele Ay", Ele, Ay
End Function

Function AwBet(Ay, FmEle, ToEle)
Dim O: O = ResiU(Ay)
Dim I: For Each I In Itr(Ay)
    If IsBet(I, FmEle, ToEle) Then
        Push O, I
    End If
Next
AwBet = O
End Function

Function AwDist(Ay, Optional IgnCas As Boolean)
AwDist = IntozItr(ResiU(Ay), DiKqCnt(Ay).Keys)
End Function

Function AwDistAsI(Ay, Optional IgnCas As Boolean) As Integer()
AwDistAsI = CvIntAy(AwDist(Ay, IgnCas))
End Function

Function AwDistAsSy(Ay, Optional IgnCas As Boolean) As String()
AwDistAsSy = CvSy(AwDist(Ay, IgnCas))
End Function

Function AwDistT1(Ay) As String()
AwDistT1 = AwDist(AmT1(Ay))
End Function

Function AwDup(Ay, Optional C As VbCompareMethod = vbTextCompare)
Dim O: O = ResiU(Ay)
Dim D As Dictionary: Set D = DiKqCnt(Ay, EiCntDup, C)
AwDup = IntozItr(O, D.Keys)
End Function

Function AwEQ(Ay, V)
If Si(Ay) <= 1 Then AwEQ = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I = V Then PushI O, I
Next
AwEQ = O
End Function

Function AwFE(Ay, FmIx, EIx)
AwFE = AwFT(Ay, FmIx, EIx - 1)
End Function

Function AwFei(Ay, b As Fei)
AwFei = AwFE(Ay, b.FmIx, b.EIx)
End Function

Function AwFm(Ay, FmIx)
Dim J&, I&
Dim O: O = ResiU(Ay, UB(Ay) - FmIx)
For J = FmIx To UB(Ay)
    Asg Ay(J), O(I)
    I = I + 1
Next
AwFm = O
End Function

Function AwFstUEle(Ay, U)
If U > UB(Ay) Then AwFstUEle = Ay: Exit Function
Dim O: O = Ay
ReDim Preserve O(U)
AwFstUEle = O
End Function

Function AwFT(Ay, FmIx, ToIx)
Dim J&, I&
Dim O: O = ResiU(Ay, ToIx - FmIx)
For J = FmIx To ToIx
    Asg Ay(J), O(I)
    I = I + 1
Next
AwFT = O
End Function

Function AwGT(Ay, V)
If Si(Ay) = 1 Then AwGT = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I > V Then PushI O, I
Next
AwGT = O
End Function

Function AwInAset(Ay, b As Aset)
AwInAset = ResiU(Ay)
Dim I
For Each I In Itr(Ay)
    If Ay.Has(I) Then Push AwInAset, I
Next
End Function

Function AwIsNm(Ay) As String()
AwIsNm = AwPred(Ay, PredIsNm)
End Function

Function AwIxCnt(Ay, Ix, Cnt)
Dim J&
Dim O: O = Ay: Erase O
For J = 0 To Cnt - 1
    Push O, Ay(Ix + J)
Next
AwIxCnt = O
End Function

Function AwIxy(Ay, Ixy&())
If Si(Ixy) = 0 Then AwIxy = Ay: Erase AwIxy: Exit Function
Dim U&: U = UB(Ixy)
Dim O: O = Ay: ReDim O(U)
Dim Ix, J&
For Each Ix In Itr(Ixy)
    If IsObject(Ay(Ix)) Then
        Set O(J) = Ay(Ix)
    Else
        O(J) = Ay(Ix)
    End If
    J = J + 1
Next
AwIxy = O
End Function

Function AwIxyAlwE(Ay, Ixy&())
Dim U&: U = UB(Ixy)
Dim O: O = ResiU(Ay)
ReDim Preserve O(U)
Dim Ix, J&
For Each Ix In Itr(Ixy)
    If Ix >= 0 Then
        If IsObject(Ay(Ix)) Then
            Set O(J) = Ay(Ix)
        Else
            O(J) = Ay(Ix)
        End If
    End If
    J = J + 1
Next
AwIxyAlwE = O
End Function

Function AwIxyzMust(Ay, Ixy&())
If IsIxyOut(Ixy, UB(Ay)) Then Thw CSub, "Some element in Ixy is outsize Ay", "UB(Ay) Ixy", UB(Ay), Ixy
Dim U&: U = UB(Ay)
Dim O: O = ResiU(Ay)
Dim Ix
For Each Ix In Itr(Ixy)
    Push O, Ay(Ix)
Next
AwIxyzMust = O
End Function

Function AwKssAy(Ay, KssAy$()) As String()
Dim LikAy$(): LikAy = LikAyzKssAy(KssAy)
Dim S: For Each S In Itr(Ay)
    If HitLikAy(S, LikAy) Then PushI AwKssAy, S
Next
End Function

Function AwLasN(Ay, N)
Dim O, J&, I&, U&, Fm&, NewU&
U = UB(Ay)
If U < N Then AwLasN = Ay: Exit Function
O = Ay
Fm = U - N + 1
NewU = N - 1
For J = Fm To U
    Asg O(J), O(I)
    I = I + 1
Next
ReDim Preserve O(NewU)
AwLasN = O
End Function

Function AwLE(Ay, V)
If Si(Ay) <= 1 Then AwLE = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I <= V Then PushI O, I
Next
AwLE = O
End Function

Function AwLik(Ay, Lik) As String()
Dim I: For Each I In Itr(Ay)
    If I Like Lik Then PushI AwLik, I
Next
End Function

Function AwLikAy(Ay, LikAy$()) As String()
Dim I, Lik
For Each I In Itr(Ay)
    If HitLikAy(I, LikAy) Then PushI AwLikAy, I
Next
End Function

Function AwLikss(Ay, Likss$) As String()
AwLikss = AwLikAy(Ay, SyzSS(Likss))
End Function

Function AwLT(Ay, V)
If Si(Ay) = 1 Then AwLT = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I < V Then PushI O, I
Next
AwLT = O
End Function

Function AwMid(Ay, Fm, Optional L = 0)
AwMid = ResiU(Ay)
Dim J&
Dim E&
    Select Case True
    Case L = 0: E = UB(Ay)
    Case Else:  E = Min(UB(Ay), L + Fm - 1)
    End Select
For J = Fm To E
    Push AwMid, Ay(J)
Next
End Function

Function AwmRmvT1(Ly$(), T1) As String()
Dim L: For Each L In Itr(Ly)
    If ShfT1(L) = T1 Then PushI AwmRmvT1, L
Next
End Function

Function AwNm(Ay, b As WhNm) As String()
Dim I
For Each I In Itr(Ay)
    If HitNm(I, b) Then PushI AwNm, I
Next
End Function

Function AwNmStr(Ay, WhNmStr$) As String()
AwNmStr = AwNm(Ay, WhNmzS(WhNmStr))
End Function

Function AwNonEmp(Ay)
AwNonEmp = Ay: Erase AwNonEmp
Dim I
For Each I In Ay
    If Not IsEmpty(I) Then
        PushI AwNonEmp, I
    End If
Next
End Function

Function AwNoNm(Sy$()) As String()
Dim Nm$, I
For Each I In Sy
    Nm = I
    If IsNm(Nm) Then PushI AwNoNm, Nm
Next
End Function

Function AwPatn(Ay, Patn$) As String()
If Si(Ay) = 0 Then Exit Function
If Patn = "" Or Patn = "." Then AwPatn = Ay: Exit Function
AwPatn = AwPred(Ay, PredzPatn(Patn))
End Function

Function AwPatn1(Ay, Patn$) As Variant()
If Si(Ay) = 0 Then Exit Function
If Patn = "" Or Patn = "." Then AwPatn1 = Ay: Exit Function
Dim Re As RegExp: Set Re = Rx(Patn)
Dim I: For Each I In Itr(Ay)
    If IsStr(I) Then
        If Re.Test(I) Then PushI AwPatn1, I
    End If
Next
End Function

Function AwPatnAy(Ay, PatnAy$()) As Variant()
If Si(Ay) = 0 Then Exit Function
Stop
End Function

Function AwPatnExl(Ay, Patn$, ExlKss$) As String()
AwPatnExl = AeKss(AwPatn(Ay, Patn), ExlKss)
End Function

Function AwPfx(Ay, Pfx) As String()
AwPfx = AwPred(Ay, PredzPfx(Pfx))
End Function

Function AwPred(Ay, P As IPred) As String()
Dim I
For Each I In Itr(Ay)
    If P.Pred(I) Then
        PushI AwPred, I
    End If
Next
End Function

Function AwPredFalse(Ay, P As IPred)
Dim X
AwPredFalse = ResiU(Ay)
For Each X In Itr(Ay)
    If Not P.Pred(X) Then
        Push AwPredFalse, X
    End If
Next
End Function

Function AwPredXAP(Ay, PredXAP$, ParamArray Ap())
AwPredXAP = ResiU(Ay)
Dim I
Dim Av()
    Av = Ap
    Av = InsEle(Av)
For Each I In Itr(Ay)
    Asg I, Av(0)
    If RunAv(PredXAP, Av) Then
        Push AwPredXAP, I
    End If
Next
End Function

Function AwPredXP(Ay, Xp$, P)
Dim X
AwPredXP = ResiU(Ay)
For Each X In Itr(Ay)
    If Run(Xp, X, P) Then
        Push AwPredXP, X
    End If
Next
End Function

Function AwPredXPNot(Ay, Xp$, P)
Dim X
AwPredXPNot = ResiU(Ay)
For Each X In Itr(Ay)
    If Not Run(Xp, X, P) Then
        Push AwPredXPNot, X
    End If
Next
End Function

Function AwPredzSy(Ay, P As IPred) As String()
Dim I
For Each I In Itr(Ay)
    If P.Pred(I) Then PushI AwPredzSy, I
Next
End Function

Function AwRe(Ay, Re As RegExp) As String()
Dim S
For Each S In Itr(Ay)
    If Re.Test(S) Then PushI AwRe, S
Next
End Function

Function AwRmvEle(Ay, Ele)
AwRmvEle = ResiU(Ay)
Dim I
For Each I In Itr(Ay)
    If I <> Ele Then PushI AwRmvEle, I
Next
End Function

Function AwRmvT1(Ay, T1) As String()
AwRmvT1 = AmzRmvT1(AwT1(Ay, T1))
End Function

Function AwRmvTT(Ay, T1, T2) As String()
AwRmvTT = AmzRmvTT(AwTT(Ay, T1, T2))
End Function

Function AwSfx(Ay, Sfx$) As String()
Dim I
For Each I In Itr(Ay)
    If HasSfx(I, Sfx) Then PushI AwSfx, I
Next
End Function

Function AwSingleEle(Ay)
Dim O: O = Ay: Erase O
Dim CntDy(): CntDy = DyoCntg(Ay)
If Si(CntDy) = 0 Then
    AwSingleEle = O
    Exit Function
End If
Dim Dr
For Each Dr In CntDy
    If Dr(1) = 1 Then
        Push O, Dr(0)
    End If
Next
AwSingleEle = O
End Function

Function AwSkip(Ay, Optional SkipN& = 1)
If SkipN <= 0 Then AwSkip = Ay: Exit Function
Dim U&: U = UB(Ay) - SkipN: If SkipN < -1 Then Thw CSub, "Ay is not enough to skip", "Si-Ay SkipN", "Si(Ay),SKipN"
Dim O: O = ResiU(Ay, U)
Dim J&, I&
For J = SkipN To UB(Ay)
    Asg Ay(J), O(I)
    I = I + 1
Next
End Function

Function AwSng(Ay)
AwSng = AyMinus(Ay, AwDup(Ay))
End Function

Function AwSngEle(Ay)
'Return Set of Element as array in {Ay} having 2 or more element
Dim O: O = ResiU(Ay)
Dim K, D As Dictionary
Set D = DiKqCnt(Ay)
For Each K In D.Keys
    If D(K) = 1 Then PushI O, K
Next
End Function

Function AwSubStr(Ay, SubStr) As String()
AwSubStr = AwPred(Ay, PredzSubStr(SubStr))
End Function

Function AwT1(Ay, T1) As String()
Dim L$, I
For Each I In Itr(Ay)
    L = I
    If HasT1(L, T1) Then
        PushI AwT1, L
    End If
Next
End Function

Function AwT1InAy(Ay, InAy) As String()
If Si(Ay) = 0 Then Exit Function
Dim O$(), L: For Each L In Ay
    If HasEle(InAy, T1(L)) Then Push O, L
Next
AwT1InAy = O
End Function

Function AwTT(Ay, T1, T2) As String()
Dim I, L$
For Each I In Itr(Ay)
    L = I
    If HasTT(L, T1, T2) Then PushI AwTT, L
Next
End Function

Function AwTTSelRst(Ay, T1, T2) As String()
Dim L$, I, X1$, X2$, Rst$
For Each I In Itr(Ay)
    L = I
    AsgTTRst L, X1, X2, Rst
    If X1 = T1 Then
        If X2 = T2 Then
            PushI AwTTSelRst, Rst
        End If
    End If
Next
End Function

Function ItrzAwRmvT1(Ay, T1)
Asg Itr(AwRmvT1(Ay, T1)), ItrzAwRmvT1
End Function

Function IxyzAyPatn(Ay, Patn$) As Long()
IxyzAyPatn = IxyzAyRe(Ay, Rx(Patn))
End Function

Function IxyzAyRe(Ay, b As RegExp) As Long()
If Si(Ay) = 0 Then Exit Function
Dim I, O&(), J&
For Each I In Ay
    If b.Test(I) Then Push O, J
    J = J + 1
Next
IxyzAyRe = O
End Function

Function IxyzCC(D As Drs, CC$) As Long()
IxyzCC = IxyzFF(D.Fny, CC)
End Function

Function IxyzFF(Fny$(), FF$) As Long()
IxyzFF = IxyzSubFny(Fny, SyzSS(FF))
End Function

Function IxyzSubAy(Ay, SubAy, Optional ThwNFnd As Boolean) As Long()
Dim E, Ix&
For Each E In SubAy
    Ix = IxzAy(Ay, E)
    If ThwNFnd Then
        If Ix = -1 Then
            Thw CSub, "Ele in SubAy not found in Ay", "Ele SubAy Ay", E, SubAy, Ay
        End If
    End If
    PushI IxyzSubAy, Ix
Next
End Function

Function IxyzSubFny(Fny$(), SubFny$()) As Long()
Dim F: For Each F In Itr(SubFny)
    Dim I&: I = IxzAy(Fny, F)
    If I >= 0 Then PushI IxyzSubFny, I
Next
End Function

Function LikAyzKssAy(KssAy$()) As String()
Dim Kss: For Each Kss In Itr(KssAy)
    PushIAy LikAyzKssAy, Kss
Next
End Function

Function PredzPatn(Patn$) As IPred
Dim O As New PredPatn
O.Init Patn
Set PredzPatn = O
End Function

Function PredzInT1Sy(AmT1$()) As IPred
Dim O As PredInT1Sy
O.Init AmT1
Set PredzInT1Sy = O
End Function

Function PredzIsOneTermLin() As IPred

End Function

Function PredzKssAy(KssAy$()) As PredLikAy
Dim O As New PredLikAy
O.Init KssAy
Set PredzKssAy = O
End Function

Function PredzLik(Lik$) As IPred
Dim O As New PredLik
O.Ini Lik
Set PredzLik = O
End Function

Function PredzEq(Eq) As IPred
Dim O As New PredEq
O.Ini Eq
Set PredzEq = O
End Function

Function PredzLikAy(LikAy$()) As PredLikAy
Set PredzLikAy = New PredLikAy
PredzLikAy.Init LikAy
End Function

Function PredzPfx(Pfx) As IPred
Dim O As New PredPfx
O.Init Pfx
Set PredzPfx = O
End Function

Function PredzSubStr(SubStr) As IPred
Dim O As New PredSubStr
O.Init SubStr
Set PredzSubStr = O
End Function

Function RmvBlnkzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    If Trim(I) <> "" Then
        PushI RmvBlnkzAy, I
    End If
Next
End Function

Function RmvFstChrzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI RmvFstChrzAy, RmvFstChr(I)
Next
End Function

Function RmvFstNonLetterzAy(Ay) As String() 'Gen:AyXXX
Dim I
For Each I In Itr(Ay)
    PushI RmvFstNonLetterzAy, RmvFstNonLetter(I)
Next
End Function

Function RmvSngQtezAy(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI RmvSngQtezAy, RmvSngQte(I)
Next
End Function

Function RplStarzAy(Ay, By) As String()
Dim I
For Each I In Itr(Ay)
    PushI RplStarzAy, Replace(I, By, "*")
Next
End Function

Function RplT1zAy(Ay, NewT1) As String()
RplT1zAy = AmAddPfx(AmzRmvT1(Ay), NewT1 & " ")
End Function

Function SyeKss(Sy$(), Kss$) As String()
If Kss = "" Then SyeKss = Sy: Exit Function
SyeKss = SyePred(Sy, PredzLikAy(SyzSS(Kss)))
End Function

Function SyeKssAy(Sy$(), KssAy$()) As String()
If Si(KssAy) = 0 Then SyeKssAy = Sy: Exit Function
SyeKssAy = SyePred(Sy, PredzKssAy(KssAy))
End Function

Function SyeLik(Sy$(), Lik$) As String()
SyeLik = SyePred(Sy, PredzLik(Lik))
End Function

Function SyeLikAy(Sy$(), LikAy$()) As String()
SyeLikAy = SyePred(Sy, PredzLikAy(LikAy))
End Function

Function SyeOneTermLin(Sy$()) As String()
SyeOneTermLin = SyePred(Sy, PredzIsOneTermLin)
End Function

Function SyePfx(Sy$(), ExlPfx$) As String()
SyePfx = SyePred(Sy, PredzPfx(ExlPfx))
End Function

Function SyePred(Sy$(), P As IPred) As String()
Dim I
For Each I In Itr(Sy)
    If Not P.Pred(I) Then
        PushI SyePred, I
    End If
Next
End Function

Function SyeT1Sy(Sy$(), ExlT1Sy$()) As String()
'Exclude those Lin in Array-Ay its T1 in ExlAmT10
If Si(ExlT1Sy) = 0 Then SyeT1Sy = Sy: Exit Function
SyeT1Sy = SyePred(Sy, PredzInT1Sy(ExlT1Sy))
End Function

Sub Z_AeAtCnt()
Dim Ay()
Ay = Array(1, 2, 3, 4, 5)
Ept = Array(1, 4, 5)
GoSub Tst
'
Exit Sub

Tst:
    Act = AeAtCnt(Ay, 1, 2)
    C
    Return
End Sub

Sub Z_AeEmpEleAtEnd()
Dim Ay: Ay = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AeEmpEleAtEnd(Ay)
Ass Si(Act) = 4
Ass Act(3) = 1
End Sub

Sub Z_AeEmpEleAtEnd1()
Dim Ay: Ay = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AeEmpEleAtEnd(Ay)
Ass Si(Act) = 4
Ass Act(3) = 1
End Sub

Sub Z_AeFei()
Dim Ay
Dim Fei1 As Fei
Dim Act
Ay = SplitSpc("a b c d e")
Fei1 = Fei(1, 2)
Act = AeFei(Ay, Fei1)
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Sub Z_AeFei1()
Dim Ay
Dim Act
Ay = SplitSpc("a b c d e")
Act = AeFei(Ay, Fei(1, 2))
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Sub Z_AeIxy()
Dim Ay(), Ixy
Ay = Array("a", "b", "c", "d", "e", "f")
Ixy = Array(1, 3)
Ept = Array("a", "c", "e", "f")
GoSub Tst
Exit Sub
Tst:
    Act = AeIxy(Ay, Ixy)
    C
    Return
End Sub

Sub Z_FmtDiKqCnt()
Dim Ay
GoSub Z
Exit Sub
Z:
    Ay = Array(1, 2, 2, 2, 3, "skldflskdfsdklf" & vbCrLf & "ksdlfj")
    Brw FmtDiKqCnt(Ay)
    
End Sub

Sub Z_SyeKss()
Dim Sy$(), Kss$
GoSub Z
GoSub T0
Exit Sub
T0:
    Sy = SyzSS("A B C CD E E1 E3")
    Kss = "C* E*"
    Ept = SyzSS("A B")
    GoTo Tst
Z:
    D SyeKss(SyzSS("A B C CD E E1 E3"), "C* E*")
    Return
Tst:
    Act = SyeKss(Sy, Kss)
    C
    Return
End Sub
