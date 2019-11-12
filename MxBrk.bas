Attribute VB_Name = "MxBrk"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxBrk."
':Dn:  :S #Dot-Nm# ! one-or-more-nm sep by Dot

Sub AsgBrkBet(L$, A$, B$, O1, O2, O3)
AsgS3 BrkBet(L, A, B), O1, O2, O3
End Sub

Sub AsgS3(A As S3, O1, O2, O3)
O1 = A.A
O2 = A.B
O3 = A.C
End Sub

Function BrkBet(L$, A$, B$) As S3
If L = "" Then Exit Function
Dim P1%, P2%, O As S3, LA%, LB%
LA = Len(A)
LB = Len(B)
P1 = InStr(L, A)
P2 = InStr(L, B)
Select Case True
Case P1 <> 0 And P2 <> 0 And P1 > P2:       Stop
Case P1 = 0 And P2 = 0: O.A = Trim(L)
Case P1 = 0:            O.A = Trim(Left(L, P2 - 1)): O.C = Trim(Mid(L, P2 + LB))
Case P2 = 0:            O.A = Trim(Left(L, P1 - 1)): O.B = Trim(Mid(L, P1 + LA))
Case Else:              O.A = Trim(Left(L, P1 - 1)): O.B = Trim(Mid(L, P1 + LA, P2 - P1 + LA - 2)): O.C = Trim(Mid(L, P2 + Len(LB)))
End Select
BrkBet = O
End Function

Sub AsgBrkSpc(S, OA$, OB$, Optional NoTrim As Boolean)
AsgS12 BrkSpc(S), OA, OB
End Sub

Sub AsgBrk1Dot(S, OA$, OB$, Optional NoTrim As Boolean)
AsgS12 Brk1Dot(S), OA, OB
End Sub

Sub AsgBrkDot(S, OA$, OB$, Optional NoTrim As Boolean)
AsgS12 BrkDot(S), OA, OB
End Sub

Function Brk1Dot(S, Optional NoTrim As Boolean) As S12
Brk1Dot = Brk1(S, ".", NoTrim)
End Function

Function Brk2Dot(S, Optional NoTrim As Boolean) As S12
Brk2Dot = Brk2(S, ".", NoTrim)
End Function

Function BrkDot(S, Optional NoTrim As Boolean) As S12
BrkDot = Brk(S, ".", NoTrim)
End Function

Function BrkTRst(S) As S12
BrkTRst = Brk1Spc(LTrim(S))
End Function

Function BrkSpc(S) As S12
BrkSpc = Brk(S, " ")
End Function

Function Brk1Spc(S) As S12
Brk1Spc = Brk1(S, " ")
End Function

Sub AsgDn1(Dn1, OA$, OB$)
AsgBrkDot Dn1, OA, OB
End Sub

Function Brk(S, Sep$, Optional NoTrim As Boolean) As S12
Const CSub$ = CMod & "Brk"
Dim P&: P = InStr(S, Sep)
If P = 0 Then Thw CSub, "{S} does not contains {Sep}", "S Sep", S, Sep
Brk = BrkAtSep(S, P, Sep, NoTrim)
End Function

Function Brk1(S, Sep$, Optional NoTrim As Boolean) As S12
Dim P&: P = InStr(S, Sep)
If P = 0 Then Brk1 = S12(S, "", NoTrim): Exit Function
Brk1 = Brk1At(S, P, Sep, NoTrim)
End Function

Sub AsgBrk1(S, Sep$, Optional O1, Optional O2, Optional NoTrim As Boolean)
AsgS12 Brk1(S, Sep, NoTrim), O1, O2
End Sub

Function Brk1Rev(S, Sep, Optional NoTrim As Boolean) As S12
Dim P&: P = InStrRev(S, Sep)
If P = 0 Then Brk1Rev = S12(S, "", NoTrim): Exit Function
Brk1Rev = Brk1At(S, P, Sep, NoTrim)
End Function

Function Brk2(S, Sep, Optional NoTrim As Boolean) As S12
Dim P&: P = InStr(S, Sep)
Brk2 = Brk2__(S, P, Sep, NoTrim)
End Function

Function Brk2__(S, P&, Sep, NoTrim As Boolean) As S12
If P = 0 Then
    If NoTrim Then
        Brk2__ = S12("", S)
    Else
        Brk2__ = S12("", Trim(S))
    End If
    Exit Function
End If
Brk2__ = Brk1At(S, P, Sep, NoTrim)
End Function

Sub AsgBrk2(S, Sep$, O1, O2, Optional NoTrim As Boolean)
AsgS12 Brk2(S, Sep, NoTrim), O1, O2
End Sub

Function Brk2Rev(S, Sep, Optional NoTrim As Boolean) As S12
Dim P&: P = InStrRev(S, Sep)
Brk2Rev = Brk2__(S, P, Sep, NoTrim)
End Function

Sub AsgBrk(S, Sep$, Optional O1, Optional O2, Optional NoTrim As Boolean)
AsgBrkAt S, InStr(S, Sep), Sep, O1, O2, NoTrim
End Sub

Function BrkAtSep(S, P&, Sep, NoTrim As Boolean) As S12
Dim S1$, S2$
S1 = Left(S, P - 1)
S2 = Mid(S, P + Len(Sep))
BrkAtSep = S12(S1, S2, NoTrim)
End Function

Function Brk1At(S, P&, Sep, NoTrim As Boolean) As S12
If P = 0 Then
    Brk1At = S12(S, "", NoTrim)
Else
    Brk1At = BrkAtSep(S, P, Sep, NoTrim)
End If
End Function

Sub AsgBrkAt(S, At&, Sep$, O1, O2, Optional NoTrim As Boolean)
Const CSub$ = CMod & "AsgBrkAt"
If At = 0 Then
    Thw CSub, "String does not have Sep", "Str Sep At NoTrim", S, Sep, At, NoTrim
    Exit Sub
End If
O1 = Left(S, At - 1)
O2 = Mid(S, At + Len(Sep))
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
End Sub

Function BrkBoth(S, Sep, Optional NoTrim As Boolean) As S12
Dim P&: P = InStr(S, Sep)
If P = 0 Then
    BrkBoth = S12(S, S, NoTrim)
    Exit Function
End If
BrkBoth = Brk1At(S, P, Sep, NoTrim)
End Function

Sub AsgBrkQte(QteStr$, O1$, O2$)
AsgS12 BrkQte(QteStr), O1, O2
End Sub

Function BrkRev(S, Sep, Optional NoTrim As Boolean) As S12
Dim P&: P = InStrRev(S, Sep)
If P = 0 Then Err.Raise "BrkRev: Str[" & S & "] does not contains Sep[" & Sep & "]"
BrkRev = Brk1At(S, P, Len(Sep), NoTrim)
End Function

Sub AsgBrk1At(S, At&, Sep$, O1, O2, Optional NoTrim As Boolean)
If At = 0 Then
    O1 = S
    O2 = ""
    Exit Sub
End If
O1 = Left(S, At - 1)
O2 = Mid(S, At + Len(Sep))
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
End Sub

Sub Z_Brk1Rev()
Dim S1$, S2$, ExpS1$, ExpS2$, S
S = "aa --- bb --- cc"
ExpS1 = "aa --- bb"
ExpS2 = "cc"
With Brk1Rev(S, "---")
    S1 = .S1
    S2 = .S2
End With
Ass S1 = ExpS1
Ass S2 = ExpS2
End Sub

Sub Z_Brk1Rev1()
Dim S1$, S2$, ExpS1$, ExpS2$, S
S = "aa --- bb --- cc"
ExpS1 = "aa --- bb"
ExpS2 = "cc"
With Brk1Rev(S, "---")
    S1 = .S1
    S2 = .S2
End With
Ass S1 = ExpS1
Ass S2 = ExpS2
End Sub

