Attribute VB_Name = "MxWrp"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxWrp."

Function WrpLines$(Lines$, Optional Wdt% = 80)
WrpLines = Lines: Exit Function
WrpLines = JnCrLf(WrpLy(SplitCrLf(Lines), Wdt))
End Function

Sub Z_WrpLy()
Dim Ly$(), Wdt%
GoSub T1
Exit Sub
T1:
    Ly = Sy("a b c d")
    Wdt = 80
    Ept = Sy("a b c d")
    GoTo Tst
Tst:
    Act = WrpLy(Ly, Wdt)
    C
    Return
End Sub


Sub Z_WrpLines()
Dim A$, W%
A = "lksjf lksdj flksdjf lskdjf lskdjf lksdjf lksdjf klsdjf klj skldfj lskdjf klsdjf klsdfj klsdfj lskdfj  sdlkfj lsdkfj lsdkjf klsdfj lskdjf lskdjf kldsfj lskdjf sdklf sdklfj dsfj "
W = 80
Ept = Sy("lksjf lksdj flksdjf lskdjf lskdjf lksdjf lksdjf klsdjf klj skldfj lskdjf klsdjf ", _
"klsdfj klsdfj lskdfj  sdlkfj lsdkfj lsdkjf klsdfj lskdjf lskdjf kldsfj lskdjf", _
"sdklf sdklfj dsfj ")
GoSub Tst
Exit Sub
Tst:
    Act = WrpLines(A, W)
    C
    Return
End Sub

Function WrpLy(Ly$(), Optional Wdt% = 80) As String()
Dim W%, Lin, I
W = EnsBet(Wdt, 10, 200)
For Each I In Itr(Ly)
    Lin = I
    PushIAy WrpLy, WrpLin(Lin, W)
Next
End Function

Function ShfWrpgLin(OLin$, W%)
If OLin = "" Then Exit Function
Dim O$, OL$, F$
O = Left(OLin, W)
OL = Mid(OLin, W + 1)
F = FstChr(OL)
Select Case True
Case OL = "" Or F = " "
Case Else:
    Dim P%: P = InStrRev(O, " ")
    If P <> 0 Then
        O = Left(O, P)
        OL = Mid(O, P + 1) & OL
    End If
End Select
ShfWrpgLin = Trim(O)
OLin = OL
End Function

Function WrpLin(Lin, W%) As String()
If Len(Lin) > W Then WrpLin = Sy(Lin): Exit Function
Dim L$: L = RTrim(Lin)
Dim J%: While L <> ""
    J = J + 1: If J > 1000 Then ThwLoopingTooMuch CSub
    PushI WrpLin, ShfWrpgLin(L, W)
Wend
End Function
