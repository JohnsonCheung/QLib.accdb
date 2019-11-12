Attribute VB_Name = "MxAyOp"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAyOp."
Enum EmIxCol
EiBeg1
EiBeg0
EiBegI
EiNoIx
End Enum

Function AyMinusAp(Ay, ParamArray Ap())
Dim IAy, O
O = Ay
For Each IAy In Ap
    O = AyMinus(Ay, IAy)
    If Si(O) = 0 Then AyMinusAp = O: Exit Function
Next
AyMinusAp = O
End Function

Function CvVy(Vy)
Const CSub$ = CMod & "CvVy"
Select Case True
Case IsStr(Vy): CvVy = SyzSS(CStr(Vy))
Case IsArray(Vy): CvVy = Vy
Case Else: Thw CSub, "VyzDicKK should either be string or array", "Vy-TypeName Vy", TypeName(Vy), Vy
End Select
End Function

Function SyShow(XX$, Sy$()) As String()
Dim O$()
Select Case Si(Sy)
Case 0
    Push O, XX & "()"
Case 1
    Push O, XX & "(" & Sy(0) & ")"
Case Else
    Push O, XX & "("
    PushAy O, Sy
    Push O, XX & ")"
End Select
SyShow = O
End Function



Sub ThwIf_Dup(Ay, Fun$)
' If there are 2 ele with same string (IgnCas), throw error
Dim Dup$()
    Dup = AwDup(Ay)
If Si(Dup) = 0 Then Exit Sub
Thw Fun, "There are dup in array", "Dup Ay", Dup, Ay
End Sub


Function OffsetzEmBeg(B As EmIxCol, Optional FmI&) As Byte
Select Case True
Case B = EiBeg0: OffsetzEmBeg = 0
Case B = EiBeg1: OffsetzEmBeg = 1
Case B = EiBegI: OffsetzEmBeg = FmI
Case Else: Thw CSub, "EmIxCol value error", "EmIxCol", B
End Select
End Function

Function AddIxPfx(Ay, Optional B As EmIxCol = EiBeg0, Optional FmI&) As String()
If B = EiNoIx Then AddIxPfx = CvSy(Ay): Exit Function
Dim L, J&, N%
J = OffsetzEmBeg(B, FmI)
N = Len(CStr(UB(Ay) + J))
For Each L In Itr(Ay)
    PushI AddIxPfx, AlignR(J, N) & ": " & L
    J = J + 1
Next
End Function

Function TabNmV$(Nm$, V, Optional NTab% = 1)
TabNmV = TabN(NTab) & Nm & V
End Function
Function TabNmLy(Nm$, Ly$(), Optional NTab% = 1, Optional Beg As EmIxCol = EiNoIx) As String()
Stop
If Si(Ly) = 0 Then
    PushI TabNmLy, TabN(NTab) & Nm
    Exit Function
End If
Dim Ly1$(), L0$, S$, J&
Ly1 = AddIxPfx(Ly, Beg)
PushI TabNmLy, TabN(NTab) & Nm & Ly1(0)
'
S = TabN(NTab) & Space(Len(Nm))
For J = 1 To UB(Ly1)
    PushI TabNmLy, S & Ly1(J)
Next
End Function

Sub SplitAsgzAyPred(Ay, P As IPred, OTrueAy, OFalseAy)
Dim V: For Each V In Itr(Ay)
    If P.Pred(V) Then
        Push OTrueAy, V
    Else
        Push OFalseAy, V
    End If
Next
End Sub

Function LookupT1$(Itm, TkssAy$())
Dim L$, I, Kss$, T1$
For Each I In TkssAy
    L = I
    AsgTRst L, T1, Kss
    If HitKss(Itm, Kss) Then LookupT1 = T1: Exit Function
Next
End Function

Function AddSS(Sy$(), SS$) As String()
AddSS = SyzAp(Sy, SyzSS(SS))
End Function


