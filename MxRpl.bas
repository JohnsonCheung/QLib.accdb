Attribute VB_Name = "MxRpl"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxRpl."
':Q: :S #Str-With-QuestionMark#
Sub Z_RplBet()
Dim S$, Exp$, By$, S1$, S2$
S1 = "Data Source="
S2 = ";"
S = "aa;Data Source=???;klsdf"
By = "xx"
Exp = "aa;Data Source=xx;klsdf"
GoSub Tst
Exit Sub
Tst:
Dim Act$
Act = RplBet(S, By, S1, S2)
Debug.Assert Exp = Act
Return
End Sub

Sub Z_RplPfx()
Ass RplPfx("aaBB", "aa", "xx") = "xxBB"
End Sub
Function RmvCr$(S)
RmvCr = Replace(S, vbCr, "")
End Function

Function RplCr$(S)
RplCr = Replace(S, vbCr, " ")
End Function
Function RplCrLf$(S)
RplCrLf = RplLf(RplCr(S))
End Function
Function RplLf$(S)
RplLf = Replace(S, vbLf, " ")
End Function
Function RplVbl$(S)
RplVbl = RplVBar(S)
End Function
Function RplVBar$(S)
RplVBar = Replace(S, "|", vbCrLf)
End Function
Function RplBet$(S, By$, S1$, S2$)
Dim P1%, P2%, B$, C$
P1 = InStr(S, S1)
If P1 = 0 Then Stop
P2 = InStr(P1 + Len(S1), S, S2)
If P2 = 0 Then Stop
B = Left(S, P1 + Len(S1) - 1)
C = Mid(S, P2 + Len(S2) - 1)
RplBet = B & By & C
End Function

Function Rpl2DblQ$(S)
'Ret :S #Rptl-2DblQ-To-Blnk#
Rpl2DblQ = Replace(S, vb2DblQ, "")
End Function
Function RplDblSpc$(S)
Dim O$: O = Trim(S)
Dim J&
While HasSubStr(O, "  ")
    J = J + 1: If J > 10000 Then Stop
    O = Replace(O, "  ", " ")
Wend
RplDblSpc = O
End Function

Function RplFstChr$(S, By$)
RplFstChr = By & RmvFstChr(S)
End Function

Function RplPfx(S, Fm$, ToPfx$)
If HasPfx(S, Fm) Then
    RplPfx = ToPfx & RmvPfx(S, Fm)
Else
    RplPfx = S
End If
End Function

Sub Z_RplPfx2()
Ass RplPfx("aaBB", "aa", "xx") = "xxBB"
End Sub

Function RplPun$(S)
Dim O$(), J&, L&, C$
L = Len(S): If L = 0 Then Exit Function
ReDim O(L - 1)
For J = 1 To L
    C = Mid(S, J, 1)
    If IsPun(C) Then
        O(J - 1) = " "
    Else
        O(J - 1) = C
    End If
Next
RplPun = Join(O, "")
End Function

Function RplQ$(Q, By)
RplQ = Replace(Q, "?", By)
End Function

Function DyoAyAv(AyAv()) As Variant()
If Si(AyAv) = 0 Then Exit Function
Dim UAy%: UAy = UB(AyAv)
Dim URec&: URec = UB(AyAv(0))
Dim ODy(): ReDim ODy(URec)
Dim R&: For R = 0 To URec
    Dim Dr(): ReDim Dr(UAy)
    Dim C%: For C = 0 To UAy
        Dr(C) = AyAv(C)(R)
    Next
    ODy(R) = Dr
Next
DyoAyAv = ODy
End Function
Function SyzMacro(RplMacro$, ParamArray ByAyAp()) As String()
Dim AyAv(): AyAv = ByAyAp
SyzMacro = SyzMacroDy(RplMacro, DyoAyAv(AyAv))
End Function

Function SyzMacroDy(RplMacro$, ByDy()) As String()
If Si(ByDy) = 0 Then Exit Function
Dim M$():     M = NyzMacro(RplMacro, InclBkt:=True)
Dim URec&: URec = UB(ByDy)
Dim UFld%: UFld = UB(ByDy(0))

If UB(M) <> UFld Then Thw CSub, "UFld should = UB(MacroNy)", "UFld UB(MacroNy)", UFld, UB(M)

'-- O --
Dim O$(): ReDim O(URec)
Dim J&, Dr: For Each Dr In ByDy
    O(J) = SzMacro(RplMacro, M, Dr)
    J = J + 1
Next
SyzMacroDy = O
End Function

Function SzMacro$(RplMacro$, OfMacroNy$(), ByDr)
Dim O$: O = RplMacro
Dim V, J%: For Each V In ByDr
    O = Replace(O, OfMacroNy(J), V)
    J = J + 1
Next
SzMacro = O
End Function

Function SyzQAy(Q, ByAy) As String()
Dim By: For Each By In Itr(ByAy)
    PushI SyzQAy, RplQ(Q, By)
Next
End Function

Sub Z_RplBet3()
Dim S$, Exp$, By$, S1$, S2$
S1 = "Data Source="
S2 = ";"
S = "aa;Data Source=???;klsdf"
By = "xx"
Exp = "aa;Data Source=xx;klsdf"
GoSub Tst
Exit Sub
Tst:
Dim Act$
Act = RplBet(S, By, S1, S2)
Debug.Assert Exp = Act
Return
End Sub


