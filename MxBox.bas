Attribute VB_Name = "MxBox"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxBox."

Function BoxLines(Lines, Optional C$ = "*") As String()
BoxLines = BoxLy(SplitCrLf(Lines))
End Function
Function BoxLy(Ly$(), Optional C$ = "*") As String()
If Si(Ly) = 0 Then Exit Function
Dim W%, L$, I
W = WdtzAy(Ly)
L = Qte(Dup("-", W), "|-*-|")
PushI BoxLy, L
For Each I In Ly
    PushI BoxLy, "| " & AlignL(I, W) & " |"
Next
PushI BoxLy, L
End Function
Function BoxzS(S, Optional C$ = "*") As String()
Dim H$: H = Dup(C, Len(S) + 6)
PushI BoxzS, H
PushI BoxzS, C & C & " " & S & " " & C & C
PushI BoxzS, H
End Function
Function Box(V, Optional C$ = "*") As String()
If IsStr(V) Then
    If V = "" Then
        Exit Function
    End If
End If
Select Case True
Case IsLines(V): Box = BoxLines(V, C)
Case IsStr(V):   Box = BoxzS(V, C)
Case IsSy(V):    Box = BoxLy(CvSy(Sy), C)
Case IsArray(V): Box = BoxAy(V)
Case Else:       Box = BoxzS(V, C)
End Select
End Function

Function BoxFny(Fny$()) As String()
If Si(Fny) = 0 Then Exit Function
Const S$ = " | ", Q$ = "| * |"
Const LS$ = "-|-", LQ$ = "|-*-|"
Dim L$, H$, Ay$(), J%
    ReDim Ay(UB(Fny))
    For J = 0 To UB(Fny)
        Ay(J) = Dup("-", Len(Fny(J)))
    Next
L = Qte(Jn(Fny, S), Q)
H = Qte(Jn(Ay, LS), LQ)
BoxFny = Sy(H, L, H)
End Function

Function BoxAy(Ay) As String()
If Si(Ay) = 0 Then Exit Function
Dim W%: W = WdtzAy(Ay)
Dim H$: H = "|" & Dup("-", W + 2) & "|"
Push BoxAy, H
Dim I
For Each I In Ay
    Push BoxAy, "| " & AlignL(I, W) + " |"
Next
Push BoxAy, H
End Function
