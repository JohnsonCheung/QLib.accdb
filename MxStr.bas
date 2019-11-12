Attribute VB_Name = "MxStr"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxStr."

Function SzIf$(IfTrue As Boolean, S)
If IfTrue Then SzIf = S
End Function

Function Pad0$(N&, NDig&)
Pad0 = Format(N, Dup("0", NDig))
End Function

Sub BrwStr(S, Optional Fnn$, Optional UseVc As Boolean)
Dim T$: T = TmpFt("BrwStr", Fnn$)
WrtStr S, T
BrwFt T, UseVc
End Sub

Sub VcStr(S, Optional Fnn$)
BrwStr S, Fnn, UseVc:=True
End Sub

Function StrDft$(S, Dft)
StrDft = IIf(S = "", Dft, S)
End Function

Function Dup$(S, N)
Dim O$, J&
For J = 0 To N - 1
    O = O & S
Next
Dup = O
End Function

Function HasSfxAs(S, AsSfxSy$()) As Boolean
Dim I, Sfx$
For Each I In Itr(AsSfxSy)
    Sfx = I
    If HasSfx(S, Sfx) Then HasSfxAs = True: Exit Function
Next
End Function

Function HasPfxAs(S, AsPfxSy$()) As Boolean
Dim I, Pfx$
For Each I In Itr(AsPfxSy)
    Pfx = I
    If HasPfx(S, Pfx) Then HasPfxAs = True: Exit Function
Next
End Function

Sub EdtStr(S, Ft)
WrtStr S, Ft, OvrWrt:=True
Brw Ft
End Sub

Function IsDigStr(S) As Boolean
Dim J&
For J = 1 To Len(S)
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then Exit Function
Next
IsDigStr = True
End Function

Function ApdStr$(S, Ft)
Dim Fno%: Fno = FnoA(Ft)
Print #Fno, S;
Close #Fno
ApdStr = Ft
End Function

Function WrtStr$(S, Ft, Optional OvrWrt As Boolean)
If OvrWrt Then DltFfnIf Ft
Dim Fno%: Fno = FnoO(Ft)
Print #Fno, S;
Close #Fno
WrtStr = Ft
End Function



Function Align$(V, W%)
Dim S: S = V
If IsStr(V) Then
    Align = AlignL(S, W)
Else
    Align = AlignR(S, W)
End If
End Function

Function AlignL$(S, W)
Dim L%: L = Len(S)
If L >= W Then
    AlignL = S
Else
    AlignL = S & Space(W - Len(S))
End If
End Function

Function AlignR$(S, W)
Dim L%: L = Len(S)
If W > L Then
    AlignR = Space(W - L) & S
Else
    AlignR = S
End If
End Function

Function AlignRzT1(Ly$()) As String()
Dim T1$(), Rst$()
AsgAmT1RstAy Ly, T1, Rst
T1 = AlignRzAy(T1)
Dim J&: For J = 0 To UB(T1)
    PushI AlignRzT1, T1(J) & " " & Rst(J)
Next
End Function

Function TrimWhite$(A)
TrimWhite = TrimWhiteL(TrimWhiteL(A))
End Function

Function TrimWhiteL$(A)
Dim J%
    For J = 1 To Len(A)
        If Not IsWhiteChr(Mid(A, J, 1)) Then Exit For
    Next
TrimWhiteL = Left(A, J)
End Function

Function TrimWhiteR$(S)
Dim J%
    Dim A$
    For J = Len(S) To 1 Step -1
        If Not IsWhiteChr(Mid(S, J, 1)) Then Exit For
    Next
    If J = 0 Then Exit Function
TrimWhiteR = Mid(S, J)
End Function
Function TabN$(N%)
TabN = Space(4 * N)
End Function


