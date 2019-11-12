Attribute VB_Name = "MxEsc"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxEsc."
Function Hex2zAsc$(Asc%)
If Asc < 16 Then
    Hex2zAsc = "0" & Hex(Asc)
Else
    Hex2zAsc = Hex(Asc)
End If
End Function
Function Hex2$(C$)
If Len(C) <> 1 Then Thw CSub, "C should have len=1", "C Len", C, Len(C)
Hex2 = Hex2zAsc(Asc(C))
End Function
Function PerHex2$(C$)
PerHex2 = "%" & Hex2(C)
End Function
Function EscChr$(S, C$)
EscChr = EscAsc(S, Asc(C))
End Function
Function EscAsc$(S, A%) 'Escaping the AscChr-A% in S$ as %HH
EscAsc = Replace(S, Chr(A), "%" & Hex2zAsc(A))
End Function
Function EscSqBkt$(S)
EscSqBkt = EscChrLis(S, "[]")
End Function
Function EscChrLis$(S, ChrLis$)
Dim O$, J: O = S
For J = 1 To Len(ChrLis)
    O = EscChr(O, Mid(ChrLis, J, 1))
Next
EscChrLis = O
End Function

':SlashC$ = "It is 1 chr.  It will combine with sfx-\.  Eg.  SlashC = 'r', it measns it will be '\r'"
Function SlashCr$(S) 'Escapeing vbCr in S.
SlashCr = Slash$(S, vbCr, "r")
End Function
Function EscOpnSqBkt$(S)
EscOpnSqBkt = EscChr(S, "[")
End Function
Function UnSlashCrLfTab$(S)
UnSlashCrLfTab = UnSlashCr(UnSlashLf(UnSlashTab(S)))
End Function
Function EscClsSqBkt$(S)
EscClsSqBkt = EscChr(S, "]")
End Function
Function UnSlashCrLf$(S)
UnSlashCrLf = UnSlashCr(UnSlashLf(S))
End Function
Function SlashAsc$(S, Asc%, C$)
SlashAsc = Slash(S, Chr(Asc), C)
End Function
Function SlashCrLf$(S)
SlashCrLf = SlashLf(SlashCr(S))
End Function
Function SlashCrLfTab$(S)
SlashCrLfTab = SlashTab(SlashLf(SlashCr(S)))
End Function
Function SlashLf$(S)
SlashLf = SlashAsc(S, 10, "n")
End Function
Function UnSlashChr$(S, C$, SlashC$)
UnSlashChr = Replace(S, "\" & SlashC, C)
End Function

Function Slash$(S, C$, SlashC$) 'Escaping C$ in S by \SlashC$.  Eg C$ is vbCr and SlashC is r.
If InStr(S, "\" & SlashC) > 0 Then
    Debug.Print FmtQQ("SlashChr: Given S has \?, when UnSlash, it will not match", SlashC)
    Debug.Print vbTab; QteSq(S)
End If
Slash = Replace(S, C, "\" & SlashC)
End Function

Function UnEscBackSlash$(S)
Stop
'UnEscBackSlash = UnEscChr(S, "\")
End Function
Function EscBackSlash$(S)
EscBackSlash = EscChr(S, "\")
End Function

Function EscCr$(S)
EscCr = Esc(S, vbCr)
End Function

Function EscCrLf$(S)
EscCrLf = EscCr(EscLf(S))
End Function

Function EscLf$(S)
EscLf = EscChr(S, vbLf)
End Function

Function SlashTab$(S)
SlashTab = SlashChr(S, vbTab, "t")
End Function
Function SlashChr$(S, C$, SlashC$)
SlashChr = Replace(S, C, "\" & SlashC)
End Function
Function Esc$(S, C$)
Esc = EscAsc(S, Asc(C))
End Function

Function UnSlashCr$(S)
UnSlashCr = Replace(S, "\r", vbCr)
End Function

Function TidleSpc$(S)
If InStr(S, "~") Then
    Debug.Print "TidleSpc: Given-S has space"
    Debug.Print vbTab; "[" & S & "]"
End If
TidleSpc = Replace(S, " ", "~")
End Function
Function UnTidleSpc$(S)
UnTidleSpc = Replace(S, "~", " ")
End Function

Function UnSlashTab(S)
UnSlashTab = Replace(S, "\t", vbTab)
End Function

Function UnSlashBackSlash$(S)
UnSlashBackSlash = Replace(S, "\\", "\")
End Function

Function UnEscCr$(S)
UnEscCr = Replace(S, "\r", vbCr)
End Function

Function UnEscCrLf$(S)
UnEscCrLf = UnEscLf(UnEscCr(S))
End Function
Function UnEscLf$(S)

End Function

Function UnSlashLf$(S)
UnSlashLf = Replace(S, "\n", vbLf)
End Function

Function UnEscSpc$(S)
UnEscSpc = Replace(S, "~", " ")
End Function

Function UnEscSqBkt$(S)
UnEscSqBkt = Replace(S, Replace(S, "\o", "["), "\c", "]")
End Function

Function UnEscTab(S)
UnEscTab = Replace(S, "\t", vbTab)
End Function
