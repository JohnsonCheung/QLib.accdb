Attribute VB_Name = "MxShfStr"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxShfStr."
Function ShfDotSeg$(OLin$)
ShfDotSeg = ShfBef(OLin, ".")
End Function
Function ShfBktStr$(OLin$)
ShfBktStr = BetBkt(OLin)
OLin = AftBkt(OLin$)
End Function

Function RmvLasChrzzLis$(S, ChrLis$) ' Rmv fst chr if it is in ChrLis
If HasSubStr(ChrLis, LasChr(S)) Then
    RmvLasChrzzLis = RmvLasChr(S)
Else
    RmvLasChrzzLis = S
End If
End Function

Function TakChr$(S, ChrLis$) ' Ret fst chr if it is in ChrLis
If HasSubStr(ChrLis, FstChr(S)) Then TakChr = FstChr(S)
End Function

Function RmvFstChrzzLis$(S, ChrLis$)
If HasSubStr(ChrLis, FstChr(S)) Then
    RmvFstChrzzLis = RmvFstChr(S)
Else
    RmvFstChrzzLis = S
End If
End Function
Function ShfChr$(OLin$, ChrList$)
Dim C$: C = TakChr(OLin, ChrList)
If C = "" Then Exit Function
ShfChr = C
OLin = Mid(OLin, 2)
End Function
Function ShfTermX(OLin$, TermX$) As Boolean
If T1(OLin) <> TermX Then Exit Function
ShfTermX = True
OLin = RmvT1(OLin)
End Function
Function ShfEq(OLin$) As Boolean
ShfEq = ShfTermX(OLin, "=")
End Function

Function ShfTy(OLin$) As Boolean
ShfTy = ShfTermX(OLin, "Ty")
End Function

Function ShfBkt(OLin$) As Boolean
ShfBkt = ShfPfx(OLin, "()")
End Function

Function ShfPfx(OLin$, Pfx$) As Boolean
If HasPfx(OLin, Pfx) Then
    OLin = RmvPfx(OLin, Pfx)
    ShfPfx = True
End If
End Function
Function ShfSfx(OLin$, Sfx$) As Boolean
If HasSfx(OLin, Sfx) Then
    OLin = RmvSfx(OLin, Sfx)
    ShfSfx = True
End If
End Function
Function ShfPfxAy$(OLin$, PfxAy$())
Dim O$: O = PfxzAy(OLin, PfxAy): If O = "" Then Exit Function
ShfPfxAy = O
OLin = RmvPfx(OLin, O)
End Function

Function ShfPfxAyS$(OLin$, PfxAy$())
Dim O$: O = PfxzAyS(OLin, PfxAy): If O = "" Then Exit Function
ShfPfxAyS = O
OLin = RmvPfxSpc(OLin, O)
End Function

Function ShfPfxSpc(OLin$, Pfx$) As Boolean
If HitPfxSpc(OLin, Pfx) Then
    OLin = Mid(OLin, Len(Pfx) + 2)
    ShfPfxSpc = True
End If
End Function

Sub Z_ShfBktStr()
Dim A$, Ept1$
A$ = "(O$()) As X": Ept = "O$()": Ept1 = " As X": GoSub Tst
Exit Sub
Tst:
    Act = ShfBktStr(A)
    C
    Ass A = Ept1
    Return
End Sub

Property Get Z_ShfPfx()
Dim O$: O = "AA{|}BB "
Ass ShfPfx(O, "{|}") = "AA"
Ass O = "BB "
End Property



