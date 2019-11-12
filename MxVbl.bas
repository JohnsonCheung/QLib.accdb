Attribute VB_Name = "MxVbl"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbl."

Function VblzLines$(Lines$)
VblzLines = Replace(RmvCr(Lines), vbLf, "|")
End Function

Function DiczVbl(Vbl$, Optional JnSep$ = vbCrLf) As Dictionary
Set DiczVbl = Dic(SplitVBar(Vbl), JnSep)
End Function


Function SyzVbl(Vbl) As String()
SyzVbl = SplitVBar(Vbl)
End Function
Function ItrzVbl(Vbl)
ItrzVbl = Itr(SyzVbl(Vbl))
End Function

Function LineszVbl$(Vbl)
LineszVbl = Replace(Vbl, "|", vbCrLf)
End Function

Function IsVbl(S) As Boolean
Select Case True
Case Not IsStr(S)
Case HasSubStr(S, vbCr)
Case HasSubStr(S, vbLf)
Case Else: IsVbl = True
End Select
End Function

Function IsVblAy(VblAy$()) As Boolean
Dim Vbl: For Each Vbl In Itr(VblAy)
    If Not IsVbl(Vbl) Then Exit Function
Next
IsVblAy = True
End Function

Function IsVdtVbl(Vbl$) As Boolean
If HasSubStr(Vbl, vbCr) Then Exit Function
If HasSubStr(Vbl, vbLf) Then Exit Function
IsVdtVbl = True
End Function

Function DrszTRst(FF$, TRstLy$()) As Drs
DrszTRst = DrszFF(FF, DyoTRst(TRstLy))
End Function
Function DyoTRst(TRstLy$()) As Variant()
Dim L: For Each L In Itr(TRstLy)
    PushI DyoTRst, SyzTRst(L)
Next
End Function
Function DyoTLiny(TLiny$()) As Variant()
Dim I
For Each I In Itr(TLiny)
    PushI DyoTLiny, TermAy(I)
Next
End Function

Function DyoVblLy(A$()) As Variant()
Dim I
For Each I In Itr(A)
    PushI DyoVblLy, AmTrim(SplitVBar(I))
Next
End Function
Function DyoSSVbl(SSVbl$) As Variant()
Dim SS: For Each SS In Itr(SplitVBar(SSVbl))
    PushI DyoSSVbl, SyzSS(SS)
Next
End Function

Sub Z_DyoVblLy()
Dim VblLy$()
GoSub T1
Exit Sub
T0:
    Erase XX
    X "1 | 2 | 3"
    X "4 | 5 6 |"
    X "| 7 | 8 | 9 | 10 | 11 |"
    X "12"
    VblLy = XX
    Ept = Array(SyzSS("1 2 3"), Sy("4", "5 6", ""), Sy("", "7", "8", "9", "10", "11", ""), Sy("12"))
    GoTo Tst
Exit Sub
T1:
    Erase XX
    X "|lskdf|sdlf|lsdkf"
    X "|lsdf|"
    X "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
    X "|sdf"
    VblLy = XX
    Ept = ""
    GoTo Tst
Tst:
    Act = DyoVblLy(VblLy)
    DmpDy CvAv(Act)
'    C
    Return
End Sub

