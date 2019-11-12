Attribute VB_Name = "MxLTDH"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxLTDH."
Public Const FFoLTDH$ = "L T1 Dta IsHdr"
Type TDoLTDH: D As Drs: End Type
Sub Z_DyoLTDH()
Dim IndentSrc$()
GoSub Z
GoSub T0
Exit Sub
T0:
    Erase XX
    X "A Bc"
    X " 1"
    X " --"
    X " 2"
    X "A 2"
    IndentSrc = XX
    Erase XX
    Ept = Array( _
        Array(0&, "A", True, "Bc"), _
        Array(1&, "A", False, "1"), _
        Array(3&, "A", False, "2"), _
        Array(4&, "A", True, "2"))
    GoTo Tst
Tst:
    Act = DyoLTDH(IndentSrc)
    C
    Return
Z:
    Erase XX
    X "A Bc"
    X " 1"
    X " --"
    X " 2"
    X "A 2"
    IndentSrc = XX
    Erase XX
    DmpDy DyoLTDH(IndentSrc)
    Return
End Sub
Sub Z_IndentedLy()
Dim IndentSrc$(), K$
GoSub Z
GoSub T0
Exit Sub
T0:
    K = "A"
    Erase XX
    X "A Bc"
    X " 1"
    X " 2"
    X "A 2"
    IndentSrc = XX
    Erase XX
    Ept = Sy("1", "2")
    GoTo Tst
Tst:
    Act = IndentedLy(IndentSrc, K)
    C
    Return
Z:
    K = "A"
    Erase XX
    X "A Bc"
    X " 1"
    X " 2"
    X "A Bc"
    X " 1 2"
    X " 2 3"
    IndentSrc = XX
    Erase XX
    D IndentedLy(IndentSrc, K)
    Return
End Sub


Function IndentedLy(IndentSrc$(), Key$) As String()
Dim O$()
Dim L, Fnd As Boolean, IsNewSection As Boolean, IsFstChrSpc As Boolean, FstA%, Hit As Boolean
Const SpcAsc% = 32
For Each L In Itr(IndentSrc)
    If Fst2Chr(LTrim(L)) = "--" Then GoTo Nxt
    FstA = FstAsc(L)
    IsNewSection = IsAscUCas(FstA)
    If IsNewSection Then
        Hit = T1(L) = Key
    End If
    
    IsFstChrSpc = FstA = SpcAsc
    Select Case True
    Case IsNewSection And Not Fnd And Hit: Fnd = True
    Case IsNewSection And Fnd:             IndentedLy = O: Exit Function
    Case Fnd And IsFstChrSpc:              PushI O, Trim(L)
    End Select
Nxt:
Next
If Fnd Then IndentedLy = O: Exit Function
End Function

Function TDoLTDH(IndentSrc$()) As TDoLTDH
'Ret :TDrs-L-T1-Dta-IsHdr @@
TDoLTDH.D = DoLTDH(IndentSrc)
End Function

Function DoLTDH(IndentSrc$()) As Drs
DoLTDH = DrszFF(FFoLTDH, DyoLTDH(IndentSrc))
End Function

Private Function DyoLTDH(IndentedSrc$()) As Variant()
Dim L&, Dta$, T1$, IsHdr As Boolean, Lin
For Each Lin In Itr(IndentedSrc)
    L = L + 1
    Dim T$: T = LTrim(Lin)
    If Fst2Chr(T) = "--" Then GoTo X
    If T = "" Then GoTo X
    IsHdr = FstChr(Lin) <> " "
    If IsHdr Then
        Dta = RmvT1(Lin)
        T1 = T1zS(Lin)
    Else
        Dta = LTrim(Lin)
    End If
    PushI DyoLTDH, Array(L, T1, Dta, IsHdr)
X:
Next
End Function

