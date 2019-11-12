Attribute VB_Name = "MxExpr"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxExpr."
Private Type LinRslt
    ExprLin As String
    OvrFlwTerm As String
    S As String
End Type
Private Type Term
    ExprTerm As String
    S As String
End Type

Function ExprLyzStr(Str, Optional MaxCdLinWdt% = 200) As String()
Dim L, Ay$(): Ay = SplitCrLf(Str)
Dim J&, Fst As Boolean
Erase XX
Fst = True
For Each L In Itr(Ay)
    If Fst Then
        Fst = False
    Else
        X J & ":" & Len(L) & ":" & L
    End If
'    Stop
    J = J + 1
'    PushIAy ExprLyzStr, ExprLyzLin(L, MaxCdLinWdt)
'    Stop
Next
Brw AddIxPfx(XX)
Stop
Erase XX
End Function
Function ExprLyzLin(Lin, W%) As String()
Dim J&
Dim S$: S = Lin
Dim CurLen&
Dim LasLen&: LasLen = Len(S)
Dim OvrFlwTerm$
While LasLen > 0
    DoEvents
    J = J + 1: If J > 10000 Then ThwLoopingTooMuch CSub
    Stop
    If J > 10 Then Stop
    With ShfLin(S, OvrFlwTerm, W)
        If .ExprLin = "" Then Exit Function
        PushI ExprLyzLin, .ExprLin
        S = .S
        OvrFlwTerm = .OvrFlwTerm
    End With
    CurLen = Len(S)
    If CurLen >= LasLen Then ThwNever CSub, "Str is not shifted by ShfLin"
    LasLen = CurLen
Wend
End Function

Function ShfLin(Str$, OvrFlwTerm$, W%) As LinRslt
Dim T$, OExprTermAy, TotW&
If OvrFlwTerm <> "" Then
    PushI OExprTermAy, OvrFlwTerm
    TotW = Len(OvrFlwTerm) + 3
End If
Dim S$: S = Str
Dim J&, OStr$, OExprTerm$

X:
ShfLin = LinRslt(ExprLin:=Jn(OExprTermAy, " & "), OvrFlwTerm:=OvrFlwTerm, S:=OStr)
End Function
Function Z_ShfTermzPrintable()
Dim S$: S = StrOfPjfP
Dim LAs&, Cur&, O$()
LAs = Len(S)
While Len(S) > 0
    PushI O, ShfTermzPrintable(S)
    Cur = Len(S)
    If Cur >= LAs Then Stop
    LAs = Cur
Wend
MsgBox Si(O)
Stop
Brw O
End Function
Function ShfTermzPrintable$(OStr$)
If OStr = "" Then Exit Function
Dim IsPrintable As Boolean
Dim J&
IsPrintable = IsAscPrintable(Asc(FstChr(OStr)))
For J = 2 To Len(OStr)
    If IsPrintable <> IsAscPrintable(Asc(Mid(OStr, J, 1))) Then
        ShfTermzPrintable = Left(OStr, J - 1)
        OStr = Mid(OStr, J)
        Exit Function
    End If
Next
ShfTermzPrintable = OStr
OStr = ""
End Function

'Fun=================================================
Function LinRslt(ExprLin, OvrFlwTerm$, S$) As LinRslt
With LinRslt
    .ExprLin = ExprLin
    .OvrFlwTerm = OvrFlwTerm
    .S = S
End With
End Function

Function ExprzQte$(BytAy() As Byte)
Dim O$(), I
For Each I In BytAy
    If I = vbDblQAsc Then PushI O, vb2DblQ Else PushI O, Chr(I)
Next
ExprzQte = QteDbl(Jn(O))
End Function

Function ExprzAndChr$(BytAy() As Byte)
Dim O$(), I
For Each I In BytAy
    PushI O, "Chr(" & I & ")"
Next
ExprzAndChr = Jn(O, " & ")
End Function

Function Term(ExprTerm$, S$) As Term
With Term
    .ExprTerm = ExprTerm
    .S = S
End With
End Function
Sub Z_ExprLyzStr()
Dim S$
GoSub ZZ1
GoSub ZZ2
GoSub T0
GoSub T1
Exit Sub
ZZ2:
    S = StrOfPjfP
    Brw ExprLyzStr(S)
    Return
ZZ1:
    S = StrOfPjfP
    Brw ExprLyzStr(S)
    Return
T0:
    S = "lksdjf lskdf dkf " & Chr(2) & Chr(11) & "ksldfj"
    Ept = Sy("")
    GoTo Tst
T1:
    GoTo Tst
Tst:
    Act = ExprLyzStr(S)
    D Act
    Stop
    C
    Return
End Sub

Sub Z_BrwRepeatedBytes()
BrwRepeatedBytes StrOfPjfP
End Sub

Function AscStr$(S)
Dim J&, O$()
For J = 1 To Len(S)
    PushI O, Asc(Mid(S, J, 1))
Next
AscStr = JnSpc(O)
End Function

Sub Z_BrkAyzPrintable1()
Dim T, O$(), J&
'For Each T In BrkAyzPrintable(JnCrLf(Srcp))
    J = J + 1
    Push O, FmtPrintableStr(T)
'Next
Brw AddIxPfx(O)
End Sub

Function FmtPrintableStr$(T)
Dim S$: S = PrintableSts(T)
Dim P$: P = S & " " & AlignL(Len(T), 8) & " : "
Select Case S
Case "Prt": FmtPrintableStr = P & T
Case "Non": FmtPrintableStr = P & AscStr(Left(T, 10))
Case "Mix": FmtPrintableStr = P & AscStr(Left(T, 10))
Case Else
    Stop
End Select
End Function
Sub Z_BrkAyzPrintable()
Brw BrkAyzPrintable(StrOfPjfP)
End Sub

Function BrkAyzRepeat(S) As String()
Dim L$: L = S
Dim T$, J&
While Len(L) > 0
    DoEvents
    T = ShfTermzRepeatedOrNot(L)
'    Debug.Print J, Len(L), Len(T), RepeatSts(T)
'    J = J + 1
    PushI BrkAyzRepeat, T
'    Stop
Wend
End Function
Function BrkAyzPrintable(S) As String()
Dim L$: L = S
#If True Then
    While Len(L) > 0
        Push BrkAyzPrintable, ShfTermzPrintable(L)
    Wend
#Else
    Dim T$, J&, I%
    While Len(L) > 0
        DoEvents
        T = ShfTermzPrintable(L)
        S = PrintableSts(T)
        Debug.Print J, Len(L), Len(T), S,
        If S = "NonPrintable" Then
            For I = 1 To Min(Len(T), 10)
                Debug.Print Asc(Mid(T, I, 1)); " ";
            Next
        End If
        Debug.Print
        
        J = J + 1
        PushI BrkAyzPrintable, T
    '    Stop
    Wend
#End If
End Function
Function PrintableSts$(T)
Dim IsPrintable As Boolean
IsPrintable = IsAscPrintable(Asc(FstChr(T)))
Dim J&
For J = 2 To Len(T)
    If IsPrintable <> IsAscPrintablezStrI(T, J) Then
        PrintableSts = "Mix"
        Stop
        Exit Function
    End If
Next
PrintableSts = IIf(IsPrintable, "Prt", "Non")
End Function

Function RepeatSts$(T)
'If Len(T) = 199 Then Stop
Select Case Len(T)
Case 0: RepeatSts = "ZeroByt": Exit Function
Case 1: RepeatSts = "OneByt":  Exit Function
Case Else
    Dim IsRepeat As Boolean, LAs$, C$, IsSam As Boolean
    LAs = SndChr(T)
    IsRepeat = FstChr(T) = LAs
    Dim J&
    For J = 3 To Len(T)
        C = Mid(T, J, 1)
        IsSam = C = LAs
        Select Case True
        Case IsRepeat And IsSam:
        Case IsRepeat: Stop: RepeatSts = "Mixed": Exit Function
        Case IsSam:    Stop: RepeatSts = "Mixed": Exit Function
        Case Else: LAs = C
        End Select
    Next
End Select
RepeatSts = IIf(IsRepeat, "Repated", "Dif")
End Function
Function ShfTermzRepeatedOrNot$(OStr$)
If OStr = "" Then Exit Function
Dim J&, C$, LAs$, IsSam As Boolean, IsRepeat As Boolean
LAs = SndChr(OStr)
IsRepeat = FstChr(OStr) = LAs
For J = 3 To Len(OStr)
    C = Mid(OStr, J, 1)
    IsSam = C = LAs
    Select Case True
    Case IsSam And IsRepeat
    Case IsSam
        ShfTermzRepeatedOrNot = Left(OStr, J - 2)
        OStr = Mid(OStr, J - 1)
        Exit Function
    Case IsRepeat
        ShfTermzRepeatedOrNot = Left(OStr, J - 1)
        OStr = Mid(OStr, J)
        Exit Function
    Case Else
        LAs = C
    End Select
Next
ShfTermzRepeatedOrNot = OStr
OStr = ""
End Function

Sub BrwRepeatedBytes(S)
Dim J&, b%, B1%, RepeatCnt&, L&
L = Len(S)
If L = 0 Then Exit Sub
b = Asc(FstChr(S)): RepeatCnt = 1
Erase XX
X FmtQQ("Len(?)", L)
For J = 2 To L
    B1 = Asc(Mid(S, J, 1))
    Select Case True
    Case b = B1:        RepeatCnt = RepeatCnt + 1
    Case Else
        If RepeatCnt > 1 Then
            X FmtQQ("Pos(?) Asc(?) RepeatCnt(?)", J, b, RepeatCnt)
            RepeatCnt = 1
        End If
        b = B1
    End Select
Next
Brw AddIxPfx(XX)
Erase XX
End Sub
