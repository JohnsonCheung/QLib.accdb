Attribute VB_Name = "MxLines"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxLines."
Function WdtzLines%(Lines)
WdtzLines = WdtzAy(SplitCrLf(Lines))
End Function

Function WdtzLinesAy%(LinesAy)
Dim O%, Lines
For Each Lines In Itr(LinesAy)
    O = Max(O, WdtzLines(Lines))
Next
WdtzLinesAy = O
End Function

Sub VcLinesAy(LinesAy)
Vc FmtLinesAy(LinesAy)
End Sub

Sub BrwLinesAy(LinesAy)
B FmtLinesAy(LinesAy)
End Sub

Function MrkLines$(Lines, Lno&)
Dim Ly$(): Ly = SplitCrLf(Lines)

End Function

Sub Z_FmtLinesAy()
Dim LinesAy
GoSub Z
Exit Sub
Z:
    BrwLinesAy Y_LinesAy
    Return
End Sub

Function Y_LinesAy() As String()
With New Bfr
.X RplVbl("sdklf|lskdjflsdf|lsdkjflsdkfjsdflsdf|skldfjdsf|dklfsjdlksjfsldkf")
.X RplVbl("sdklf2-49230  sdfjldf|lskdjflsdf|lsdkjflsdkfjsdflsdf|skldfjdsf|dklfsjdlksjfsldkf")
.X RplVbl("sdsdfklf2-49230  sdfjldf|lskdjflsdf|lsdkjflsdkfjsdflsdf|skldfjdsf|dklfsjdlksjfsldkf")
Y_LinesAy = .Ly
End With
End Function

Function FmtLinesAy(LinesAy, Optional B As EmIxCol = EiBeg0) As String()
If Si(LinesAy) = 0 Then Exit Function
Dim W%: W = WdtzLinesAy(LinesAy)
Dim LinzSep: LinzSep = Qte(Dup("-", W + 2), "|")
Dim Lines
PushI FmtLinesAy, LinzSep
For Each Lines In Itr(LinesAy)
    PushIAy FmtLinesAy, AddIxPfxzLineszW(Lines, W, B)
    PushI FmtLinesAy, LinzSep
Next
End Function
Function AddIxPfxzLineszW(Lines, W%, Optional B As EmIxCol = EiBeg0) As String()
Dim L
For Each L In Itr(SplitCrLf(Lines))
    PushI AddIxPfxzLineszW, "| " & AlignL(L, W) & " |"
Next
End Function


Function EnsBet%(I%, A%, B%)
Select Case True
Case I < A: EnsBet = A
Case I > B: EnsBet = B
Case Else: EnsBet = I
End Select
End Function


Sub Z_LineszRTrim()
Dim Lines$: Lines = LineszVbl("lksdf|lsdfj|||")
Dim Act$: Act = LineszRTrim(Lines)
Debug.Print Act & "<"
Stop
End Sub

Sub Z_LineszLasN()
Dim Ay$(), A$, J%
For J = 0 To 9
Push Ay, "Line " & J
Next
A = Join(Ay, vbCrLf)
Debug.Print LineszLasN(A, 3)
End Sub

Function FstLin(Lines$)
FstLin = BefOrAll(Lines, vbCrLf)
End Function

Function LinesApp$(A, L)
If A = "" Then LinesApp = L: Exit Function
LinesApp = A & vbCrLf & L
End Function

Sub Z_LineszRTrim1()
Dim Lines$: Lines = LineszVbl("lksdf|lsdfj|||")
Dim Act$: Act = LineszRTrim(Lines)
Debug.Print Act & "<"
Stop
End Sub

Function LineszLasN$(Lines$, N%)
LineszLasN = JnCrLf(AwLasN(SplitCrLf(Lines), N))
End Function

Function LinCnt&(Lines)
LinCnt = Si(SplitCrLf(Lines))
End Function

Function MaxLinCnt&(LinesAy$())
Dim O&, Lines: For Each Lines In Itr(LinesAy)
    O = Max(O, LinCnt(Lines))
Next
MaxLinCnt = O
End Function

Function KE24Rs() As Recordset
Set KE24Rs = SampDb.TableDefs("KE24").OpenRecordset
End Function

Sub Z_DrszRs()
BrwDrs DrszRs(KE24Rs)
End Sub

Function SqhzLines(Lines$) As Variant()
SqhzLines = SqH(SplitCrLf(Lines))
End Function

Function SqvzLines(Lines$) As Variant()
SqvzLines = SqV(SplitCrLf(Lines))
End Function

Function LineszRTrim$(Lines)
Dim At&
For At = Len(Lines) To 1 Step -1
    If Not IsStrAtSpcCrLf(Lines, At) Then LineszRTrim = Left(Lines, At): Exit Function
Next
End Function

Function LasLin(Lines$)
LasLin = LasEle(SplitCrLf(Lines))
End Function

Function LineszAlign$(Lines$, W%)
Dim LAs$: LAs = LasLin(Lines)
Dim N%: N = W - Len(LAs)
If N > 0 Then
    LineszAlign = Lines & Space(N)
Else
    Warn CSub, "W is too small", "Lines.LasLin W", LAs, W
    LineszAlign = Lines
End If
End Function

Function NLines&(Lines)
Dim R As RegExp: Set R = Rx("\n", MultiLine:=True, IsGlobal:=True)
NLines = Mch(R, Lines).Count + 1
End Function

Function LineszV$(V)
LineszV = JnCrLf(FmtV(V))
End Function

Function ColRgAyzLinesDr(LinesDr$()) As Variant()
Dim Lines: For Each Lines In LinesDr
    PushI ColRgAyzLinesDr, SplitCrLf(Lines)
Next
End Function

Function SqzLinesDr(LinesDr$()) As Variant()
Dim NRow%: NRow = MaxLinCnt(LinesDr)
Dim NCol%: NCol = Si(LinesDr)
Dim ColAy(): ColAy = ColRgAyzLinesDr(LinesDr)
Dim O(): ReDim O(1 To NRow, 1 To NCol)
Dim Col$(), ICol%, S$: For ICol = 0 To NCol - 1
    Col = ColAy(ICol)
    Dim IRow%: For IRow = 0 To UB(Col)
        S = Col(IRow)
        O(IRow + 1, ICol + 1) = S
    Next
Next
SqzLinesDr = O
End Function
