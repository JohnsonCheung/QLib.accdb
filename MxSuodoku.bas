Attribute VB_Name = "MxSuodoku"
Option Compare Text
Option Explicit
Const CLib$ = "QSudoku."
Const CMod$ = CLib & "MxSuodoku."
Private Type RRCC
    R1 As Byte 'all started from 1
    R2 As Byte
    C1 As Byte
    C2 As Byte
End Type
Private Type SolveRslt
    SudokuSq() As Variant
    HasSolve As Boolean
End Type
Private Type NineEleRslt
    NineEle() As Variant
    HasSolve As Boolean
End Type

Private Function RRCC(R1 As Byte, R2 As Byte, C1 As Byte, C2 As Byte) As RRCC
With RRCC
    .R1 = R1
    .R2 = R2
    .C1 = C1
    .C2 = C2
End With
End Function

Function SolveFstRound(Sq()) As Variant()
Dim J%
For J = 1 To 9
    SetNineEleRow Sq(), J, SolveNineEleFstRnd(NineEleRow(Sq(), J))
Next
SolveFstRound = Sq
End Function

Function Solve(SudokuSq()) As Variant()
Dim O(), HasSolve As Boolean, J%
O = SolveFstRound(SudokuSq)
HasSolve = True
While HasSolve
    J = J + 1: If J > 1000 Then Stop
    HasSolve = False
    With SolveRow(O): O = IIf(.HasSolve, .SudokuSq, O): HasSolve = IIf(.HasSolve, True, HasSolve): End With
    With SolveCol(O): O = IIf(.HasSolve, .SudokuSq, O): HasSolve = IIf(.HasSolve, True, HasSolve): End With
    With SolveDiag(O): O = IIf(.HasSolve, .SudokuSq, O): HasSolve = IIf(.HasSolve, True, HasSolve): End With
    With SolveSmallSq(O): O = IIf(.HasSolve, .SudokuSq, O): HasSolve = IIf(.HasSolve, True, HasSolve): End With
Wend
Solve = O
End Function

Function SolveRow(Sq()) As SolveRslt
Dim J%, O()
O = Sq
For J = 1 To 9
    With SolveNineEle(NineEleRow(O, J))
        If .HasSolve Then
            SolveRow.HasSolve = True
            SetNineEleRow O, J, .NineEle
        End If
    End With
Next
SolveRow.SudokuSq = O
End Function

Function NineEleRow(Sq(), Row%) As Variant()
Dim J%
For J = 1 To 9
    PushI NineEleRow, Sq(Row, J)
Next
End Function

Sub SetNineEleRow(Sq(), Row%, NineEle())
Dim J%
For J = 1 To 9
    Sq(Row, J) = NineEle(J - 1)
Next
End Sub

Function SolveSmallSq(Sq()) As SolveRslt
Dim J%, O()
O = Sq
For J = 1 To 9
    With SolveNineEle(NineEleSmallSq(O, J))
        If .HasSolve Then
            SolveSmallSq.HasSolve = True
            SetNineEleSmallSq O, J, .NineEle
        End If
    End With
Next
SolveSmallSq.SudokuSq = O
End Function

Function NineEleSmallSq(Sq(), J%) As Variant()
Dim R As Byte, C As Byte
With RRCCzJ(J)
For R = .R1 To .R2
    For C = .C1 To .C2
        PushI NineEleSmallSq, Sq(R, C)
    Next
Next
End With
End Function

Function RRCCzJ(J%) As RRCC
Select Case J
Case 1: RRCCzJ = RRCC(1, 3, 1, 3)
Case 2: RRCCzJ = RRCC(1, 3, 4, 6)
Case 3: RRCCzJ = RRCC(1, 3, 7, 9)
Case 4: RRCCzJ = RRCC(4, 6, 1, 3)
Case 5: RRCCzJ = RRCC(4, 6, 4, 6)
Case 6: RRCCzJ = RRCC(4, 6, 7, 9)
Case 7: RRCCzJ = RRCC(7, 9, 1, 3)
Case 8: RRCCzJ = RRCC(7, 9, 4, 6)
Case 9: RRCCzJ = RRCC(7, 9, 7, 9)
Case Else: Thw CSub, "Invalid J, should be 1 to 9", "J", J
End Select
End Function

Sub SetNineEleSmallSq(Sq(), J%, NineEle())
Dim R As Byte, C As Byte
Dim I%
With RRCCzJ(J)
For R = .R1 To .R2
    For C = .C1 To .C2
        Sq(R, C) = NineEle(I)
        I = I + 1
    Next
Next
End With
End Sub

Function SolveCol(Sq()) As SolveRslt
Dim J%, O()
O = Sq
For J = 1 To 9
    With SolveNineEle(NineEleCol(O, J))
        If .HasSolve Then
            SolveCol.HasSolve = True
            SetNineEleCol O, J, .NineEle
        End If
    End With
Next
SolveCol.SudokuSq = O
End Function

Function NineEleCol(Sq(), Col%) As Variant()
Dim J%
For J = 1 To 9
    PushI NineEleCol, Sq(J, Col)
Next
End Function

Sub SetNineEleCol(Sq(), Col%, NineEle())
Dim J%
For J = 1 To 9
    Sq(J, Col) = NineEle(J - 1)
Next
End Sub

Function SolveDiag(Sq()) As SolveRslt
Dim J%, O()
O = Sq
With SolveNineEle(NineEleDiag1(O))
    If .HasSolve Then
        SolveDiag.HasSolve = True
        SetNineEleDiag1 O, .NineEle
    End If
End With
With SolveNineEle(NineEleDiag2(O))
    If .HasSolve Then
        SolveDiag.HasSolve = True
        SetNineEleDiag2 O, .NineEle
    End If
End With
SolveDiag.SudokuSq = O
End Function

Function NineEleDiag1(Sq()) As Variant()
Dim J%
For J = 1 To 9
    PushI NineEleDiag1, Sq(J, J)
Next
End Function

Sub SetNineEleDiag1(Sq(), NineEle())
Dim J%
For J = 1 To 9
    Sq(J, J) = NineEle(J - 1)
Next
End Sub

Function NineEleDiag2(Sq()) As Variant()
Dim J%
For J = 1 To 9
    PushI NineEleDiag2, Sq(10 - J, 10 - J)
Next
End Function

Sub SetNineEleDiag2(Sq(), NineEle())
Dim J%
For J = 1 To 9
    Sq(10 - J, 10 - J) = NineEle(J - 1)
Next
End Sub

Function SolveNineEleFstRnd(NineEle()) As Variant()
Dim Should() As Byte: Should = ShouldBe(NineEle)
Dim J%, I
Dim O(): O = NineEle
For Each I In NineEle
    If IsEmpty(I) Then
        O(J) = Should
    End If
    J = J + 1
Next
SolveNineEleFstRnd = O
End Function

Function SolveNineEle(NineEle()) As NineEleRslt
Dim Should() As Byte: Should = ShouldBe(NineEle)
Dim O(): O = NineEle
Dim M
Dim I, J%
For Each I In NineEle
    If IsBytAy(I) Then
        M = AyIntersect(CvBytAy(I), Should)
        If Si(I) > Si(M) Then
            SolveNineEle.HasSolve = True
            O(J) = M
        End If
    Else
        If Not IsByt(I) Then Stop
    End If
    J = J + 1
Next
SolveNineEle.NineEle = O
End Function

Function ShouldBe(NineEle()) As Byte()
Dim Certain() As Byte
Dim I
For Each I In NineEle
    If IsByt(I) Then PushI Certain, I
Next
Dim J As Byte
For J = 1 To 9
    If Not HasEle(Certain, J) Then PushI ShouldBe, J
Next
End Function

Sub SolveSudoku(Ws As Worksheet)
PutSudokuSolution Ws, Solve(SudokuSq(Ws))
End Sub

Function SudokuSq(Ws As Worksheet) As Variant()
Dim O(): O = RgRCRC(A1zWs(Ws), 1, 1, 9, 9)
Dim I%, J%
For J = 1 To 9
    For I = 1 To 9
        If Not IsEmpty(O(J, I)) Then
            O(J, I) = CByte(O(J, I))
        End If
    Next
Next
SudokuSq = O
End Function

Sub PutSudokuSolution(Ws As Worksheet, Sq())
SolutionRg(Ws).Value = Sq
End Sub

Function SolutionRg(Ws As Worksheet) As Range
Set SolutionRg = RgRCRC(A1zWs(Ws), 11, 1, 19, 9)
End Function

Property Get SampSudokuSq() As Variant()
Dim E: E = Empty
SampSudokuSq = SqzDy(Av( _
Array(5, E, 7, 6, 9, E, E, E, 2), _
Array(9, 3, E, E, E, 2, 7, 4, 5), _
Array(E, E, E, 3, E, 7, 1, E, E), _
Array(E, 4, 5, E, 6, E, 3, E, 8), _
Array(2, E, E, 4, E, E, E, E, E), _
Array(E, E, E, E, E, 8, 1, E, 2), _
Array(E, E, 9, E, 2, E, E, 1, 3), _
Array(3, E, E, E, E, 6, E, 5, 7), _
Array(7, E, E, 1, 3, E, 9, 8, 4)))
End Property

Sub Z_PutSampSudoku()
PutSampSudoku WsRC(ActiveSheet, 1, "L")
End Sub

Sub PutSampSudoku(At As Range)
RgRCRC(At, 1, 1, 9, 9).Value = SampSudokuSq
FmtSudoku At
End Sub

Sub FmtSudoku(At As Range)
BdrAround RgRCRC(At, 1, 1, 3, 3)
BdrAround RgRCRC(At, 1, 4, 3, 6)
BdrAround RgRCRC(At, 1, 7, 3, 9)
BdrAround RgRCRC(At, 4, 1, 6, 3)
BdrAround RgRCRC(At, 4, 4, 6, 6)
BdrAround RgRCRC(At, 4, 7, 6, 9)
BdrAround RgRCRC(At, 7, 1, 9, 3)
BdrAround RgRCRC(At, 7, 4, 9, 6)
BdrAround RgRCRC(At, 7, 7, 9, 9)
RgCC(At, 1, 9).EntireColumn.ColumnWidth = 2
End Sub

Sub Z_SolveSudoku()
Dim Ws As Worksheet
GoSub T0
Exit Sub
T0:
    Set Ws = ActiveSheet
    GoTo Tst
Tst:
    SolveSudoku Ws
    Return
End Sub

