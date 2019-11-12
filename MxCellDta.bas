Attribute VB_Name = "MxCellDta"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxCellDta."

Function WdtAyzCellDy(CellDy()) As Integer()
':CellDy: :Dy ! each dr is sy and  ele of the sy is a cell
Dim J&: For J = 0 To NColzDy(CellDy) - 1
    Push WdtAyzCellDy, WdtzCellCol(StrColzDy(CellDy, J))
Next
End Function

Function FmtCellDr$(CellDr$(), W%(), S$, Q$)
'@S :SepChr
'@Q :QteChr
Dim Sq(): Sq = SqzLinesDr(CellDr)
Dim Sq1(): Sq1 = AlignSqzW(Sq, W)
Dim Dr, Lin$, O$(), IR%: For IR = 1 To UBound(Sq, 1)
    Dr = DrzSq(Sq1, IR)
    Lin = LinzDr(Dr, S, Q)
    PushI O, Lin
Next
FmtCellDr = JnCrLf(O)
End Function

Function CellDy(Dy(), ShwZer As Boolean, MaxColWdt%) As Variant()
Dim J%
Dim Dr, StrDr$(): For Each Dr In Itr(Dy)
    StrDr = CellDr(Dr, ShwZer, MaxColWdt)
    J = J + 1
    Push CellDy, StrDr
Next
'Insp "QDta_Fun_FmtDy.CellDy", "Inspect", "Oup(CellDy) Dy ShwZer MaxColWdt", "NoFmtr(Variant())", "NoFmtr(())", ShwZer, "NoFmtr(% = 100)": Stop
End Function

Function CellDr(Dr, ShwZer As Boolean, MaxWdt%) As String()
Dim V, S$: For Each V In Itr(Dr)
    S = Cell(V, ShwZer, MaxWdt)
    PushI CellDr, S
Next
End Function

Function FmtCellDy(CellDy(), W%(), Fmt As EmTblFmt) As String()
'Ret : sam ele as @Dy.  One ele of @Ret may be a lin or lines depending on any cell of Dr of @Dy has lines.
Dim S$: S = IIf(Fmt = EiSSFmt, " ", " | ")
Dim Q$: Q = IIf(Fmt = EiSSFmt, "", "| * |")
Dim Dr: For Each Dr In Itr(CellDy)
    PushI FmtCellDy, FmtCellDr(CvSy(Dr), W, S, Q)
Next
'Insp "QDta_Fun_FmtDy.FmtCellDy", "Inspect", "Oup(FmtCellDy) CellDy W Fmt", FmtCellDy, "NoFmtr(())", W, "NoFmtr(EmTblFmt)": Stop
End Function

Function WdtzCellCol%(CellCol$())
Dim O%, Cell
For Each Cell In Itr(CellCol)
    O = Max(O, WdtzCell(Cell))
Next
WdtzCellCol = O
End Function

Function WdtzCell%(Cell)
If IsLines(Cell) Then
    WdtzCell = WdtzLines(Cell)
Else
    WdtzCell = Len(Cell)
End If
End Function
