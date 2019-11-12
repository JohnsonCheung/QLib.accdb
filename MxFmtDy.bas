Attribute VB_Name = "MxFmtDy"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxFmtDy."

Sub Z_FmtDy()
Dim A$()
A = FmtDy(SampDy3, Fmt:=EiSSFmt, BrkCCIxy0:=Array(0))
A = FmtDy(SampDy3, Fmt:=EiSSFmt)
Dmp A
End Sub

Function FmtDy(Dy(), _
Optional MaxColWdt% = 100, _
Optional BrkCCIxy0, _
Optional ShwZer As Boolean, _
Optional Fmt As EmTblFmt) _
As String()
If Si(Dy) = 0 Then Exit Function
Dim Dy1(): Dy1 = CellDy(Dy, ShwZer, MaxColWdt) ' :CellDy: ! each ele is a :Cell
Dim W%():                    W = WdtAyzCellDy(Dy1)
Dim Bdy$():                Bdy = FmtCellDy(Dy1, W, Fmt)
Dim Sep$:                  Sep = SepLin(W, Fmt)
Dim ShouldBrk() As Boolean:  ShouldBrk = ShouldBrkAy(Dy, BrkCCIxy0)
Dim BrkBdy$():          BrkBdy = InsSepLin(Bdy, ShouldBrk, Sep)
Dim Bdy1$():                Bdy1 = LyzLinesAy(BrkBdy)
                         FmtDy = Sy(Sep, Bdy1, Sep)
End Function

Function SepLin$(W%(), Fmt As EmTblFmt)
Dim Sep$():  Sep = SepDr(W)
Dim S$:        S = IIf(Fmt = EiSSFmt, " ", "-|-")
Dim Q$:        Q = IIf(Fmt = EiSSFmt, "", "|-*-|")
            SepLin = LinzDr(Sep, S, Q)
'Insp "QDta_Fun_FmtDy.SepLin", "Inspect", "Oup(SepLin) W Fmt", SepLin, W, "NoFmtr(EmTblFmt)": Stop
End Function

Function ShouldBrkAy(Dy(), BrkCCIxy0) As Boolean()
'Ret : ! no or sam ele of @Dy.  Telling that row of @Dy should ins a brk aft the row.  The las ele is always false.  @@
If Si(BrkCCIxy0) = 0 Then Exit Function
Dim Ixy&(): Ixy = IntozAy(EmpLngAy, BrkCCIxy0)
Dim LasK: LasK = AwIxy(Dy(0), Ixy)
Dim CurK
Dim Dr: For Each Dr In Itr(Dy)
    CurK = AwIxy(Dr, Ixy)
           PushI ShouldBrkAy, Not IsEqAy(CurK, LasK)
    LasK = CurK
Next
'Insp "QDta_Fun_FmtDy.ShouldBrkAy", "Inspect", "Oup(ShouldBrkAy) Dy BrkCCIxy0", "NoFmtr(Boolean())", "NoFmtr(())", BrkCCIxy0: Stop
End Function

Function InsSepLin(Bdy$(), ShouldBrk() As Boolean, SepLin$) As String()
If Si(ShouldBrk) = 0 Then InsSepLin = Bdy: Exit Function
Dim L, J&: For Each L In Bdy
    If ShouldBrk(J) Then PushI InsSepLin, SepLin
    PushI InsSepLin, L
    J = J + 1
Next
'Insp "QDta_Fun_FmtDy.InsSepLin", "Inspect", "Oup(InsSepLin) WrpBdy BrkAy Sep", InsSepLin, WrpBdy, "NoFmtr(() As Boolean)", Sep: Stop
End Function
