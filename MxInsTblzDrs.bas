Attribute VB_Name = "MxInsTblzDrs"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxInsTblzDrs."
Sub InsTblzDrs(D As Database, T, b As Drs)
Dim F$(): F = AyIntersect(Fny(D, T), b.Fny)
InsRszDy RszTFny(D, T, F), SelDrsFny(b, F).Dy
End Sub

Sub InsTblzDy(D As Database, T, Dy())
InsRszDy RszT(D, T), Dy
End Sub
