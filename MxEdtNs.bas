Attribute VB_Name = "MxEdtNs"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxEdtNs."

Sub EdtNs()
WrtDrsAsXls DoMdP, ResFcsv("DoMdP")
BrwNs
End Sub

Sub SetNs()
'Do : Set CNsv in each module from ResFcsv("DoMdP")
Dim D As Drs
D = ResDrs("DoMdP")
D = SelDrs(D, "Mdn CNsv")
D = DwNBlnk(D, "CNsv")
Dim Dr: For Each Dr In Itr(D.Dy)
    Dim M As CodeModule: Set M = Md(Dr(0))
    Dim Ns$: Ns = Dr(1)
    EnsCNsLin M, Ns
Next
End Sub

Sub BrwNs()
OpnFcsv ResFcsv("DoMdP")
ShwXls Xls
LasWb.Activate
End Sub
