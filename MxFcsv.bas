Attribute VB_Name = "MxFcsv"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxFcsv."

Function DrszFcsvXls(Fcsv$) As Drs
OpnFcsv Fcsv
DrszFcsvXls = DrszAldta(LasWs(LasWb))
ClsLasWbNoSav
End Function

Function DrszFcsv(Fcsv$) As Drs
Dim Ly$(): Ly = LyzFt(Fcsv)
Dim Fny$(): Fny = DrzCsvLin(Ly(0))
Fny(0) = RmvUtfSig(Fny(0))
Dim Dy()
    Dim J&: For J = 1 To UB(Ly)
        PushSomSi Dy, SplitComma(Ly(J))
    Next
DrszFcsv = Drs(Fny, Dy)
End Function
