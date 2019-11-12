Attribute VB_Name = "MxEndDte"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxEndDte."

Sub UpdEndDte(D As Database, T, EndDteFld$, BegDteFld$, GpFF)
Dim LasBegDte As Date
LasBegDte = DateSerial(2099, 12, 31)
Dim Q$
''Q = SqlSel_FF_Fm_Ordff(Sy(BegDteFld, EndDteFld), T, BegDteFld)
Stop
With Rs(D, Q)
    While Not .EOF
        .Edit
        .Fields(EndDteFld).Value = LasBegDte
        .Update
        .MoveNext
    Wend
    .Close
End With
End Sub
