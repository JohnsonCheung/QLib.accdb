Attribute VB_Name = "MxFds"
Option Compare Text
Option Explicit
Const CNs$ = "as"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxFds."

Function CsvzFds$(A As DAO.Fields)
CsvzFds = CsvLinzDr(AvzItr(A))
End Function

Function DrzFds(A As DAO.Fields, Optional FF$) As Variant()
If FF = "" Then
    Dim F As DAO.Field
    For Each F In A
        PushI DrzFds, EmptyIfNull(F.Value)
    Next
    Exit Function
End If
DrzFds = DrzFdsFny(A, Ny(FF))
End Function

Function DrzFdsFny(A As DAO.Fields, Fny$()) As Variant()
Dim I, O()
For Each I In Fny
    Push O, EmptyIfNull(A(I).Value)
Next
End Function


Sub Z_DrzFds()
Dim Rs As DAO.Recordset, Dy()
Set Rs = Db(SampFbzShpRate).OpenRecordset("Select * from YMGRnoIR")
With Rs
    While Not .EOF
        PushI Dy, DrzFds(Rs.Fields)
        .MoveNext
    Wend
    .Close
End With
BrwDy Dy
End Sub

Sub Z_DrzFds1()
Dim Rs As DAO.Recordset, Dr(), D As Database
Set Rs = RszQ(SampDbDutyDta, "Select * from SkuB")
With Rs
    While Not .EOF
        Dr = DrzRs(Rs)
        Debug.Print JnComma(Dr)
        .MoveNext
    Wend
    .Close
End With
End Sub
