Attribute VB_Name = "MxCsv"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxCsv."

Function DrzCsvLin(CsvLin) As Variant()
'If Not HasDblQ(CsvLin) Then DrzCsvLin = SplitComma(CsvLin): Exit Function
Stop
End Function

Function Csv$(V)
Select Case True
Case IsStr(V): Csv = """" & V & """"
Case IsDte(V): Csv = Format(V, "YYYY-MM-DD HH:MM:SS")
Case IsEmpty(V):
Case Else: Csv = V
End Select
End Function

Function CsvlzDrs$(D As Drs)
':Csvl: :Lines #Csv-Lines#
CsvlzDrs = JnCrLf(CsvLyzDrs(D))
End Function

Function CsvLyzDrs(D As Drs) As String()
':CsvLy: :Ly #Csv-Ly#
PushI CsvLyzDrs, JnComma(D.Fny)
Dim Dr: For Each Dr In Itr(D.Dy)
    PushI CsvLyzDrs, CsvLinzDr(Dr)
Next
End Function

Sub WrtDrsAsXls(D As Drs, Fcsv$)
'Do Wrt @D to @Fcvs using Xls-Style @@
DltFfnIf Fcsv
PushXlsVisHid
ClsWbNoSav SavWbCsv(NewWbzDrs(D), Fcsv)
PopXlsVis
End Sub

Sub WrtDrs(D As Drs, Fcsv$)
WrtStr CsvlzDrs(D), Fcsv, OvrWrt:=True
End Sub

Sub WrtDrsRes(D As Drs, ResFnn$, Optional Pseg$)
Dim F$: F = ResFfn(ResFnn & ".csv", Pseg)
WrtDrs D, F
End Sub

Function CsvLinzDr$(Dr)
If Si(Dr) = 0 Then Exit Function
Dim O$(), U&, J&, V
U = UB(Dr)
ReDim O(U)
For Each V In Dr
    O(J) = Csv(V)
    J = J + 1
Next
CsvLinzDr = Join(O, ",")
End Function
