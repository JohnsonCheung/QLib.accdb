Attribute VB_Name = "MxCsv"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxCsv."

Function DrzCsvLin(CsvLin) As String()
If Not HasDblQ(CsvLin) Then DrzCsvLin = SplitComma(CsvLin): Exit Function
End Function

Function CvCsv$(V)
Select Case True
Case IsStr(V): CvCsv = """" & V & """"
Case IsDte(V): CvCsv = Format(V, "YYYY-MM-DD HH:MM:SS")
Case IsEmpty(V):
Case Else: CvCsv = V
End Select
End Function

Function CsvStrzDrs$(D As Drs)
CsvStrzDrs = JnCrLf(CsvLyzDrs(D))
End Function

Function CsvLyzDrs(D As Drs) As String()
PushI CsvLyzDrs, JnComma(D.Fny)
Dim Dr: For Each Dr In Itr(D.Dy)
    PushI CsvLyzDrs, CsvLinzDr(Dr)
Next
End Function


Sub WrtDrsXls(D As Drs, Fcsv$)
'Do Wrt @D to @Fcvs using Xls-Style @@
DltFfnIf Fcsv
PushXlsVisHid
ClsWbNoSav SavWbCsv(NewWbzDrs(D), Fcsv)
PopXlsVis
End Sub

Sub WrtDrs(D As Drs, Fcsv$)
WrtStr CsvStrzDrs(D), Fcsv, OvrWrt:=True
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
    O(J) = CvCsv(V)
    J = J + 1
Next
CsvLinzDr = Join(O, ",")
End Function
