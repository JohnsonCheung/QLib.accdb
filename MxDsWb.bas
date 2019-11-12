Attribute VB_Name = "MxDsWb"
Option Compare Database

Function WbzDs(D As Ds, Optional DsWbSpec$) As Workbook
Dim O As Workbook: Set O = NewWb
DsWb_PutWs O, D
DsWb_PutPt O, D, DsWbSpec
Set WbzDs = O
End Function

Sub DsWb_PutWs(b As Workbook, D As Ds)
Dim J%: For J = 0 To D.N - 1
    WszDt
Next
End Sub

Sub DsWb_PutPt(b As Workbook, D As Ds, DsWbSpec$)

End Sub

