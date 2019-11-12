Attribute VB_Name = "MxIsCell"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxIsCell."
Function IsCellInRg(A As Range, Rg As Range) As Boolean
Dim R&, C%, R1&, R2&, C1%, C2%
R = A.Row
R1 = Rg.Row
If R < R1 Then Exit Function
R2 = R1 + Rg.Rows.Count
If R > R2 Then Exit Function
C = A.Column
C1 = Rg.Column
If C < C1 Then Exit Function
C2 = C1 + Rg.Columns.Count
If C > C2 Then Exit Function
IsCellInRg = True
End Function

Function IsCellInRgAp(Cell As Range, ParamArray RgAp()) As Boolean
Dim Av(): Av = RgAp
'IsCellInRgAp = IsCellInRgAv(A, Av)
End Function

Function IsCellInRgAv(A As Range, RgAv()) As Boolean
Dim V
For Each V In RgAv
    If IsCellInRg(A, CvRg(V)) Then IsCellInRgAv = True: Exit Function
Next
End Function

