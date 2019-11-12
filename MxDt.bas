Attribute VB_Name = "MxDt"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDt."
Type Dt: DtNm As String: Fny() As String: Dy() As Variant: End Type

Sub BrwDt(A As Dt, Optional Fnn$)
BrwAy FmtDt(A), Dft(Fnn, A.DtNm)
End Sub

Function CsvLyzDt(A As Dt) As String()
Dim Dy(): Dy = A.Dy
Push CsvLyzDt, JnComma(QteSyDbl(A.Fny))
Dim QQStr$: QQStr = CsvQQStrzDr(Dr(Dy, 0))
Dim IDr
For Each IDr In A.Dy
   PushI CsvLyzDt, FmtQQAv(QQStr, CvAv(IDr))
Next
End Function

Function CsvQQStrzDr$(Dr())
Dim O$(), I
For Each I In Dr
    If IsStr(I) Then
        PushI O, """?"""""
    Else
        PushI O, "?"
    End If
Next
CsvQQStrzDr = JnComma(O)
End Function

Sub DmpDt(A As Dt)
DmpAy FmtDt(A)
End Sub

Function Dr(Dy(), R&) As Variant()
Dr = Dy(R)
End Function

Function DrszDt(A As Dt) As Drs
DrszDt = Drs(A.Fny, A.Dy)
End Function

Function Dt(DtNm, Fny() As String, Dy() As Variant) As Dt
With Dt
    .DtNm = DtNm
    .Fny = Fny
    .Dy = Dy
End With
End Function

Function DtDrpCol(A As Dt, CC$, Optional DtNm$) As Dt
DtDrpCol = DtzDrs(DrpCol(DrszDt(A), CC), Dft(DtNm, A.DtNm))
End Function

Function DtReOrd(A As Dt, BySubFF$) As Dt
DtReOrd = DtzDrs(ReOrdCol(DrszDt(A), BySubFF), A.DtNm)
End Function

Function DtSelCol(A As Dt, CC$, Optional DtNm$) As Dt
DtSelCol = DtzDrs(SelDrs(DrszDt(A), CC), Dft(DtNm, A.DtNm))
End Function

Function DtzDrs(A As Drs, Optional DtNm$ = "Dt") As Dt
DtzDrs = Dt(DtNm, A.Fny, A.Dy)
End Function

Function DtzFF(DtNm$, FF$, Dy()) As Dt
DtzFF = Dt(DtNm, Ny(FF), Dy)
End Function

Function DtzNmDrs(DtNm$, A As Drs) As Dt
DtzNmDrs = DtzDrs(A, DtNm)
End Function

Property Get EmpDtAy() As Dt()
End Property

Function IsEmpDt(A As Dt) As Boolean
IsEmpDt = Si(A.Dy) = 0
End Function

Function NRowzDrs&(A As Drs)
NRowzDrs = Si(A.Dy)
End Function

Function NRowzDt&(A As Dt)
NRowzDt = Si(A.Dy)
End Function
