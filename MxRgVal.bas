Attribute VB_Name = "MxRgVal"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxRgVal."
Function VvyzRgH(A As Range) As Variant()
VvyzRgH = VvyzRg(RgC(A, 1))
End Function
Function VvyzRgV(A As Range) As Variant()
VvyzRgV = VvyzRg(RgC(A, 1))
End Function
Function VvyzRg(A As Range) As Variant()
VvyzRg = Vvy(A)
End Function
Function SvyzRgV(A As Range) As String()
SvyzRgV = Svy(RgC(A, 1))
End Function


