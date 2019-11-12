Attribute VB_Name = "MxLDta"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxLDta."
Public Const FFoLDta$ = "L Dta"
Type TDoLDta: D As Drs: End Type
Private Function DoLDtazT1(DoLTD As Drs, T1$) As Drs
DoLDtazT1 = DwEqExl(DoLTD, "T1", T1)
End Function

Function TDoLDtazT1(A As TDoLTD, T1$) As TDoLDta
TDoLDtazT1.D = DoLDtazT1(A.D, T1)
End Function

