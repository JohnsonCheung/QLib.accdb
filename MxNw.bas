Attribute VB_Name = "MxNw"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxNw."
Function NwApp(AppHom$, Appn$, Appv$, OupHom$) As App
Dim A As New App
A.Ini AppHom, Appn, Appv, OupHom
Set NwApp = A
End Function


