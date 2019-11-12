Attribute VB_Name = "MxDtaEr"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaEr."

Function EoColDup(D As Drs, C$) As String()
Dim b As Drs: b = DwDup(D, C)
Dim Msg$: Msg = "Dup [" & C & "]"
EoColDup = EoDrsMsg(b, Msg)
End Function

