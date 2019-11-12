Attribute VB_Name = "MxXlsDft"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsDft."
Function DftWsn$(Wsn0$, Fx)
If Wsn0 = "" Then
    DftWsn = FstWsn(Fx)
    Exit Function
End If
DftWsn = Wsn0
End Function

Function DftWny(Wny0, Fx) As String()
Dim O$()
    O = CvSy(Wny0)
If Si(O) = 0 Then
    DftWny = WnyzFx(Fx)
Else
    DftWny = O
End If
End Function
