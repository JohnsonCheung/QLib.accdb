Attribute VB_Name = "MxDtaS12"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaS12."

Function DrszS12s(A As S12s, Optional FF$ = "S1 S2") As Drs
DrszS12s = DrszFF(FF, DyzS12s(A))
End Function

Function AvzS12(A As S12) As Variant()
AvzS12 = Array(A.S1, A.S2)
End Function

Function DyzS12s(A As S12s) As Variant()
'Ret : a 2 col of dry with fst row is @N1..2 and snd row is ULin and rst from @A @@
Dim J&: For J& = 0 To A.N - 1
    With A.Ay(J)
    PushI DyzS12s, Array(.S1, .S2)
    End With
Next
End Function

