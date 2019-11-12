Attribute VB_Name = "MxDoMthn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoMthn."


Function DoMthnP() As Drs
DoMthnP = SelDrs(DoMthcP, FFoMthn)
End Function

Function DoMthnzV(V As Vbe) As Drs

End Function

Function DoMthnV() As Drs
DoMthnV = DoMthnzV(CVbe)
End Function

Function DoMthnM() As Drs
DoMthnM = DoMthnzM(CMd)
End Function
