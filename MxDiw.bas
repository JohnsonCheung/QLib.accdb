Attribute VB_Name = "MxDiw"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxDiw."

Function DiwAy(DiAqB As Dictionary, Ay) As Dictionary
'Ret : :DiAqB #SubSet-Of-Dic-By-Ay#
Set DiwAy = New Dictionary
Dim A: For Each A In Itr(Ay)
    DiwAy.Add A, DiAqB(A)
Next
End Function

