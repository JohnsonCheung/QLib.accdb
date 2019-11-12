Attribute VB_Name = "MxLLin"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxLLin."
Type LLin
    Lno As Integer
    Lin As String
End Type

Function LLin(Lno&, Lin$) As LLin
LLin.Lno = Lno
LLin.Lin = Lin
End Function
