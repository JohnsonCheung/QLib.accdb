Attribute VB_Name = "MxCnstStrv"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCnstStrv."

Function CnstStrvzN$(Lin, Cnstn$)
If IsLinCnstzN(Lin, Cnstn) Then CnstStrvzN = BetDblQ(Lin)
End Function

Function CnstStrvzDcl(Dcl$(), Cnstn$)
Dim L, O$: For Each L In Itr(Dcl)
    O = CnstStrvzN(L, Cnstn)
    If O <> "" Then CnstStrvzDcl = O: Exit Function
Next
End Function

Function CnstStrv$(Lin)
If IsLinCnstStr(Lin) Then CnstStrv = BetDblQ(Lin)
End Function

Sub Z_CnstStrv()
Dim O$()
Dim L: For Each L In SrczP(CPj)
    PushNB O, CnstStrv(L)
Next
BrwAy O
End Sub

