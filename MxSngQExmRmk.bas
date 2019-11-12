Attribute VB_Name = "MxSngQExmRmk"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxSngQExmRmk."
Function RmkzSngQExmLin$(Lin)

End Function
Function RmkzTyDfnRmkLy$(TyDfnRmkLy$())
Dim R$, O$()
Dim L: For Each L In Itr(TyDfnRmkLy)
    If FstChr(L) = "'" Then
        Dim A$: A = LTrim(RmvFstChr(L))
        If FstChr(A) = "!" Then
            PushNB O, LTrim(RmvFstChr(A))
        End If
    End If
Next
RmkzTyDfnRmkLy = JnCrLf(O)
End Function
Function SngQExmRe() As RegExp
Static O As RegExp
If IsNothing(O) Then
End If
End Function

Function IsLinSngQExm(L) As Boolean
IsLinSngQExm = SngQExmRe.Test(L)
End Function
