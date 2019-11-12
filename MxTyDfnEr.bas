Attribute VB_Name = "MxTyDfnEr"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxTyDfnEr."
Function IsLinTyDfnEr(Lin) As Boolean
Dim L$: L = Lin
If ShfTyDfnNm(L) = "" Then Exit Function
ShfDfnTy L
ShfMemNm L
If L = "" Then Exit Function
If FstChr(L) <> "!" Then IsLinTyDfnEr = True
End Function

Function TyDfnErLinAyP() As String()
Dim L: For Each L In SrczP(CPj)
    If IsLinTyDfnEr(L) Then
        PushI TyDfnErLinAyP, L
    End If
Next
End Function

Sub Z_TyDfnErLinAyP()
Brw TyDfnErLinAyP
End Sub
