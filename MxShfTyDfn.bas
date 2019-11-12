Attribute VB_Name = "MxShfTyDfn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxShfTyDfn."

Function ShfTyDfnNm$(OLin$)
Dim A$: A = T1(OLin)
If IsTyDfnNm(A) Then
    ShfTyDfnNm = A
    OLin = RmvT1(OLin)
End If
End Function

Function ShfDfnTy$(OLin$)
Dim A$: A = T1(OLin)
If IsDfnTy(A) Then
    ShfDfnTy = A
    OLin = RmvT1(OLin)
End If
End Function

Function ShfMemNm$(OLin$)
Dim A$: A = T1(OLin)
If IsMemNm(A) Then
    ShfMemNm = A
    OLin = RmvT1(OLin)
End If
End Function

