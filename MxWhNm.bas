Attribute VB_Name = "MxWhNm"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxWhNm."
Type WhNm
    Re As RegExp
    ExlLikAy() As String
    LikAy() As String
    IsEmp As Boolean
End Type

Function EmpWhNm() As WhNm
End Function

Function IsEqWhNm(A As WhNm, B As WhNm) As Boolean
With A
Select Case True
Case _
    ObjPtr(.Re) <> ObjPtr(.Re), _
    IsEqAy(.ExlLikAy, B.ExlLikAy), _
    IsEqAy(.LikAy, B.LikAy)
Case Else
    IsEqWhNm = True
End Select
End With
End Function

Function WhNmStr$(A As WhNm)
'If IsEmpWhNm(A) Then ToStr = "#Emp": Exit Function
Dim O$()
'Push O, Qte(X_Re.Pattern, "Patn(*)")
'Push O, Qte(TLin(X_LikAy), "LikAy(*)")
'Push O, Qte(TLin(X_ExlLikAy), "ExlLikAy(*)")
'ToStr = JnCrLf(O)
End Function

Function IsEmpWhNm(A As WhNm) As Boolean
IsEmpWhNm = IsEqWhNm(A, EmpWhNm)
End Function
