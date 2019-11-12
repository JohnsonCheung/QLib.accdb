Attribute VB_Name = "MxIsXXXDic"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxIsXXXDic."
Function IsDicKeyNm(A As Dictionary) As Boolean
Dim K
For Each K In A.Keys
    If Not IsNm(K) Then Exit Function
Next
IsDicKeyNm = True
End Function

Function IsDicKeyStr(A As Dictionary) As Boolean
IsDicKeyStr = IsItrSy(A.Keys)
End Function

Function IsDicEmp(A As Dictionary) As Boolean
IsDicEmp = A.Count = 0
End Function

Function IsDicLy(A As Dictionary) As Boolean
Dim D As Dictionary, I, V
If Not IsDic(A) Then Exit Function
Dim Ly: For Each Ly In A.Items
    If Not IsLy(Ly) Then Exit Function
Next
IsDicLy = True
End Function
Function IsDic(V) As Boolean
IsDic = TypeName(V) = "Dictionary"
End Function

Function IsDicSy(A As Dictionary) As Boolean
Dim D As Dictionary, I, V
If Not IsDic(A) Then Exit Function
IsDicSy = IsItrSy(CvDic(A).Items)
End Function

Function IsDicLines(A As Dictionary) As Boolean
IsDicLines = True
If IsItrLines(A.Items) Then Exit Function
If IsItrStr(A.Keys) Then Exit Function
IsDicLines = False
End Function

Function IsStrDic(A As Dictionary) As Boolean
If Not IsItrStr(A.Keys) Then Exit Function
IsStrDic = IsItrStr(A.Items)
End Function

Function IsDicPrim(A As Dictionary) As Boolean
If Not IsItrPrim(A.Keys) Then Exit Function
IsDicPrim = IsItrPrim(A.Items)
End Function

