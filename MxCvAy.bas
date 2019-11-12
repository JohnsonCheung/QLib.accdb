Attribute VB_Name = "MxCvAy"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxSeq."


Function CvBytAy(A) As Byte()
CvBytAy = A
End Function

Function CvIntAy(A) As Integer()
On Error Resume Next
CvIntAy = A
End Function

Function CvLngAy(A) As Long()
On Error Resume Next
CvLngAy = A
End Function
Function CvAv(A) As Variant()
If VarType(A) = vbArray + vbVariant Then
    If Si(A) = -1 Then Exit Function
End If
CvAv = A
End Function
Function CvObj(A) As Object
Set CvObj = A
End Function

Function CvSy(Str_or_Sy_or_Ay_or_EmpMis_or_Oth) As String()
Dim A: A = Str_or_Sy_or_Ay_or_EmpMis_or_Oth
Select Case True
Case IsStr(A): PushI CvSy, A
Case IsSy(A): CvSy = A
Case IsArray(A): CvSy = SyzAy(A)
Case IsEmpty(A) Or IsMissing(A)
Case Else: CvSy = Sy(A)
End Select
End Function

