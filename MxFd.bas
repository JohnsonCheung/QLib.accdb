Attribute VB_Name = "MxFd"
Option Compare Text
Option Explicit
Const CNs$ = "sf"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxFd."

Function CvFd(A) As DAO.Field
Set CvFd = A
End Function

Function CvFd2(A As DAO.Field) As DAO.Field2
Set CvFd2 = A
End Function

Function FdClone(A As DAO.Field2, FldNm) As DAO.Field2
Set FdClone = New DAO.Field
With FdClone
    .Name = FldNm
    .Type = A.Type
    .AllowZeroLength = A.AllowZeroLength
    .Attributes = A.Attributes
    .DefaultValue = A.DefaultValue
    .Expression = A.Expression
    .Required = A.Required
    .ValidationRule = A.ValidationRule
    .ValidationText = A.ValidationText
End With
End Function

Function IsEqFd(A As DAO.Field2, B As DAO.Field2) As Boolean
With A
    If .Name <> B.Name Then Exit Function
    If .Type <> B.Type Then Exit Function
    If .Required <> B.Required Then Exit Function
    If .AllowZeroLength <> B.AllowZeroLength Then Exit Function
    If .DefaultValue <> B.DefaultValue Then Exit Function
    If .ValidationRule <> B.ValidationRule Then Exit Function
    If .ValidationText <> B.ValidationText Then Exit Function
    If .Expression <> B.Expression Then Exit Function
    If .Attributes <> B.Attributes Then Exit Function
    If .Size <> B.Size Then Exit Function
End With
IsEqFd = True
End Function

Function Fv(A As DAO.Field)
On Error Resume Next
Fv = A.Value
End Function
