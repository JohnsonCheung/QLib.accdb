Attribute VB_Name = "MxIdx"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxIdx."

Function CvIdx(A) As DAO.Index
Set CvIdx = A
End Function

Function FnyzIdx(A As DAO.Index) As String()
If IsNothing(A) Then Exit Function
FnyzIdx = Itn(A.Fields)
End Function

Function IdxStr$(A As DAO.Index)
Dim X$, F$
With A
IdxStr = FmtQQ("Idx;?;?;?", .Name, X, F)
End With
End Function

Function IdxStrAyIdxs(A As DAO.Indexes) As String()
Dim I As DAO.Index
For Each I In A
    PushI IdxStrAyIdxs, IdxStr(I)
Next
End Function

Function IsEqIdx(A As DAO.Index, b As DAO.Index) As Boolean
With A
Select Case True
Case .Name <> b.Name
Case .Primary <> b.Primary
Case .Unique <> b.Unique
Case Not IsEqAy(Itn(.Fields), Itn(b.Fields))
Case Else: IsEqIdx = True
End Select
End With
End Function

Function IsEqIdxs(A As DAO.Indexes, b As DAO.Indexes) As Boolean
If A.Count <> b.Count Then Exit Function
If Not IsEqNmItr(A, b) Then Exit Function
Dim I
For Each I In A
    If Not IsEqIdx(CvIdx(I), b(CvIdx(I).Name)) Then Exit Function
Next
End Function

Function IsIdxSk(A As DAO.Index, T) As Boolean
If A.Name <> T Then Exit Function
IsIdxSk = A.Unique
End Function

Function IsIdxUniq(A As DAO.Index) As Boolean
If IsNothing(A) Then Exit Function
IsIdxUniq = A.Unique
End Function
