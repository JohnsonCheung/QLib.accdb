Attribute VB_Name = "MxRx"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxRx."
Function MchzPatn(Patn$, S) As MatchCollection
Set MchzPatn = Mch(Rx(Patn), S)
End Function
Function SyzMch(M As MatchCollection) As String()
Dim I As Match: For Each I In M
    PushI SyzMch, I.Value
Next
End Function
Function MchzPatnF(Patn$, S) As Match
Dim M As MatchCollection: Set M = MchzPatn(Patn, S)
If M.Count = 0 Then Exit Function
Set MchzPatnF = CvMch(M(0))
End Function

Function Mch(Re As RegExp, S) As MatchCollection
Set Mch = Re.Execute(S)
End Function

Function CvRe(A) As RegExp
Set CvRe = A
End Function

Function Rx(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As RegExp
If Patn = "" Or Patn = ".*" Then Exit Function
Dim O As New RegExp
With O
   .Pattern = Patn
   .MultiLine = MultiLine
   .IgnoreCase = IgnoreCase
   .Global = IsGlobal
End With
Set Rx = O
End Function

