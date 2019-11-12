Attribute VB_Name = "MxWsoTyDfn"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxWsoTyDfn."

Function WsoTyDfn() As Worksheet
Set WsoTyDfn = WsoTyDfnzP(CPj)
End Function

Function WsoTyDfnzP(P As VBProject) As Worksheet
Dim O As New Worksheet
Set O = NewWs("TyDfn")
'RgzSq DocSqzP(P), A1zWs(O)
Stop
FmtWsoTyDfn O
Set WsoTyDfnzP = O
End Function

Sub FmtWsoTyDfn(WsoTyDfn As Worksheet)

End Sub
