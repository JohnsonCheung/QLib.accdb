Attribute VB_Name = "MxRe"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxRe."

Sub Z_ReMatch()
Dim A As MatchCollection
Dim R  As RegExp: Set R = Rx("m[ae]n")
Set A = R.Execute("alskdflfmEnsdklf")
Stop
End Sub

Sub Z_ReRpl()
Dim R As RegExp: Set R = Rx("(.+)(m[ae]n)(.+)")
Dim Act$: Act = R.Replace("a men is male", "$1male$3")
Ass Act = "a male is male"
End Sub
