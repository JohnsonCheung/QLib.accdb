Attribute VB_Name = "MxMnm"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMnm."
Private Property Get MnmRe() As RegExp
Static X As RegExp
If IsNothing(X) Then
    Set X = Rx("#[A-Z]\S+#", IsGlobal:=True)
End If
Set MnmRe = X
End Property
Function MnmAyP() As String()
MnmAyP = MnmAy(SrclP)
End Function
Private Sub Z_MnmAyP()
Brw MnmAyP
End Sub
Function MnmAy(S) As String()
':Mnm: :NoSpcStr #Memonic-Name# ! A string without space quoted with [#]
MnmAy = SyzMch(MnmRe.Execute(S))
End Function
