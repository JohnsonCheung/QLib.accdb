Attribute VB_Name = "MxAset"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAset."
Property Get EmpAset() As Aset
Set EmpAset = New Aset
End Property

Function CvAset(V) As Aset
Set CvAset = V
End Function

Function IsAset(V) As Boolean
IsAset = TypeName(V) = "Aset"
End Function

Function AsetzAp(ParamArray Ap()) As Aset
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
Set AsetzAp = AsetzAy(Av)
End Function
Function AsetzItr(Itr) As Aset
Set AsetzItr = EmpAset
AsetzItr.PushItr Itr
End Function
Function AsetzFF(FF$) As Aset
Set AsetzFF = AsetzAy(TermAy(FF))
End Function
Function AsetzSsl(Ssl$) As Aset
Set AsetzSsl = EmpAset
Dim Sy$(): Sy = SyzSS(Ssl)
If HasDup(Sy) Then Thw CSub, "Ssl has dup", "Ssl DupEle", Ssl, AwDup(Sy)
AsetzSsl.PushAy SyzSS(Ssl)
End Function

Function AsetzAy(Ay) As Aset
Set AsetzAy = EmpAset
AsetzAy.PushAy Ay
End Function

Function AsetzItm(Itm) As Aset
Set AsetzItm = EmpAset
AsetzItm.PushItm Itm
End Function
