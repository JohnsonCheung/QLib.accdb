Attribute VB_Name = "MxRel"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxRel."

Property Get SampRel() As Rel
Set SampRel = Rel(SampRelLy)
End Property

Property Get SampRelLy() As String()
Erase XX
X "A B"
X "B A"
SampRelLy = XX
Erase XX
End Property

Property Get SamPubMthRel() As Rel
Set SamPubMthRel = Rel(SamPubMthRelLy)
End Property

Property Get SamPubMthRelLy() As String()
'SampMthRelLy = RelOf_MthSDNm_To_Mdn_InVbe
End Property



Function CvRel(A) As Rel
Set CvRel = A
End Function

Property Get EmpRel() As Rel
Set EmpRel = New Rel
End Property

Function IsRel(A) As Boolean
IsRel = TypeName(A) = "Rel"
End Function

Function RelzVbl(RelVbl$) As Rel
Set RelzVbl = Rel(SplitVBar(RelVbl))
End Function

Function Rel(RelLy$()) As Rel
Dim O As New Rel
Set Rel = O.Init(RelLy)
End Function

Function RelVbl(Vbl$) As Rel
Set RelVbl = Rel(SplitVBar(Vbl))
End Function
