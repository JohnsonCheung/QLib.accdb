Attribute VB_Name = "MxSampDta"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxSampDta."

Property Get SampDr_AToJ() As Variant()
Const NC% = 10
Dim J%
For J = 0 To NC - 1
    PushI SampDr_AToJ, Chr(Asc("A") + J)
Next
End Property

Property Get SampSq1() As Variant()
Dim O(), R&, C&
Const NR& = 1000
Const NC& = 100
ReDim O(1 To NR, 1 To NC)
For R = 1 To NR
For C = 1 To NC
    O(R, C) = R + C
Next
Next
SampSq1 = O
End Property
Property Get SampSqWithHdr() As Variant()
SampSqWithHdr = InsSqr(SampSq, SampDr_AToJ)
End Property

