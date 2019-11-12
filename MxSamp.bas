Attribute VB_Name = "MxSamp"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxSamp."
Property Get SampDr1() As Variant()
SampDr1 = Array(1, 2, 3)
End Property

Property Get SampDr2() As Variant()
SampDr2 = Array(2, 3, 4)
End Property

Property Get SampDr3() As Variant()
SampDr3 = Array(3, 4, 5)
End Property

Property Get SampDr4() As Variant()
SampDr4 = Array(43, 44, 45)
End Property

Property Get SampDr5() As Variant()
SampDr5 = Array(53, 54, 55)
End Property

Property Get SampDr6() As Variant()
SampDr6 = Array(63, 64, 65)
End Property

Property Get SampDrs2() As Drs
SampDrs2 = DrszFF("A B C", SampDy2)
End Property

Property Get SampDrs1() As Drs
SampDrs1 = DrszFF("A B C", SampDy1)
End Property

Property Get SampDrs() As Drs
SampDrs = DrszFF("A B C D E G H I J K", SampDy)
End Property

Property Get SampDy1() As Variant()
SampDy1 = Array(SampDr1, SampDr2, SampDr3)
End Property

Property Get SampDy2() As Variant()
SampDy2 = Array(SampDr3, SampDr4, SampDr5)
End Property

Property Get SampDy3() As Variant()
PushI SampDy3, Array("A", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(100, "A") & vbCrLf & String(100, "X"))
PushI SampDy3, Array("B", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(100, "A") & vbCrLf & String(100, "X"))
End Property

Property Get SampDy() As Variant()
PushI SampDy, Array("A", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "A"))
PushI SampDy, Array("B", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "B"))
PushI SampDy, Array("C", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "C"))
PushI SampDy, Array("D", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "D"))
PushI SampDy, Array("E", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "E"))
PushI SampDy, Array("F", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "F"))
PushI SampDy, Array("G", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "G"))
End Property

Property Get SampDs() As Ds
AddDt SampDs, SampDt1
AddDt SampDs, SampDt2
SampDs.DsNm = "Ds"
End Property

Property Get SampDt1() As Dt
SampDt1 = DtzFF("SampDt1", "A B C", SampDy1)
End Property

Property Get SampDt2() As Dt
SampDt2 = DtzFF("SampDt2", "A B C", SampDy2)
End Property
