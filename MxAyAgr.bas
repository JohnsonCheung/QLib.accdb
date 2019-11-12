Attribute VB_Name = "MxAyAgr"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAyAgr."


Function LcAgrP() As Drs
LcAgrP = AgrzNum(LinCntP)
End Function

Function LinCntP() As Long()
LinCntP = LinCntzP(CPj)
End Function

Function LinCntzP(P As VBProject) As Long()
Dim C As VBComponent: For Each C In P.VBComponents
    PushI LinCntzP, C.CodeModule.CountOfLines
Next
End Function

Function CntNo0&(NumAy)
Dim O&
Dim V: For Each V In Itr(NumAy)
    If V <> 0 Then O = O + 1
Next
CntNo0 = O
End Function

Function AgrzNum(NumAy) As Drs
'Ret : Agr Val ! where *Arg has Cnt Avg Max Min Sum
Dim ODy()
Dim Sum#: Sum = AySum(NumAy)
Dim NNo0&: NNo0 = CntNo0(NumAy)
Dim N&: N = Si(NumAy)
Dim AvgAll#, AvgNo0#
If N <> 0 Then AvgAll = Sum / N
If NNo0 <> 0 Then AvgNo0 = Sum / NNo0

Push ODy, Array("CntNo0", NNo0)
Push ODy, Array("CntAll", N)
Push ODy, Array("AvgNo0", AvgNo0)
Push ODy, Array("AvgAll", AvgAll)
Push ODy, Array("Sum", Sum)
Push ODy, Array("Max", AyMax(NumAy))
Push ODy, Array("Min", AyMin(NumAy))
Push ODy, Array("MinGT0", AyMinzGT0(NumAy))
AgrzNum = DrszFF("Agr Val", ODy)
End Function
