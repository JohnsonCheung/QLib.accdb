Attribute VB_Name = "MxSimTy"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxSimTy."
Enum EmSimTy
    EiUnk
    EiEmp
    EiYes
    EiNum
    EiDte
    EiStr
End Enum

Function SimTy(V) As EmSimTy
SimTy = SimTyzV(VarType(V))
End Function

Function SimTyzCol(Col()) As EmSimTy
Dim V: For Each V In Itr(Col)
    Dim O As EmSimTy: O = MaxSim(O, SimTy(V))
    If O = EiStr Then SimTyzCol = O: Exit Function
Next
End Function

Function SimTyzLo(L As ListObject) As EmSimTy()
Dim Sq(): Sq = SqzLo(L)
Dim C%: For C = 1 To UBound(Sq, 2)
    PushI SimTyzLo, SimTyzCol(ColzSq(Sq, C))
Next
End Function

Function SimTyzV(V As VbVarType) As EmSimTy
Dim O As EmSimTy
Select Case True
Case V = Empty: O = EiEmp
Case V = vbBoolean: O = EiYes
Case V = vbByte, V = vbCurrency, V = vbDecimal, V = vbDouble, V = vbInteger, V = vbLong, V = vbSingle: O = EiNum
Case V = vbDate: O = EiDte
Case V = vbString: O = EiStr
End Select
SimTyzV = O
End Function

