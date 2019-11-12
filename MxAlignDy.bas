Attribute VB_Name = "MxAlignDy"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxAlignDy."

Function AlignDrWAsLin$(Dr, WdtAy%())
'Ret : a lin by joing [ | ] and quoting [| * |] after aligng @Dr with @WdtAy. @@
AlignDrWAsLin = QteJnzAsTLin(AlignDrzW(Dr, WdtAy))
End Function

Function AlignDrzWAsLy(Ay, WdtAy%()) As String()
Dim S, J&: For Each S In Ay
    PushI AlignDrzWAsLy, Align(S, WdtAy(J))
    J = J + 1
Next
End Function

Function AlignSqzW(Sq(), W%()) As Variant()
Dim O(): O = Sq
Dim IC%: For IC = 1 To UBound(Sq, 2)
    Dim Wdt%: Wdt = W(IC - 1)
    Dim IR&: For IR = 1 To UBound(Sq, 1)
        O(IR, IC) = Align(O(IR, IC), Wdt)
    Next
Next
AlignSqzW = O
End Function

Function AlignSq(Sq()) As Variant()
If Si(Sq) = 0 Then Exit Function
Dim C&, O(), NR&, NC&
NR = UBound(Sq, 1)
NC = UBound(Sq, 2)
ReDim O(1 To NR, 1 To NC)
For C = 1 To UBound(O, 2)
    AlignColzSq O, Sq, C, WdtzSqc(Sq, C)
Next
AlignSq = O
End Function

Function AlignDrzW(Dr, WdtAy%())
Dim O: O = Dr
Dim J%: For J = 0 To Min(UB(Dr), UB(WdtAy))
    O(J) = AlignL(Dr(J), WdtAy(J))
Next
AlignDrzW = O
End Function

Sub AlignDyzCol(ODy(), C)
'Fm ODy : the col @C will be aligned
'Fm C    : the column ix
'Ret     : column-@C of @ODy will be aligned
Dim Col(): Col = ColzDy(ODy, C)
Dim ACol$(): ACol = AlignAy(Col)
Dim J&: For J = 0 To UB(ODy)
    ODy(J)(C) = ACol(J)
Next
End Sub


Function AlignDyzCix(Dy(), Cix&()) As Variant()
Dim O(): O = Dy
Dim C: For Each C In Cix
    AlignDyzCol O, C
Next
AlignDyzCix = O
End Function

Function AlignDy(Dy(), Optional FstNTerm%) As Variant()
Dim W%(): W = WdtAyzDy(Dy, FstNTerm)
AlignDy = AlignDyzW(Dy, W)
End Function

Function AlignDyAsLy(Dy()) As String()
AlignDyAsLy = JnDy(AlignDy(Dy))
End Function

Function AlignDyzW(Dy(), FstNTermWdtAy%()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    PushI AlignDyzW, AlignDrzW(Dr, FstNTermWdtAy)
Next
End Function
