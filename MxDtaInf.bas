Attribute VB_Name = "MxDtaInf"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaInf."
Public Const vbFldSep$ = ""

Function ColzDrs(A As Drs, ColNm$) As Variant()
ColzDrs = ColzDy(A.Dy, IxzAy(A.Fny, ColNm))
End Function

Function VzColEq(A As Drs, SelC$, Col$, Eq)
Dim Dr, I&
I = IxzAy(A.Fny, Col)
For Each Dr In Itr(A.Dy)
    If Dr(I) = Eq Then VzColEq = Dr(IxzAy(A.Fny, SelC))
Next
End Function

Function WdtzCol%(A As Drs, C$)
WdtzCol = WdtzAy(StrCol(A, C))
End Function

Function JnDyCC(Dy(), CCIxy&(), Optional FldSep$ = vbFldSep) As String()
Dim Dr
For Each Dr In Itr(Dy)
    PushI JnDyCC, Jn(AeIxy(Dr, CCIxy), FldSep)
Next
End Function

Function StrColzDy(Dy(), C) As String()
StrColzDy = IntozDyC(EmpSy, Dy, C)
End Function

Function DblColzDy(Dy(), C) As Double()
DblColzDy = IntozDyC(EmpDblAy, Dy, C)
End Function

Function StrColzDyFst(Dy()) As String()
StrColzDyFst = StrColzDy(Dy, 0)
End Function

Function StrColzDySnd(Dy()) As String()
StrColzDySnd = StrColzDy(Dy, 1)
End Function

Function BoolColzDy(Dy(), C&) As Boolean()
BoolColzDy = IntozDyC(EmpBoolAy, Dy, C)
End Function


Function IntCol(A As Drs, C) As Integer()
IntCol = IntColzDy(A.Dy, IxzAy(A.Fny, C))
End Function

Function IntColzDy(Dy(), C) As Integer()
IntColzDy = IntozDyC(EmpIntAy, Dy, C)
End Function

Function Col(D As Drs, C) As Variant()
Col = ColzDy(D.Dy, IxzAy(D.Fny, C))
End Function

Function ColzDy(Dy(), C) As Variant()
ColzDy = IntozDyC(EmpAv(), Dy, C)
End Function

Function IntozDyC(Into, Dy(), C)
Dim O, J&, Dr, U&
O = Into
U = UB(Dy)
O = ResiU(O, U)
For Each Dr In Itr(Dy)
    If UB(Dr) >= C Then
        O(J) = Dr(C)
    End If
    J = J + 1
Next
IntozDyC = O
End Function

Function HasReczDy2V(Dy(), C1, C2, V1, V2) As Boolean
Dim Dr
For Each Dr In Itr(Dy)
    If Dr(C1) = V1 Then
        If Dr(C2) = V2 Then
            HasReczDy2V = True
            Exit Function
        End If
    End If
Next
End Function

Function IsSamNCol(A As Drs, NCol%) As Boolean
Dim Dr
For Each Dr In Itr(A.Dy)
    If Si(Dr) = NCol Then Exit Function
Next
IsSamNCol = True
End Function

Function ResiDrs(A As Drs, NCol%) As Drs
If IsSamNCol(A, NCol) Then ResiDrs = A: Exit Function
Dim O As Drs, U%, Dr, J%
U = NCol - 1
For J = 0 To UB(O.Dy)
    Dr = O.Dy(J)
    ReDim Preserve Dr(U)
    O.Dy(J) = Dr
Next
End Function

Function LngAyzColEqSel(A As Drs, C$, V, Sel$) As Long()
LngAyzColEqSel = LngAyzDrs(DwEQSel(A, C, V, Sel), Sel)
End Function

Function LngAyzDrs(A As Drs, C$) As Long()
LngAyzDrs = IntozDrsC(EmpLngAy, A, C)
End Function

Function LngAyzDyC(Dy(), C) As Long()
LngAyzDyC = IntozDyC(EmpLngAy, Dy, C)
End Function
