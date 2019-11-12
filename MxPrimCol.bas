Attribute VB_Name = "MxPrimCol"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxPrimCol."

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

Function FstStrCol(A As Drs) As String()
FstStrCol = StrColzDy(A.Dy, 0)
End Function

Function SndStrCol(A As Drs) As String()
SndStrCol = StrColzDy(A.Dy, 1)
End Function

Function StrCol(A As Drs, C) As String()
StrCol = StrColzDy(A.Dy, IxzAy(A.Fny, C))
End Function

Function LngCol(A As Drs, C) As Long()
LngCol = LngColzDy(A.Dy, IxzAy(A.Fny, C))
End Function

Function StrColLines$(A As Drs, C)
StrColLines = JnCrLf(StrCol(A, C))
End Function

Function DblCol(A As Drs, C) As Double()
DblCol = DblColzDy(A.Dy, IxzAy(A.Fny, C))
End Function

Function BoolCol(A As Drs, C) As Boolean()
BoolCol = BoolColzDy(A.Dy, IxzAy(A.Fny, C))
End Function

Function FstColzDy(Dy()) As Variant()
FstColzDy = ColzDy(Dy, 0)
End Function

Function Fst3ColzDy(Dy()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    PushI Fst3ColzDy, FstNEle(Dr, 3)
Next
End Function

Function FstCol(A As Drs) As Variant()
FstCol = FstColzDy(A.Dy)
End Function

Function StrColzEq(A As Drs, Col$, V, ColNm$) As String()
Dim B As Drs
B = DwEqSel(A, Col, V, ColNm)
StrColzEq = StrCol(B, ColNm)
End Function

Function ColzDrs(A As Drs, ColNm$) As Variant()
ColzDrs = ColzDy(A.Dy, IxzAy(A.Fny, ColNm))
End Function

Function StrColzDy(Dy(), C) As String()
StrColzDy = IntozDyC(EmpSy, Dy, C)
End Function

Function LngColzDy(Dy(), C) As Long()
LngColzDy = IntozDyC(EmpLngAy, Dy, C)
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
