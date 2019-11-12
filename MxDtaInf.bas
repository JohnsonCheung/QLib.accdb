Attribute VB_Name = "MxDtaInf"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaInf."
Public Const vbFldSep$ = ""

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
LngAyzColEqSel = LngAyzDrs(DwEqSel(A, C, V, Sel), Sel)
End Function

Function LngAyzDrs(A As Drs, C$) As Long()
LngAyzDrs = IntozDrsC(EmpLngAy, A, C)
End Function

Function LngAyzDyC(Dy(), C) As Long()
LngAyzDyC = IntozDyC(EmpLngAy, Dy, C)
End Function
