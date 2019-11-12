Attribute VB_Name = "MxAddCol"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CNs$ = "Dta.AddColl"
Const CMod$ = CLib & "MxAddCol."
Function AddCol(A As Drs, C$, V) As Drs
Dim Dr, Dy()
For Each Dr In Itr(A.Dy)
    PushI Dr, V
    PushI Dy, Dr
Next
AddCol = AddColzFFDy(A, C, Dy)
End Function

Function AddColzDyCC(Dy(), V1, V2) As Variant()
AddColzDyCC = AddColzDyAv(Dy, Av(V1, V2))
End Function

Function AddColzDyC3(Dy(), V1, V2, V3) As Variant()
AddColzDyC3 = AddColzDyAv(Dy, Av(V1, V2, V3))
End Function

Function AddColz2(A As Drs, FF2$, C1, C2) As Drs
Dim Fny$(), Dy()
Fny = AddAy(A.Fny, TermAy(FF2))
Dy = AddColzDyCC(A.Dy, C1, C2)
AddColz2 = Drs(Fny, Dy)
End Function

Function AddColz3(A As Drs, FF3$, C1, C2, C3) As Drs
Dim Fny$(), Dy()
Fny = AddAy(A.Fny, TermAy(FF3))
Dy = AddColzDyC3(A.Dy, C1, C2, C3)
AddColz3 = Drs(Fny, Dy)
End Function

Function AddColzDyColVy(Dy(), ColVy, AtIx&) As Variant()
Dim Dr, J&, O(), U&
U = UB(ColVy)
If U = -1 Then Exit Function
If U <> UB(Dy) Then Thw CSub, "Row-in-Dy <> Si-ColVy", "Row-in-Dy Si-ColVy", Si(Dy), Si(ColVy)
ReDim O(U)

For Each Dr In Itr(Dy)
    If Si(Dr) > AtIx Then Thw CSub, "Some Dr in Dy has bigger size than AtIx", "DrSz AtIx", Si(Dr), AtIx
    ReDim Preserve Dr(AtIx)
    Dr(AtIx) = ColVy(J)
    O(J) = Dr
    J = J + 1
Next
AddColzDyColVy = O
End Function

