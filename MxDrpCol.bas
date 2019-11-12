Attribute VB_Name = "MxDrpCol"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDrpCol."

Function DrpCol(A As Drs, CC$) As Drs
Dim C$(), Dr, Ixy&(), OFny$(), ODy()
C = SyzSS(CC)
Ixy = IxyzSubAy(A.Fny, C)
OFny = AyMinus(A.Fny, C)
ODy = DrpColzDy(A.Dy, CvLngAy(AySrt(Ixy)))
DrpCol = Drs(OFny, ODy)
End Function

Function DrpColzDy(Dy(), Ixy&()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    PushI DrpColzDy, AeIxy(Dr, Ixy)
Next
End Function

Function DrpColzDyIxy(Dy(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
   Push DrpColzDyIxy, AeIxy(Dr, Ixy)
Next
End Function

