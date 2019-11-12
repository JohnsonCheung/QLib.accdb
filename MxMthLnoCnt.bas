Attribute VB_Name = "MxMthLnoCnt"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMthLnoCnt."
'*LnoCnt:Fun|Sub has one StartLineNo|Count.  Prp may have 2.
Type Lnocnt2: S1 As Long: C1 As Long: S2 As Long: C2 As Long: End Type
Type Lnocnt: S As Long: C As Long: End Type

Function MthLnoC(M As CodeModule, MthLno&) As Lnocnt
If MthLno = 0 Then Thw CSub, "MthLno cannot be zero"
With MthLnoC
    .S = MthLno
    If .S = 0 Then Exit Function
    Dim A&: A = EndLnozM(M, MthLno)
    .C = A - .S + 1: If .C <= 0 Then Thw CSub, FmtQQ("MthLineCnt[?] cannot be 0 or neg", .C)
End With
End Function

Function FmtLnoC2$(A As Lnocnt2)
With A
FmtLnoC2 = FmtQQ("LnoCnt(? ? ? ? ?)", .S1, .C1, "|", .S2, .C2)
End With
End Function

Function MthLnoC2(M As CodeModule, Mthn) As Lnocnt2
Dim A&(): A = MthLnoAy(M, Mthn)
Dim O As Lnocnt2
Select Case Si(A)
Case 0
Case 1: GoSub X1
Case 2: GoSub X1: GoSub X2
Case Else: ThwNever CSub, "There is error in MthLnoC, it should return 0,1 or 2 Lno"
End Select
MthLnoC2 = O
Exit Function
X1:
    With MthLnoC(M, A(0))
    O.C1 = .C
    O.S1 = .S
    End With
    Return
X2:
    With MthLnoC(M, A(1))
    O.C2 = .C
    O.S2 = .S
    End With
    Return
End Function
