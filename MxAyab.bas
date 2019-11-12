Attribute VB_Name = "MxAyab"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAyab."
Type Ayabc
    A As Variant
    B As Variant
    C As Variant
End Type
Type Ayab
    A As Variant
    B As Variant
End Type
Function Ayab(A, B) As Ayab
ThwIf_NotAy A, CSub
ThwIf_NotAy B, CSub
With Ayab
    .A = A
    .B = B
End With
End Function

Function Ayabc(A, B, C) As Ayabc
ThwIf_NotAy A, CSub
ThwIf_NotAy B, CSub
ThwIf_NotAy C, CSub
With Ayabc
    .A = A
    .B = B
    .C = C
End With
End Function

Function AyabzAyPfx(Ay, Pfx$) As Ayab
Dim O As Ayab
O.A = ResiU(Ay)
O.B = O.A
Dim S$, I
For Each I In Itr(Ay)
    S = I
    If HasPfx(S, Pfx) Then
        PushI O.B, S
    Else
        PushI O.A, S
    End If
Next
AyabzAyPfx = O
End Function

Function AyabzAyN(Ay, N&) As Ayab
AyabzAyN = Ayab(FstNEle(Ay, N), AeFstNEle(Ay, N))
End Function

Function AyabczAyFE(Ay, FmIx&, EIx&) As Ayabc
Dim O As Ayabc
AyabczAyFE = Ayabc( _
    AwFE(Ay, 0, FmIx), _
    AwFE(Ay, FmIx, EIx), _
    AwFm(Ay, EIx))
End Function
Function AyabJn(A, B, Sep$) As String()
Dim J&: For J = 0 To Min(UB(A), UB(B))
    PushI AyabJn, A(J) & Sep & B(J)
Next
End Function
Function AyabJnDot(A, B) As String()
AyabJnDot = AyabJn(A, B, ".")
End Function
Function AyabJnSngQ(A, B) As String()
AyabJnSngQ = AyabJn(A, B, "'")
End Function
Function AyabczAyFei(Ay, B As Fei) As Ayabc
AyabczAyFei = AyabczAyFE(Ay, B.FmIx, B.EIx)
End Function


Function DyoAyab(A, B) As Variant()
Dim J&
For J = 0 To Min(UB(A), UB(B))
    PushI DyoAyab, Array(A(J), B(J))
Next
End Function
Function DrszAyab(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2") As Drs
DrszAyab = Drs(Sy(N1, N2), DyoAyab(A, B))
End Function




Function LyzAyab(AyA, AyB, Optional Sep$) As String()
ThwIf_DifSi AyA, AyB, CSub
Dim A, J&: For Each A In Itr(AyA)
    PushI LyzAyab, A & Sep & AyB(J)
    J = J + 1
Next
End Function

Function LyzAyabSpc(AyA, AyB) As String()
LyzAyabSpc = LyzAyab(AyA, AyB, " ")
End Function

Function FmtAyab(A, B, Optional FF$ = "Ay1 Ay2") As String()
FmtAyab = FmtS12s(S12szAyab(A, B), FF)
End Function

Function LyzAyabNEmpB(A, B, Optional Sep$ = " ") As String()
Dim J&, O$()
For J = 0 To UB(A)
    If Not IsEmp(B(J)) Then
        Push O, A(J) & Sep & B(J)
    End If
Next
LyzAyabNEmpB = O
End Function

Sub AsgAyaReSzMax(A, B, OA, OB)
OA = A
OB = B
ResiMax OA, OB
End Sub

Function DyzAyab(A, B) As Variant()
ThwIf_DifSi A, B, CSub
Dim I, J&: For Each I In Itr(A)
    PushI DyzAyab, Array(I, B(J))
    J = J + 1
Next
End Function
