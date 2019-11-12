Attribute VB_Name = "MxAyab"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAyab."
Type Ayabc
    A As Variant
    b As Variant
    C As Variant
End Type
Type Ayab
    A As Variant
    b As Variant
End Type
Function Ayab(A, b) As Ayab
ThwIf_NotAy A, CSub
ThwIf_NotAy b, CSub
With Ayab
    .A = A
    .b = b
End With
End Function

Function Ayabc(A, b, C) As Ayabc
ThwIf_NotAy A, CSub
ThwIf_NotAy b, CSub
ThwIf_NotAy C, CSub
With Ayabc
    .A = A
    .b = b
    .C = C
End With
End Function

Function AyabzAyPfx(Ay, Pfx$) As Ayab
Dim O As Ayab
O.A = ResiU(Ay)
O.b = O.A
Dim S$, I
For Each I In Itr(Ay)
    S = I
    If HasPfx(S, Pfx) Then
        PushI O.b, S
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
Function AyabJn(A, b, Sep$) As String()
Dim J&: For J = 0 To Min(UB(A), UB(b))
    PushI AyabJn, A(J) & Sep & b(J)
Next
End Function
Function AyabJnDot(A, b) As String()
AyabJnDot = AyabJn(A, b, ".")
End Function
Function AyabJnSngQ(A, b) As String()
AyabJnSngQ = AyabJn(A, b, "'")
End Function
Function AyabczAyFei(Ay, b As Fei) As Ayabc
AyabczAyFei = AyabczAyFE(Ay, b.FmIx, b.EIx)
End Function


Function DyoAyab(A, b) As Variant()
Dim J&
For J = 0 To Min(UB(A), UB(b))
    PushI DyoAyab, Array(A(J), b(J))
Next
End Function
Function DrszAyab(A, b, Optional N1$ = "Ay1", Optional N2$ = "Ay2") As Drs
DrszAyab = Drs(Sy(N1, N2), DyoAyab(A, b))
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

Function FmtAyab(A, b, Optional FF$ = "Ay1 Ay2") As String()
FmtAyab = FmtS12s(S12szAyab(A, b), FF)
End Function

Function LyzAyabNEmpB(A, b, Optional Sep$ = " ") As String()
Dim J&, O$()
For J = 0 To UB(A)
    If Not IsEmp(b(J)) Then
        Push O, A(J) & Sep & b(J)
    End If
Next
LyzAyabNEmpB = O
End Function

Sub AsgAyaReSzMax(A, b, OA, OB)
OA = A
OB = b
ResiMax OA, OB
End Sub

Function DyzAyab(A, b) As Variant()
ThwIf_DifSi A, b, CSub
Dim I, J&: For Each I In Itr(A)
    PushI DyzAyab, Array(I, b(J))
    J = J + 1
Next
End Function
