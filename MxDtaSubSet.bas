Attribute VB_Name = "MxDtaSubSet"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaSubSet."

Function DeIn(A As Drs, C, InVy) As Drs
Dim Ix&: Ix = IxzAy(A.Fny, C)
If Not IsArray(InVy) Then Thw CSub, "Given InVy is not an array", "Ty-InVy", TypeName(InVy)
Dim Dr, Dy(): For Each Dr In Itr(A.Dy)
    If Not HasEle(InVy, Dr(Ix)) Then
        PushI Dy, Dr
    End If
Next
DeIn = Drs(A.Fny, Dy)
End Function

Function DeRxy(A As Drs, Rxy&()) As Drs
Dim ODy(), Dy()
    Dy = A.Dy
    Dim J&, I&
    For J = 0 To UB(Dy)
        If Not HasEle(Rxy, J) Then
            PushI ODy, Dy(J)
        End If
    Next
DeRxy = Drs(A.Fny, ODy)
End Function

Function DeVap(D As Drs, CC$, ParamArray Vap()) As Drs
'Fm D : ..@CC..            ! to be selected.  It has col-@CC
'Fm Vap : #Val-Ay-of-Param ! to select what rec in @D to be selected
'Ret    : ..@D..           ! sam stru as @D.  Subset of @D.  @@
Dim Vy(): Vy = Vap
DeVap = DeVy(D, CC, Vy)
End Function

Function DeVy(D As Drs, CC$, Vy) As Drs
'Fm D  : ..@CC..            ! to be selected.  It has col-@CC
'Fm Vy : #Val-Ay-of-Param ! to select what rec in @D to be selected
'Ret   : ..@D..           ! sam stru as @D.  Subset of @D.  @@
Dim KeyDy(): KeyDy = SelDrs(D, CC).Dy
Dim Rxy&(): Rxy = RxyeDyVy(KeyDy, Vy)
Dim ODy(): ODy = AwIxy(D.Dy, Rxy)
DeVy = Drs(D.Fny, ODy)
End Function
Function Dw2Eq(D As Drs, C2$, V1, V2) As Drs
Dim A$, b$: AsgTRst C2, A, b
Dw2Eq = DwEq(DwEq(D, A, V1), b, V2)
End Function
Function DwNBlnk(D As Drs, C$) As Drs
DwNBlnk = DwNe(D, C, "")
End Function

Function Dw2EqE(D As Drs, C2$, V1, V2) As Drs
Dw2EqE = DrpCol(Dw2Eq(D, C2, V1, V2), C2)
End Function

Function Dw2Patn(A As Drs, TwoC$, Patn1$, Patn2$) As Drs
Dim C1$, C2$: AsgBrkSpc TwoC, C1, C2
Dw2Patn = DwPatn(DwPatn(A, C1, Patn1), C2, Patn2)
End Function

Function Dw3Eq(D As Drs, C3$, V1, V2, V3) As Drs
Dim A$, b$, C$: AsgTTRst C3, A, b, C
Dw3Eq = DwEq(DwEq(DwEq(D, A, V1), b, V2), C, V3)
End Function

Function Dw3EqE(D As Drs, C3$, V1, V2, V3) As Drs
Dw3EqE = DrpCol(Dw3Eq(D, C3, V1, V2, V3), C3)
End Function

Function DwColGt(A As Drs, C$, V) As Drs
Dim Dy(), Ix%, Fny$()
Fny = A.Fny
'Ix = Ixy(Fny, C)
DwColGt = Drs(Fny, DywColGt(A.Dy, Ix, V))
End Function

Function DwColNe(A As Drs, C$, V) As Drs
Dim Dy(), Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C)
DwColNe = Drs(Fny, DywColNe(A.Dy, Ix, V))
End Function

Function DwDup(D As Drs, DupFF$) As Drs
DwDup = F_SubDrs_BySubRxy(D, F_SubRxy_ByDupFF(D, DupFF))
End Function

Function DwDupC(A As Drs, C$) As Drs
Dim Dup(): Dup = AwDup(ColzDrs(A, C))
DwDupC = DwIn(A, C, Dup)
End Function

Function DwBlnk(A As Drs, C$) As Drs
DwBlnk = DwEq(A, C, "")
End Function

Function F_SubDrs_ByNBlnkC(A As Drs, NBlnkC$) As Drs
F_SubDrs_ByNBlnkC = DwNe(A, NBlnkC, "")
End Function

Function DwCC2Eq(D As Drs, CC2$, EqV1, EqV2) As Drs
Dim C1$, C2$: AsgBrkSpc CC2, C1, C2
DwCC2Eq = DwEq(DwEq(D, C1, EqV1), C2, EqV2)
End Function

Function DwCC2EqExl(D As Drs, CC2$, EqV1, EqV2) As Drs
DwCC2EqExl = DrpCol(DwCC2Eq(D, CC2, EqV1, EqV2), CC2)
End Function

Function DwEq(A As Drs, ByC$, Eq) As Drs
Dim Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, ByC, ThwEr:=EiThwEr)
DwEq = Drs(Fny, DywEq(A.Dy, Ix, Eq))
End Function

Function DwEqStr(A As Drs, C$, Str$) As Drs
If Str = "" Then DwEqStr = A: Exit Function
DwEqStr = DwEq(A, C, Str)
End Function

Function F_SubDrs_ByC_SubStr(A As Drs, C$, SubStr) As Drs
Dim Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C, ThwEr:=EiThwEr)
F_SubDrs_ByC_SubStr = Drs(Fny, DywSubStr(A.Dy, Ix, SubStr))
End Function

Function F_SubDrs_ByC_Lik(A As Drs, C$, Lik) As Drs
Dim Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C, ThwEr:=EiThwEr)
F_SubDrs_ByC_Lik = Drs(Fny, DywLik(A.Dy, Ix, Lik))
End Function

Function DwFalse(A As Drs, C$) As Drs
DwFalse = DwEq(A, C, False)
End Function
Function DwFalseExl(A As Drs, C$) As Drs
DwFalseExl = DrpCol(DwFalse(A, C), C)
End Function

Function DwEqExl(A As Drs, C$, V) As Drs
DwEqExl = DrpCol(DwEq(A, C, V), C)
End Function

Function DwEqSel(A As Drs, C$, V, Sel$) As Drs
DwEqSel = SelDrs(DwEq(A, C, V), Sel)
End Function

Function DwFFNe(A As Drs, F1, F2) As Drs 'FFNe = Two Fld Not Eq
Dim Fny$()
Fny = A.Fny
'DwFFNe = Drs(Fny, DyWhCCNe(A.Dy, Ixy(Fny, F1), Ixy(Fny, F2)))
End Function

Function DwFldEqV(A As Drs, F, EqVal) As Drs
'DwFldEqV = Drs(A.Fny, DyWh(A.Dy, Ixy(A.Fny, F), EqVal))
End Function

Function DwIn(A As Drs, C, InVy) As Drs
Dim Ix&: Ix = IxzAy(A.Fny, C)
DwIn = Drs(A.Fny, DywIn(A.Dy, Ix, InVy))
End Function

Function DwNe(A As Drs, C$, V) As Drs
DwNe = DwColNe(A, C, V)
End Function

Function DwNeSel(A As Drs, C$, V, Sel$) As Drs
DwNeSel = SelDrs(DwNe(A, C, V), Sel)
End Function

Function DwNotRxy(A As Drs, Rxy&()) As Drs
Dim O(), Dy()
    Dy = A.Dy
    Dim J&
    For J = 0 To UB(Dy)
        If Not HasEle(Rxy, J) Then
            Push O, Dy(J)
        End If
    Next
DwNotRxy = Drs(A.Fny, O)
End Function

Function DwPatn(A As Drs, C$, Patn$, Optional Patn1$, Optional Patn2$) As Drs
If Patn = "" And Patn1 = "" And Patn2 = "" Then DwPatn = A: Exit Function
Dim I%: I = IxzAy(A.Fny, C, ThwEr:=EiThwEr)
Dim P As IPred: Set P = PredHasPatn(Patn, Patn1, Patn2)
Dim Dy(), Dr: For Each Dr In Itr(A.Dy)
    If P.Pred(Dr(I)) Then PushI Dy, Dr
Next
DwPatn = Drs(A.Fny, Dy)
End Function

Function DwPatnSS3(A As Drs, C$, PatnSS3$) As Drs
Dim P$, P1$, P2$
AsgPatn123 PatnSS3, P, P1, P2
DwPatnSS3 = DwPatn(A, C, P, P1, P2)
End Function

Function DwPfx(D As Drs, C$, Pfx) As Drs
DwPfx = Drs(D.Fny, DywPfx(D.Dy, IxzAy(D.Fny, C), Pfx))
End Function

Function F_SubDrs_BySubRxy(A As Drs, Rxy&()) As Drs
Dim Dy(): Dy = AwIxy(A.Dy, Rxy)
F_SubDrs_BySubRxy = Drs(A.Fny, Dy)
End Function

Function DwTop(A As Drs, Optional NTop& = 50) As Drs
DwTop = Drs(A.Fny, CvAv(FstNEle(A.Dy, NTop)))
End Function

Function DwVap(D As Drs, CC$, ParamArray Vap()) As Drs
'Fm D : ..@CC..            ! to be selected.  It has col-@CC
'Fm Vap : #Val-Ay-of-Param ! to select what rec in @D to be returned
'Ret    : ..@D..           ! sam stru as @D.  Subset of @D.  @@
Dim Vy(): Vy = Vap
Dim KeyDy(): KeyDy = SelDrs(D, CC).Dy
Dim Rxy&(): Rxy = RxywDyVy(KeyDy, Vy)
Dim ODy(): ODy = AwIxy(D.Dy, Rxy)
DwVap = Drs(D.Fny, ODy)
End Function

Function DwEqFny(A As Drs, C$, V, SelFny$()) As Drs
DwEqFny = SelDrsFny(DwEq(A, C, V), SelFny)
End Function

Function DwInSel(A As Drs, C, InVy, Sel$) As Drs
DwInSel = SelDrs(DwIn(A, C, InVy), Sel)
End Function
