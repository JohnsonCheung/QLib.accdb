Attribute VB_Name = "MxDtaCol"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaCol."

Function AddColzDy(Dy(), ValToBeAddAsLasCol) As Variant()
'Ret : a new :Dy with a col of value all eq to @ValToBeAddAsLasCol at end
Dim O(): O = Dy
Dim ToU&
    ToU = NColzDy(Dy)
Dim J&, Dr
For Each Dr In Itr(O)
    ReDim Preserve Dr(ToU)
    Dr(ToU) = ValToBeAddAsLasCol
    O(J) = Dr
    J = J + 1
Next
AddColzDy = O
End Function

Function AddColzDyAv(Dy(), Av()) As Variant()
Dim O(): O = Dy
Dim ToU&
    ToU = NColzDy(Dy) + 1
Dim J&, Dr, I1%, I2%
I2 = ToU
I1 = I2 - 1
For Each Dr In Itr(O)
    ReDim Preserve Dr(ToU)
    PushAy Dr, Av
    O(J) = Dr
    J = J + 1
Next
AddColzDyAv = O
End Function

Function AddColzDyBy(Dy(), Optional ByNCol% = 1) As Variant()
Dim NewU&
    NewU = NColzDy(Dy) + ByNCol - 1
Dim O()
    Dim UDy&: UDy = UB(Dy)
    O = AyReSzU(O, UDy)
    Dim J&
    For J = 0 To UDy
        O(J) = AyReSzU(Dy(J), NewU)
    Next
AddColzDyBy = O
End Function

Function AddColzDyC(Dy(), C) As Variant()
If Si(Dy) = 0 Then Exit Function
Dim O(): O = AddColzDyBy(Dy)
    Dim UCol%: UCol = UB(Dy(0))
    Dim J&
    For J = 0 To UB(Dy)
       O(J)(UCol) = C
    Next
AddColzDyC = O
End Function

Function AddColzMap(A As Drs, NewFldEqFunQteFmFldSsl$) As Drs
Dim NewColVy(), FmVy()
Dim I, S$, NewFld$, Fun$, FmFld$
For Each I In SyzSS(NewFldEqFunQteFmFldSsl)
    S = I
    NewFld = Bef(S, "=")
    Fun = IsBet(S, "=", "(")
    FmFld = BetBkt(S)
    FmVy = ColzDrs(A, FmFld)
    NewColVy = MapAy(FmVy, Fun)
    Stop '
Next
End Function

Function AddColzVy(A As Drs, ColNm$, ColVy) As Drs
Dim Fny$(): Fny = AddEle(A.Fny, ColNm)
Dim AtIx&: AtIx = UB(Fny)
Dim Dy(): Dy = AddColzDyColVy(A.Dy, ColVy, AtIx)
AddColzVy = Drs(Fny, Dy)
End Function

Function CntColEq&(A As Drs, C$, V)
Dim I%: I = IxzAy(A.Fny, C)
Dim O&, Dr: For Each Dr In Itr(A.Dy)
    If Dr(I) = V Then O = O + 1
Next
CntColEq = O
End Function

Function CntColNe&(A As Drs, C$, V)
Dim I%: I = IxzAy(A.Fny, C)
Dim O&, Dr: For Each Dr In Itr(A.Dy)
    If Dr(I) <> V Then O = O + 1
Next
CntColNe = O
End Function
Function ColNoSng(A As Drs, C$) As Drs
'@A : has a column-C
'Ret   : sam stru as A and som row removed.  rmv row are its col C value is Single. @@
Dim Col(): Col = ColzDrs(A, C)
Dim Sng(): Sng = AwSng(Col)
ColNoSng = DeIn(A, C, Sng)
End Function

Function FstDr(A As Drs, C$, V) As Variant()
Dim Ix&: Ix = IxzAy(A.Fny, C)
FstDr = A.Dy(Ix)
End Function

Function FstDrSel(A As Drs, C$, V, Sel$) As Variant()
FstDrSel = AwIxy(FstDr(A, C, V), IxyzFF(A.Fny, Sel))
End Function

Function FstRec(A As Drs, C$, V) As Drs
FstRec = Drs(A.Fny, FstDr(A, C, V))
End Function

Function HasColEq(A As Drs, C$, V) As Boolean
HasColEq = HasColEqzDy(A.Dy, IxzAy(A.Fny, C), V)
End Function

Function InsColzDyAv(Dy(), Av()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    PushI InsColzDyAv, AddAy(Av, Dr)
Next
End Function

Function InsColzDy(A(), V, Optional At& = 0) As Variant()
Dim Dr
For Each Dr In Itr(A)
    PushI InsColzDy, InsEle(Dr, V, At)
Next
End Function

Function InsColzDyV2(A(), V1, V2) As Variant()
InsColzDyV2 = InsColzDyAv(A, Av(V1, V2))
End Function

Function InsColzDyV3(Dy(), V1, V2, V3) As Variant()
InsColzDyV3 = InsColzDyAv(Dy, Av(V1, V2, V3))
End Function

Function InsColzDyV4(A(), V1, V2, V3, V4) As Variant()
InsColzDyV4 = InsColzDyAv(A, Av(V1, V2, V3, V4))
End Function

Function RmvPfxzDrs(A As Drs, C$, Pfx$) As Drs
Dim Dr, ODy(), J&, I%
ODy = A.Dy
I = IxzAy(A.Fny, C)
For Each Dr In Itr(A.Dy)
    Dr(I) = RmvPfx(Dr(I), Pfx)
    ODy(J) = Dr
    J = J + 1
Next
RmvPfxzDrs = Drs(A.Fny, ODy)
End Function

Function RxyeDyVy(Dy(), Vy) As Long()
'Fm Dy: ! to be selected if it ne to @Vy.  It has sam NCol as Si-Vy
'Fm Vy : ! to select @Dy
'Ret   : Rxy of @Dy if the rec ne @Vy
Dim Rix&, Dr: For Each Dr In Itr(Dy)
    If Not IsEqAy(Dr, Vy) Then PushI RxyeDyVy, Rix
    Rix = Rix + 1
Next
End Function

Function RxywDyVy(Dy(), Vy) As Long()
'Fm Dy: ! to be selected if it eq to @Vy.  It has sam NCol as Si-Vy
'Fm Vy : ! to select @Dy
'Ret   : Rxy of @Dy if the rec eq @Vy
Dim Rix&, Dr: For Each Dr In Itr(Dy)
    If IsEqAy(Dr, Vy) Then PushI RxywDyVy, Rix
    Rix = Rix + 1
Next
End Function

Function DwTopN(A As Drs, Optional N = 50) As Drs
If N <= 0 Then DwTopN = A: Exit Function
DwTopN = Drs(A.Fny, CvAv(FstNEle(A.Dy, N)))
End Function

Function ValzDrs(A As Drs, C$, V, ColNm$)
Dim Dr, Ix%, IxRet%
Ix = IxzAy(A.Fny, C)
IxRet = IxzAy(A.Fny, ColNm)
For Each Dr In Itr(A.Dy)
    If Dr(Ix) = V Then
        ValzDrs = Dr(IxRet)
        Exit Function
    End If
Next
Thw CSub, "In Drs, there is no record with Col-A eq Value-B, so no Col-C is returened", "Col-A Value-B Col-C Drs-Fny Drs-NRec", C, V, ColNm, A.Fny, NReczDrs(A)
End Function
