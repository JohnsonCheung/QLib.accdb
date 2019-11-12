Attribute VB_Name = "MxGp"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxGp."

Function GpAsAyDy(D As Drs, Gpcc$) As Variant()
'@D : ..{Gpcc}..        ! it has col-@Gpcc
'Ret   : gp-of-:Dy:-of-@D ! each gp of dry has same @Gpcc val
Dim O()
    Dim SDy():  SDy = D.Dy              ' the src-dy to be gp
    Dim K As Drs: K = SelDrs(D, Gpcc)   ' a drs fm @D with grouping columns only
    Dim G():      G = GRxy(K.Dy)        ' gp-of-rxy-pointing-to-@D-row
    Dim Rxy: For Each Rxy In Itr(G)     ' Rxy is gp-of-rix-poiting-to-@D-row
        Dim ODy(): Erase ODy            ' Gp-of-@D-row with sam val of @Gpcc
        Dim Rix: For Each Rix In Rxy    ' Rix is rix-poiting-to-@D-row
            PushI ODy, SDy(Rix)         ' Push the @D-row to @ODy
        Next
        PushI O, ODy                    ' <-- put to @O, the output
    Next
GpAsAyDy = O
End Function

Function Gp(D As Drs, Gpcc$, Optional C$) As Drs
'@D : ..@Gpcc..@C.. ! it has col-Gpcc and optional col-C
'Ret   : @Gpcc #Gp ! where #Gp is opt gp of col-C, in :Av: @@
Dim OKey(), OGp()  ' Sam Si
    Dim SDy(): SDy = D.Dy               ' #Src-Dy. Source Dy to be gp
    Dim K As Drs: K = SelDrs(D, Gpcc)   ' #Key.    Only those @Gpcc column
    Dim IxGp%: IxGp = Si(K.Fny)         '          The column to put gp
    Dim KDr: For Each KDr In Itr(K.Dy)
        Dim Ix&: Ix = IxzDyDr(OKey, KDr)
        If Ix = -1 Then
            PushI OGp, Array(SDy(Ix))
            PushI OKey, KDr
        Else
            PushI OGp(Ix), SDy(Ix)
        End If
    Next
Dim ODy()
Dim GpDr, J&: For Each GpDr In Itr(OKey)  ' For Each ele-of-OKey put corresponding ele-of-OGp at end to form a gp-rec
    PushI GpDr, OGp(J)               ' Put OGp(J) at end of #GpDr, now GpDr is a gp-rec
    PushI ODy, GpDr
    J = J + 1
Next
Gp = AddColzFFDy(D, "Gp", ODy)
End Function

Function GRxyzCy(Dy(), Cxy&()) As Variant()
'Fm Dy : #Dta-Row-arraY# ! Dy to be gp.  It has all col as stated in @Cxy.
'Fm Cxy : #Col-Ix-Array# ! Gpg which col of @Dy
'Ret    : Ay-of-Dy.  Each ele is a subset of @Dy in same gp.  @@
GRxyzCy = GRxy(SelDy(Dy, Cxy)) ' sel the gp-ing col and gp it
End Function

Function GRxy(Dy()) As Variant()
'Fm Dy :              ! to be gp.  All col will be used to gp.
'Ret    : #Gp-of-Rxy#  ! ##Gp-Of-Row-Index-Array##.  Each gp is an ay of rix.  each rix is pointing to a @Dy rec.  All the rec in a gp will have sam val.
Dim I&, Key(), O(), Dr, R&: For Each Dr In Itr(Dy)
    I = IxzDyDr(Key, Dr)
    If I = -1 Then
        PushI Key, Dr
        PushI O, LngAy(R)
    Else
        PushI O(I), R
    End If
    R = R + 1
Next
GRxy = O
End Function


Function AgrMin(D As Drs, Gpcc$, MinC$) As Drs
Dim Dy()
    Dim A As Drs: A = Gp(D, Gpcc, MinC)
    Dim Dr: For Each Dr In Itr(A.Dy)
        Dim Col(): Col = Pop(Dr)
        PushI Dr, AyMin(Col)
        PushI Dy, Dr
    Next
AgrMin = Drs(D.Fny, Dy)
End Function

Function AgrMax(D As Drs, Gpcc$, MaxC$) As Drs
Dim Dy()
    Dim A As Drs: A = Gp(D, Gpcc, MaxC)
    Dim Dr: For Each Dr In Itr(A.Dy)
        Dim Col(): Col = Pop(Dr)
        PushI Dr, AyMax(Col)
        PushI Dy, Dr
    Next
AgrMax = Drs(D.Fny, Dy)
End Function

Function Agr(D As Drs, Gpcc$, ArgColn$) As Drs
Dim Dy()
    Dim A As Drs: A = Gp(D, Gpcc, ArgColn)
    Dim Dr: For Each Dr In Itr(A.Dy)
        Dim Col(): Col = Pop(Dr)
        Dim Sum#: Sum = AySum(Col)
        Dim N&:   N = Si(Col)
        Dim Avg: If N <> 0 Then Avg = Sum / N
        PushI Dr, N
        PushI Dr, Avg
        PushI Dr, AySum(Col)
        PushI Dr, AyMin(Col)
        PushI Dr, AyMax(Col)
        PushI Dy, Dr
    Next
Dim NewFny$(): NewFny = SyzSS(RplQ("?Cnt ?Avg ?Sum ?Min ?Max", ArgColn))
Dim Fny$(): Fny = AddSy(D.Fny, NewFny)
Agr = Drs(Fny, Dy)
End Function

Sub Z_AgrWdt()
BrwDrs AgrWdt(DoPubFun, "Mdn Ty", "Mthn")
End Sub

Function AgrWdt(D As Drs, Gpcc$, C$) As Drs
Dim A As Drs: A = Gp(D, Gpcc, C)
Dim Dr, Dy(): For Each Dr In Itr(A.Dy)
    Dim Col(): Col = Pop(Dr)
    PushI Dr, WdtzAy(Col)
    PushI Dy, Dr
Next
Dim Fny$(): Fny(UB(Fny)) = "W" & C
AgrWdt = Drs(Fny, Dy)
End Function
