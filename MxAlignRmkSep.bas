Attribute VB_Name = "MxAlignRmkSep"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxAlignRmkSep."
Function SelRmkSep(Wi_MthLin As Drs) As Drs
'@Wi_MthLin : MthLin #Mth-Context.
'Ret : select where LTrim-*MthLin has pfx '-- '== or '..
Dim IxMthLin%: IxMthLin = IxzAy(Wi_MthLin.Fny, "L")
Dim Dr, Dy(): For Each Dr In Itr(Wi_MthLin.Dy)
    Dim L$: L = LTrim(Dr(IxMthLin))
    If FstChr(L) = "'" Then
        L = Left(RmvFstChr(L), 2)
        Select Case L
        Case "==", "--", "..": PushI Dy, Dr
        End Select
    End If
Next
SelRmkSep.Fny = Wi_MthLin.Fny
SelRmkSep.Dy = Dy
End Function

Function AlignRmkSepzD(Wi_L_MthLin As Drs) As Drs
'@Wi_L_MthLin De : L MthLin ! Where MthLin is {spc}'-- '== '..
'Ret   : L NewL OldL        ! Where NewL is aligned with 120 @@
Dim IxL%, IxMthLin%: AsgIx Wi_L_MthLin, "L MthLin", IxL, IxMthLin
Dim Dr, Dy(): For Each Dr In Itr(Wi_L_MthLin.Dy)
    Dim L&:       L = Dr(IxL)
    Dim OldL$: OldL = Dr(IxMthLin)
    Dim C$:       C = Mid(LTrim(OldL), 2, 1)
    Dim NewL$: NewL = Left(OldL, 120) & Dup(C, 120 - Len(OldL))
    If OldL <> NewL Then
        Push Dy, Array(L, NewL, OldL)
    End If
Next
AlignRmkSepzD = LNewO(Dy)
'Insp "QIde_B_AlignMth.XDeLNewO", "Inspect", "Oup(XDeLNewO) De", FmtDrszNoRdu(XDeLNewO), FmtDrszNoRdu(De): Stop
End Function

Function AlignRmkSep(Upd As Boolean, M As CodeModule, Wi_L_MthLin As Drs) As Drs
Dim D As Drs:   D = SelRmkSep(Wi_L_MthLin)
Dim D1 As Drs: D1 = AlignRmkSepzD(D)
:                    If Upd Then RplLin M, D ' <==
End Function
