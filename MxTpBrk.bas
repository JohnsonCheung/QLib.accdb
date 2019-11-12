Attribute VB_Name = "MxTpBrk"
Option Compare Text
Option Explicit
Const CLib$ = "QTp."
Const CMod$ = CLib & "MxTpBrk."
':TyPfx:Ro: #Record-Of#
Type RoBlk
    BlkTy As String
    SepLin As String
    DyoLLin() As Variant
    BlkIx As Integer
    SepLno As Long
End Type
Type Blks
    D As Drs ' BlkTy SepLin DyoLLin:<Lno Lin> BlkIx SepLno ' 1-Dr-is-1-Blk
End Type
Type Blk
    DroBlk() As Variant ' BlkTy SepLin Dyo<Lno Lin>  BlkIx SepLno
End Type
Type DyoLLin
    Dy() As Variant ' Av<Lno Lin>
End Type
Type TpBrk
    BlkTyAyoMul() As String
    BlkTyAyoSng() As String
    RmkLy() As String ' Those Ly above fst SepLin
    Ok As Blks        ' Those Blks in *BlkTyAy
    Er As Blks        ' Those Blks not in *BlkTyAy
    Excess As Blks    ' Those Blks in *BlkTyAyoSng, but excess
End Type
Public Const FFoBlks$ = "BlkTy SepLin DyoLLin BlkIx SepLno"
':SepLin: :Lin ! is a lin with Pfx of [:SepLinPfx & " " & One-Of-@BlkAyTy & " "]
    
Function TpBrk(Tp$, Mul_BlkTyss$, Sng_BlkTyss$, Optional SepLinPfx$ = "==") As TpBrk
'Fm Mul_BlkTyss : these blk ty alw mul, in a tp more than 1 is alw
'Fm Sng_BlkTyss : these blk ty is sng.  2nd and rest blk is Excess
'Ret :TpBrk     : brk @Tp by using @BlkTyss @BlkTyssoSng
Dim Mul$():                     Mul = SyzSS(Mul_BlkTyss)                     ' BlkTyAyoMul
Dim Sng$():                     Sng = SyzSS(Sng_BlkTyss)                     ' BlkTyAyoSng
Dim TpLy$():                   TpLy = SplitCrLf(Tp)
Dim Rmk$():                     Rmk = RmkLy(TpLy, SepLinPfx)                 ' :RmkLy, the ly above fst
Dim NoRmkLno%:             NoRmkLno = Si(Rmk)                                ' Lno of fst lin in                                   #Rmk
Dim NoRmk$():                 NoRmk = AwFm(TpLy, NoRmkLno)                   ' Ly-of-NoRmk.  The fst-ele-NoRmk should be a :SepLin
Dim Blks As Blks:              Blks = BlkszNoRmk(NoRmkLno, NoRmk, SepLinPfx) ' Brk NoRmk into :Blks
                  TpBrk.BlkTyAyoMul = Mul
                  TpBrk.BlkTyAyoSng = Sng
                        TpBrk.RmkLy = Rmk
                           TpBrk.Er = BlkswEr(Blks, Mul, Sng)
                           TpBrk.Ok = BlkswOk(Blks, Mul, Sng)
                       TpBrk.Excess = BlkswExcess(Blks, Mul)
End Function

Function RmkLy(TpLy$(), SepLinPfx$) As String()
'Ret : :RmkLy from @TpLy using @BlkTyAy @@
':RmkLy: :Ly ! Fst chr of :Ly is [']
Dim L, O$(): For Each L In TpLy
    If HasPfx(L, SepLinPfx) Then ' when True, means it a :SepLin
        RmkLy = O
        Exit Function
    End If
    PushI O, L
Next
RmkLy = O
Stop
End Function

Function XDroBlks(BlkTy$, SepLin$, DyoLLin(), BlkIx%, SepLno%) As Variant()
XDroBlks = Array(BlkTy, SepLin, DyoLLin, BlkIx, SepLno)
End Function

Function FoBlks() As String()
FoBlks = SyzSS(FFoBlks)
End Function

Function BlkszNoRmk(NoRmkLno%, NoRmk$(), SepLinPfx$) As Blks
'Fm NoRmkLno : is the Lno of fst-ele-of-@NoRmk, which should be a :SepLin
'Fm NoRmk    : is :TpLy after rmv :RmkLy
'Ret         : :Blks<BlkTy SepLin DyoLLin:Dyo<Lno Lin> BlkIx SepLno
Dim ODy() 'DyoBlks
Dim Fst As Boolean: Fst = True
Dim BlkTy$, SepLin$, DyoLLin(), BlkIx%, SepLno%, Dr(), Lno%, DroBlks()
Lno = NoRmkLno
Dim L: For Each L In Itr(NoRmk)
    If HasPfx(L, SepLinPfx) Then
        SepLin = L
        BlkTy = BlkTyzSepLin(SepLin, SepLinPfx)
        If Fst Then
            Fst = False
        Else
            PushI ODy, DroBlks
        End If
        DroBlks = XDroBlks(BlkTy, SepLin, EmpAv, BlkIx, Lno)
        BlkIx = BlkIx + 1
    Else
        PushI DroBlks(2), Array(Lno, L)
    End If
    Lno = Lno + 1
Next
BlkszNoRmk = Blks(ODy)
End Function

Function BlkTyzSepLin$(SepLin$, SepLinPfx$)
'Ret : the xxx of [== xxxx ] is BlkTy.  ie, Aft @SepLinPfx Bet 2 spc
Dim L$: L = RmvPfx(SepLin, SepLinPfx)
If FstChr(L) <> " " Then Exit Function
L = RmvFstChr(L)
Dim P%: P = InStr(L, " "): If P = 0 Then Exit Function
BlkTyzSepLin = Left(L, P - 1)
End Function

Function BlkswEr(B As Blks, Mul$(), Sng$()) As Blks
'Fm Mul : :BlkTyAy which alw more than 1 blk
'Fm Sng : only alw 1 blk
'Ret    : those blk in @B not in neither
Dim Dy(): Dy = B.D.Dy
Dim BTy$(): BTy = AddSy(Mul, Sng)
Dim Dr, ODy(): For Each Dr In Itr(Dy)
    If Not HasEle(BTy, RoBlkTy(Blk(Dr))) Then PushI ODy, Dr
Next
BlkswEr = Blks(ODy)
End Function

Function BlkswExcess(B As Blks, Sng$()) As Blks
'Fm Sng : sng blk instance blk ty.
'Ret : subset of @B which is excess blk.  For those blk type is @Sng, the 2nd and rest blk is treated as excess.
Dim Dy(): Dy = B.D.Dy
Dim Fnd$()
Dim ODy(), Dr: For Each Dr In Itr(Dy)
    Dim BlkTy$:  BlkTy = RoBlkTy(Blk(Dr))
    If HasEle(Sng, BlkTy) Then
        If HasEle(Fnd, BlkTy) Then
            PushI ODy, Dr
        Else
            PushI Fnd, BlkTy
        End If
    End If
Next
BlkswExcess = Blks(ODy)
End Function

Function BlkswOk(B As Blks, BtoMul$(), BtoSng$()) As Blks
'Ret : subset of @B which is ok.  It is ok when the #BlkTy is :Mul or is (:Sng and fst blk)
Dim Dy(): Dy = B.D.Dy
Dim BtoSngDone$()
Dim ODy(), Dr: For Each Dr In Itr(Dy)
    Dim BlkTy$:  BlkTy = RoBlkTy(Blk(Dr))
    Dim IsOk As Boolean: IsOk = False
    Select Case True
    Case HasEle(BtoMul, BlkTy): IsOk = True     '<-- Is Ok
    Case HasEle(BtoSng, BlkTy)
        If Not HasEle(BtoSngDone, BlkTy) Then
            IsOk = True                         '<-- Is Ok
            PushI BtoSngDone, BlkTy
        End If
    End Select
    If IsOk Then PushI ODy, Dr      '<==
Next
BlkswOk = Blks(ODy)
End Function

Function Blks(DyoBlk()) As Blks
'Fm DyoBlk: Dyo<BlkTy SepLin Dyo<Lno Lin> BlkIx SepLno>
Blks.D = Drs(FoBlks, DyoBlk)
End Function
Sub Z_TpBrk()
Dim Tp$, SepLinPfx$
GoSub Z
Exit Sub
Z:
    Tp = SampSqTp
    SepLinPfx = "=="
    Dim B As TpBrk: B = TpBrk(Tp, "SQ", "PM SW")
    Vc FmtTpBrk(B)
    Return
End Sub

Function FmtTpBrk(A As TpBrk) As String()
Dim B As New Bfr
With B
.Box "TpBrk"
.Lin
.ULin "BlkTyoMul: [" & JnSpc(A.BlkTyAyoMul) & "]"
.ULin "BlkTyoSng: [" & JnSpc(A.BlkTyAyoSng) & "]"
.Lin
.Box "Rmk", "&"
.Var A.RmkLy
.Lin
.Box "Ok Blks", "&"
.Var FmtBlks(A.Ok)
.Lin
.Box "Er Blks", "&"
.Var FmtBlks(A.Er)
End With
FmtTpBrk = B.Ly
End Function

Sub BrwBlks(B As Blks)
Brw FmtBlks(B)
End Sub

Sub VcBlks(B As Blks)
Vc FmtBlks(B)
End Sub

Function CntBlk%(B As Blks, BlkTy$)
Dim O&
Dim DroBlk: For Each DroBlk In B.D.Dy
    If DroBlk(0) = BlkTy Then O = O + 1
Next
CntBlk = O
End Function

Function DyoLLinzBlk(B As Blk) As Variant()
DyoLLinzBlk = B.DroBlk(2)
End Function

Function FmtBlks(B As Blks) As String()
Dim DroBlk, J&: For Each DroBlk In B.D.Dy
    PushIAy FmtBlks, FmtBlk(Blk(DroBlk), J)
Next
End Function

Function Blk(DroBlk) As Blk
'@DroBlk :Av-BlkTy-SepLin-Dyo<Lno Lin> BlkIx SepLno>
If Si(DroBlk) <> 5 Then Stop
Blk.DroBlk = DroBlk
End Function

Function FmtLLin(DyoLLin()) As String()
FmtLLin = AmRTrim(AlignDyAsLy(DyoLLin))
End Function

Function RoBlkIx%(B As Blk)
RoBlkIx = B.DroBlk(3)
End Function

Function RoBlkTy%(B As Blk)
RoBlkTy = B.DroBlk(0)
End Function

Function RoBlkSepLin$(B As Blk)
RoBlkSepLin = B.DroBlk(1)
End Function

Function RoBlkSepLno&(B As Blk)
RoBlkSepLno = B.DroBlk(4)
End Function

Function RoBlk(B As Blk) As RoBlk
Dim Dr(): Dr = B.DroBlk
With RoBlk
    .BlkTy = Dr(0)
    .SepLin = Dr(1)
    .DyoLLin = Dr(2)
    .BlkIx = Dr(3)
    .SepLno = Dr(4)
End With
End Function

Function FmtBlk(B As Blk, Optional BlkIx = -1) As String()
'Fm Blk : Dro<BlkTy SepLin Dyo<L Lin>>
Dim P$: If BlkIx >= 0 Then P = "(BlkIx:" & BlkIx & ") "
Dim RoB As RoBlk: RoB = RoBlk(B)
With RoB
PushI FmtBlk, FmtQQ("**BlkIx(?) BlkTy(?) SepLno(?) SepLin(?)", .BlkIx, .BlkTy, .SepLno, .SepLin)
PushIAy FmtBlk, FmtLLin(.DyoLLin)
End With
End Function

Function BlkzTy(B As Blks, BlkTy$, Optional FmIx = 0) As Blk
Dim Dy(): Dy = B.D.Dy
Dim Dr$(), Ix%: For Ix = FmIx To UB(Dy) ' Dro<BlkTy SepLin Dyo<Lno Lin> BlkIx>
    Dr = Dy(Ix)
    If Dr(0) = BlkTy Then
        BlkzTy = Blk(Dr): Exit Function
    End If
Next
End Function

Function LyzBlkTy(B As Blks, BlkTy$) As String()
LyzBlkTy = StrColzDySnd(DyoLLinzBlk(BlkzTy(B, BlkTy)))
End Function

Function LyAyzBlkTy(B As Blks, BlkTy$, Optional FmIx% = 0) As Variant()
Dim D As Drs: D = DwEQ(B.D, "BlkTy", BlkTy)
Dim Dr: For Each Dr In D.Dy
    Dim DyoBlk(): DyoBlk = Dr(2)  ' Av<BlkTy SepLin Dyo<Lno Lin>>
    PushI LyAyzBlkTy, StrColzDySnd(DyoBlk)
Next
End Function

Function IsBlkEmp(B As Blk) As Boolean
IsBlkEmp = Si(B.DroBlk) = 0
End Function

Function EoBlkTy(B As Blks, BlkTy$, Optional FmIx = 0) As String()
'Ret: Those Blk @FmIx and eq @BlkTy in @B are considered as error.  Upd them as :ErLy @@
Dim Dy(): Dy = B.D.Dy
Dim Ix%: For Ix = FmIx To UB(Dy)
    PushIAy EoBlkTy, EoBlk(Blk(Dy(Ix)))
Next
End Function

Function EoBlk(Er As Blk) As String()
Stop
End Function

Function NxtBlk(B As Blks, M As Blk) As Blk
'Ret : Nxt :Blk of @M fm @B.  M.BlkIx
NxtBlk = BlkzTy(B, RoBlkTy(M), BlkIx(M) + 1)
End Function

Function EoAftBlk(B As Blks, Aft As Blk) As String()
'Ret : :ErLy of all :Blk @Aft in @B of-sam-ty-as-@Aft which are considered as er.  `Aft` means those blks in @B aft @Aft.BlkIx
Dim Dy(): Dy = B.D.Dy
Dim TyAft$: TyAft = RoBlkTy(Aft) ' The-Ty-of-@Aft
Dim Ix%: For Ix = BlkIx(Aft) + 1 To UB(Dy)
    Dim Dr: Dr = Dy(Ix)
    Dim TyCur$: TyCur = RoBlkTy(Blk(Dr))
    If TyAft = TyCur Then
        PushIAy EoAftBlk, EoBlk(Blk(Dy(Ix)))
    End If
Next
End Function
