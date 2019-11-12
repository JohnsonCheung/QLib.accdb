Attribute VB_Name = "MxAddMthCol"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxAddMthCol."

Function Add5MthCol(Wi_MthLin As Drs) As Drs
Dim ITyChr   As Drs:     ITyChr = AddMthColTyChr(Wi_MthLin)
Dim IPm      As Drs:        IPm = AddMthColMthPm(ITyChr)
Dim IShtPm   As Drs:     IShtPm = AddMthColShtPm(IPm)
Dim IRetAs   As Drs:     IRetAs = AddMthColRetAs(IShtPm)
Add5MthCol = AddMthColIsRetObj(IRetAs)
End Function

Function Add6MthCol(Wi_MthLin_Mthl As Drs) As Drs
Add6MthCol = AddMthColRmk(Add5MthCol(Wi_MthLin_Mthl))
End Function

Function AddMthColTyChr(Wi_MthLin As Drs) As Drs
'Ret         : Add col-HasPm
Dim I%: I = IxzAy(Wi_MthLin.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In Itr(Wi_MthLin.Dy)
    Dim MthLin$: MthLin = Dr(I)
    Dim TyChr$: TyChr = MthTyChr(MthLin)
    PushI Dr, TyChr
    PushI Dy, Dr
Next
AddMthColTyChr = AddColzFFDy(Wi_MthLin, "TyChr", Dy)
End Function

Function AddMthColShtPm(Wi_MthPm As Drs) As Drs
'Ret         : Add col-ShtPm
Dim I%: I = IxzAy(Wi_MthPm.Fny, "MthPm")
Dim Dr, Dy(): For Each Dr In Itr(Wi_MthPm.Dy)
    Dim MthPm$: MthPm = Dr(I)
    Dim ShtPm1$: ShtPm1 = ShtPm(MthPm)
    PushI Dr, ShtPm1
    PushI Dy, Dr
Next
AddMthColShtPm = AddColzFFDy(Wi_MthPm, "ShtPm", Dy)
End Function

Function AddMthColMthPm(Wi_MthLin As Drs, Optional IsDrp As Boolean) As Drs
AddMthColMthPm = AddColzBetBkt(Wi_MthLin, "MthLin:MthPm", IsDrp)
End Function

Function AddMthColIsRetObj(Wi_RetAs As Drs) As Drs
'@Wi_RetAs :Drs..RetAs..
'Ret       :Drs..IsRetObj @@
Dim IxRetAs%: IxRetAs = IxzAy(Wi_RetAs.Fny, "RetAs")
Dim Dr, Dy(): For Each Dr In Itr(Wi_RetAs.Dy)
    Dim RetAs$: RetAs = Dr(IxRetAs)
    Dim R As Boolean: R = IsRetObj(RetAs)
    PushI Dr, R
    PushI Dy, Dr
Next
AddMthColIsRetObj = AddColzFFDy(Wi_RetAs, "IsRetObj", Dy)
End Function
Function AddMthColRetAs(Wi_MthLin As Drs) As Drs
Dim I%: I = IxzAy(Wi_MthLin.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In Itr(Wi_MthLin.Dy)
    Dim MthLin$: MthLin = Dr(I)
    Dim Ret$: Ret = RetAs(MthLin)
    PushI Dr, Ret
    PushI Dy, Dr
Next
AddMthColRetAs = AddColzFFDy(Wi_MthLin, "RetAs", Dy)
End Function

Function AddMthColRmk(Wi_Mthl As Drs) As Drs
Dim I%: I = IxzAy(Wi_Mthl.Fny, "Mthl")
Dim Dr, Dy(): For Each Dr In Itr(Wi_Mthl.Dy)
    Dim Mthl$: Mthl = Dr(I)
    PushI Dr, MthRmkzMthLy(SplitCrLf(Mthl))
    PushI Dy, Dr
Next
AddMthColRmk = AddColzFFDy(Wi_Mthl, "Rmk", Dy)
End Function

Function IsRetObj(RetSfx$) As Boolean
':IsRetObj: :B ! False if @RetSfx (isBlnk | IsAy | IsPrimTy | Is in TyNyP)
If RetSfx = "" Then Exit Function
If HasSfx(RetSfx, "()") Then Exit Function
If IsPrimTy(RetSfx) Then Exit Function
If HasEle(TyNyP, RetSfx) Then Exit Function
IsRetObj = True
End Function


