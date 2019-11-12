Attribute VB_Name = "MxAddMthCol"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxAddMthCol."

Function Add5MthCol(Wi_MthLin As Drs) As Drs
Dim ITyChr   As Drs:   ITyChr = AddMthColTyChr(Wi_MthLin)
Dim IPm      As Drs:      IPm = AddMthColMthPm(ITyChr)
Dim IShtPm   As Drs:   IShtPm = AddMthColShtPm(IPm)
Dim IRetAs   As Drs:   IRetAs = AddMthColRetAs(IShtPm)
                   Add5MthCol = AddMthColRmk(IRetAs)
End Function

Function AddMthColHasPm(WiMthLin As Drs) As Drs
'Ret         : Add col-HasPm
Dim I%: I = IxzAy(WiMthLin.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In Itr(WiMthLin.Dy)
    Dim MthLin$: MthLin = Dr(I)
    Dim HasPm As Boolean: HasPm = BetBkt(MthLin) <> ""
    PushI Dr, HasPm
    PushI Dy, Dr
Next
AddMthColHasPm = AddColzFFDy(WiMthLin, "HasPm", Dy)
End Function

Function AddMthColTyChr(WiMthLin As Drs) As Drs
'Ret         : Add col-HasPm
Dim I%: I = IxzAy(WiMthLin.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In Itr(WiMthLin.Dy)
    Dim MthLin$: MthLin = Dr(I)
    Dim TyChr$: TyChr = MthTyChr(MthLin)
    PushI Dr, TyChr
    PushI Dy, Dr
Next
AddMthColTyChr = AddColzFFDy(WiMthLin, "TyChr", Dy)
End Function

Function AddMthColShtPm(WiMthPm As Drs) As Drs
'Ret         : Add col-ShtPm
Dim I%: I = IxzAy(WiMthPm.Fny, "MthPm")
Dim Dr, Dy(): For Each Dr In Itr(WiMthPm.Dy)
    Dim MthPm$: MthPm = Dr(I)
    Dim ShtPm1$: ShtPm1 = ShtPm(MthPm)
    PushI Dr, ShtPm1
    PushI Dy, Dr
Next
AddMthColShtPm = AddColzFFDy(WiMthPm, "ShtPm", Dy)
End Function

Function AddMthColMthPm(WiMthLin As Drs, Optional IsDrp As Boolean) As Drs
AddMthColMthPm = AddColzBetBkt(WiMthLin, "MthLin:MthPm", IsDrp)
End Function

Function AddMthColIsRetObj(WiRetSfx As Drs) As Drs
'Fm WiRetSfx : ..RetSfx..
'Ret         : ..IsRetObj @@
Dim IxRetAs%: IxRetAs = IxzAy(WiRetSfx.Fny, "RetSfx")
Dim Dr, Dy(): For Each Dr In Itr(WiRetSfx.Dy)
    Dim RetSfx$: RetSfx = Dr(IxRetAs)
    Dim R As Boolean: R = IsRetObj(RetSfx)
    PushI Dr, R
    PushI Dy, Dr
Next
AddMthColIsRetObj = AddColzFFDy(WiRetSfx, "IsRetObj", Dy)
End Function
