Attribute VB_Name = "MxDoDic"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDoDic."

Function DoDic(A As Dictionary, Optional InclValTy As Boolean, Optional Tit$ = "Key Val") As Drs
DoDic = Drs(FoDic(InclValTy), DyzDi(A, InclValTy))
End Function

Function FoDic(Optional InclValTy As Boolean) As String()
FoDic = SyzSS("Key Val" & IIf(InclValTy, " Type", ""))
End Function
