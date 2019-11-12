Attribute VB_Name = "MxDic"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDic."

Function DoDic(A As Dictionary, Optional InclDicValOptTy As Boolean, Optional Tit$ = "Key Val") As Drs
DoDic = Drs(FoDic(InclDicValOptTy), DyzDi(A, InclDicValOptTy))
End Function

Function DtzDic(A As Dictionary, Optional DtNm$ = "Dic", Optional InclDicValOptTy As Boolean) As Dt
Dim Dy()
Dy = DyzDi(A, InclDicValOptTy)
Dim F$
    If InclDicValOptTy Then
        F = "Key Val Ty"
    Else
        F = "Key Val"
    End If
DtzDic = DtzFF(DtNm, F, Dy)
End Function

Function FoDic(Optional InclValTy As Boolean) As String()
FoDic = SyzSS("Key Val" & IIf(InclValTy, " Type", ""))
End Function
