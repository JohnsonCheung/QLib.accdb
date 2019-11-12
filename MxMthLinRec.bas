Attribute VB_Name = "MxMthLinRec"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMthLinRec."
Type MthLinRec
    ShtMdy As String
    ShtTy As String
    Nm As String
    TyChr As String
    RetTy As String
    Pm As String
    Rmk As String
    IsRetVal As Boolean
    ShtRetTy As String
End Type
Function MthLinRec(MthLin) As MthLinRec
Dim L$: L = MthLin
With MthLinRec
    .ShtMdy = ShfShtMdy(L)
    .ShtTy = ShfShtMthTy(L): If .ShtTy = "" Then Exit Function
    .Nm = ShfNm(L)
    .TyChr = ShfTyChr(L)
    .Pm = ShfBktStr(L)
    .RetTy = ShfRetTyzAftPm(L)
    .Rmk = MthLinRmkzAftRetTy(L)
    .IsRetVal = HasEle(SyzSS("Get Fun"), .ShtTy)
    .ShtRetTy = ShtRetTy(.TyChr, .RetTy, .IsRetVal)
End With
End Function

