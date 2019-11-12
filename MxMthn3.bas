Attribute VB_Name = "MxMthn3"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMthn3."
Type Mthn3: Nm As String: ShtTy As String: ShtMdy As String: End Type
':Mthn-ZDash: :Mthn-Rul ! All mth beg with Z_XXX is n rule.  When beg with Z_, it is trying to test.  Try use own resource, like Y_.  Don't use other md resource."
':Mthn-ZZDash :Mthn-Rul ! All mth with XXX__YYY means sub-mth.  XXX is ParNm and YYY is ChdNm
':Mthn-Z:     :Mthn-Rul ! It must be Sub.  CdZ
':Mthn-YDash: :Mthn-Rul ! PurePrp for testing.  Used to Z_ mth."

Function Mthn3(Nm, ShtMdy, ShtTy) As Mthn3
With Mthn3
    .Nm = Nm
    .ShtMdy = ShtMdy
    .ShtTy = ShtTy
End With
End Function

Function Mthn3zL(Lin) As Mthn3
Mthn3zL = ShfMthn3(CStr(Lin))
End Function


Function ShfMthn3(OLin$) As Mthn3
Dim M$: M = ShfShtMdy(OLin)
Dim T$: T = ShfShtMthTy(OLin):: If T = "" Then Exit Function
ShfMthn3 = Mthn3(ShfNm(OLin), M, T)
End Function

Function RmvMthn3$(Lin)
Dim L$: L = Lin
RmvMthMdy L
If ShfMthTy(L) = "" Then Exit Function
If ShfNm(L) = "" Then Thw CSub, "Not as SrcLin", "Lin", Lin
RmvMthn3 = L
End Function

Function FmtMthn3$(A As Mthn3)
With A
FmtMthn3 = JnDotAp(.Nm, .ShtMdy, .ShtTy)
End With
End Function
Sub DmPubMthn3(A As Mthn3)
D FmtMthn3(A)
End Sub
