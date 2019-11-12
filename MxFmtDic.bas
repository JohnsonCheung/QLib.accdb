Attribute VB_Name = "MxFmtDic"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxFmtDic."
Sub Z_BrwDic()
Dim R As DAO.Recordset
Set R = Rs(SampDbDutyDta, "Select Sku,BchNo from PermitD where BchNo<>''")
BrwDic JnStrDicTwoFldRs(R), True
End Sub

Sub VcDic(A As Dictionary, Optional InclValTy As Boolean, Optional ExlIx As Boolean)
BrwDic A, InclValTy, ExlIx, OupTy:=EiOtVc
End Sub

Sub BrwDic(A As Dictionary, Optional InclValTy As Boolean, Optional ExlIx As Boolean, Optional OupTy As EmOupTy = EmOupTy.EiOtBrw)
BrwAy FmtDic(A, InclValTy), OupTy:=OupTy
End Sub

Sub DmpDic(A As Dictionary, Optional InclDicValOptTy As Boolean, Optional Tit$ = "Key Val")
D FmtDic(A, InclDicValOptTy, Tit)
End Sub

Function S12szDiT1qLy(A As Dictionary) As S12s
Dim K: For Each K In A.Keys
    PushS12 S12szDiT1qLy, S12(K, JnCrLf(A(K)))
Next
End Function

Function FmtDicTit(A As Dictionary, Tit$) As String()
PushI FmtDicTit, Tit
PushI FmtDicTit, vbTab & "Count=" & A.Count
PushIAy FmtDicTit, AmAddPfx(FmtDic(A, InclValTy:=True), vbTab)
End Function

Function FmtDic(A As Dictionary, Optional InclValTy As Boolean, Optional FF$ = "Key Val", Optional IxCol As EmIxCol) As String()
ThwIf_Nothing A, "Dic", CSub
Select Case True
Case IsDicSy(A):    FmtDic = FmtS12s(S12szDiT1qLy(A), FF, IxCol)
Case IsDicLines(A): FmtDic = FmtS12s(S12szDic(A), FF, IxCol)
Case Else:          FmtDic = FmtDiczLin(A, " ", InclValTy, FF)
End Select
End Function

Function FmtDiczLin(A As Dictionary, Optional Sep$ = " ", Optional InclValTy As Boolean, Optional FF$ = "Key Val") As String()
If A.Count = 0 Then Exit Function
Dim Key: Key = A.Keys
Dim N1$, N2$: AsgAp SyzSS(FF), N1, N2
Dim O$(): O = ItmAddAy(N1, SyzItr(A.Keys))
O(0) = O(0) + Sep + N2
Dim J&, I
For Each I In A.Items
   O(J) = O(J) & Sep & I
   J = J + 1
Next
FmtDiczLin = O
End Function
