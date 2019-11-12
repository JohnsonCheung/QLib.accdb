Attribute VB_Name = "MxDtaAy"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaAy."

Sub BrwGRxyzCyCntzAy(Ay)
Brw JnGRxyzCyCntzAy(Ay)
End Sub

Function DtzAy(Ay, Optional FldNm$ = "Itm", Optional DtNm$ = "Ay") As Dt
Dim Dy(), J&
For J = 0 To UB(Ay)
    PushI Dy, Array(Ay(J))
Next
DtzAy = Dt(DtNm, Sy(FldNm), Dy)
End Function

Function GRxyzCyCntzAy(Ay) As Variant()
If Si(Ay) = 0 Then Exit Function
Dim Dup, O(), X, T&, Cnt&
Dup = AwDup(Ay)
For Each X In Itr(Dup)
    Cnt = AyItmCnt(Ay, X)
    Push O, Array(X, AyItmCnt(Ay, X))
    T = T + Cnt
Next
Push O, Array("~Tot", T)
GRxyzCyCntzAy = O
End Function

Function GRxyzCyCntzAyWhDup(A) As Variant()
GRxyzCyCntzAyWhDup = DywColGt(GRxyzCyCntzAy(A), 1, 1)
End Function

Function JnGRxyzCyCntzAy(Ay) As String()
JnGRxyzCyCntzAy = JnDy(GRxyzCyCntzAy(Ay))
End Function

Sub Z_JnGRxyzCyCntzAy()
Dim Ay()
Brw JnGRxyzCyCntzAy(Ay)
End Sub
