Attribute VB_Name = "MxFTIx"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxFTIx."
Type FCnt
    FmLno As Long
    Cnt As Long
End Type
Type Fei
FmIx As Long
EIx As Long ' To loop For J=FmIx To EIx-1, Always ToIx <-- EIx-1
End Type
Type Feis: N As Long: Ay() As Fei: End Type
Function FCnt(FmLno, Cnt) As FCnt
If FmLno <= 0 Then Exit Function
If Cnt <= 0 Then Exit Function
FCnt.FmLno = FmLno
FCnt.Cnt = Cnt
End Function
Function FCntzFei(A As Fei) As FCnt
With A
    FCntzFei = FCnt(.FmIx + 1, LinCntzFei(A))
End With
End Function
Function Fei(FmIx, EIx) As Fei
If 0 > FmIx Then Exit Function
If 0 > EIx Then Exit Function
If FmIx > EIx Then Exit Function
Fei.FmIx = FmIx
Fei.EIx = EIx
End Function
Sub PushFei(O As Feis, M As Fei)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Sub PushFeis(O As Feis, M As Feis)
Dim J&
For J = 0 To M.N - 1
    PushFei O, M.Ay(J)
Next
End Sub
Function AddFeis(A As Feis, b As Feis) As Feis
AddFeis = A
PushFeis A, b
End Function
Function SngFei(A As Fei) As Feis
PushFei SngFei, A
End Function

Function BetFei(Ix, A As Fei) As Boolean
BetFei = IsBet(Ix, A.FmIx, A.EIx - 1)
End Function

Function CntzFei&(A As Fei)
Dim O&
O = A.EIx - A.FmIx
If O < 0 Then Stop
CntzFei = O
End Function
Function FeizFC(FmIx, Cnt) As Fei
FeizFC = Fei(FmIx, FmIx + Cnt - 1)
End Function

Function IsEqFeis(A As Feis, b As Feis) As Boolean
If A.N <> b.N Then Exit Function
Dim J&
For J = 0 To A.N - 1
    If Not IsEqFei(A.Ay(J), b.Ay(J)) Then Exit Function
    J = J + 1
Next
IsEqFeis = True
End Function
Function IsFeizEmp(A As Fei) As Boolean

End Function
Function IsFeisInOrd(A As Feis) As Boolean
Dim J%
For J = 0 To A.N - 1
    With FCntzFei(A.Ay(J))
        If .FmLno = 0 Then Exit Function
        If .Cnt = 0 Then Exit Function
        If .FmLno + .Cnt > FCntzFei(A.Ay(J + 1)).FmLno Then Exit Function
    End With
Next
IsFeisInOrd = True
End Function

Function Positive(N)
If N > 0 Then Positive = N
End Function
Function LinCntzFei&(A As Fei)
LinCntzFei = Positive(A.EIx - A.FmIx)
End Function
Function LinCntzFeis&(A As Feis)
Dim J&, O&
For J = 0 To A.N - 1
    O = O + LinCntzFei(A.Ay(J))
Next
LinCntzFeis = O
End Function

Function IsEqFei(A As Fei, b As Fei) As Boolean
With A
    If .FmIx <> b.FmIx Then Exit Function
    If .EIx <> b.EIx Then Exit Function
End With
IsEqFei = True
End Function

