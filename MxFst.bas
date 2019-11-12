Attribute VB_Name = "MxFst"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxFst."

Function RmvFstEle(Ay)
If Si(Ay) = 0 Then Thw CSub, "No FstEle"
Dim O: O = Ay: Erase O
Dim J&: For J = 1 To UB(Ay)
    PushI O, Ay(J)
Next
RmvFstEle = O
End Function

Function FstEleInAset(Ay, InAset As Aset)
Dim I
For Each I In Ay
    If InAset.Has(I) Then FstEleInAset = I: Exit Function
Next
End Function

Function FstEleLik$(A, Lik$)
If Si(A) = 0 Then Exit Function
Dim X
For Each X In A
    If X Like Lik Then FstEleLik = X: Exit Function
Next
End Function

Function FstElePredPX(A, PX$, P)
If Si(A) = 0 Then Exit Function
Dim X
For Each X In Itr(A)
    If Run(PX, P, X) Then Asg X, FstElePredPX: Exit Function
Next
End Function

Function FstElePredXABTrue(Ay, XAB$, A, b)
Dim X
For Each X In Itr(Ay)
    If Run(XAB, X, A, b) Then Asg X, FstElePredXABTrue: Exit Function
Next
End Function

Function FstElePredXP(A, Xp$, P)
If Si(A) = 0 Then Exit Function
Dim X
For Each X In Itr(A)
    If Run(Xp, X, P) Then Asg X, FstElePredXP: Exit Function
Next
End Function

Function FstElewRmvT1$(Sy$(), T1)
FstElewRmvT1 = RmvT1(FstElewT1(Sy, T1))
End Function

Function FstElewT1$(Ay, T1)
Dim I
For Each I In Itr(Ay)
    If T1zS(I) = T1 Then FstElewT1 = I: Exit Function
Next
End Function

Function FstElezPfxSy$(PfxSy$(), Lin)
Dim I, P$
For Each I In PfxSy
    P = I
    If HasPfx(Lin, P) Then FstElezPfxSy = P: Exit Function
Next
End Function

Function FstElezRmvT1$(Sy$(), T1)
FstElezRmvT1 = RmvT1(FstElezT1(Sy, T1))
End Function

Function FstElezT1$(Sy$(), T1)
Dim I, S$
For Each I In Itr(Sy)
    S = I
    If HasT1(S, T1) Then FstElezT1 = S: Exit Function
Next
End Function

Function FstElezT2$(Sy$(), T2)
Dim I, S$
For Each I In Itr(Sy)
    S = I
    If HasT2(S, T2) Then FstElezT2 = S: Exit Function
Next
End Function

Function FstElezTT$(Sy$(), T1, T2)
Dim I, S$
For Each I In Itr(Sy)
    S = I
    If HasTT(S, T1, T2) Then FstElezTT = S: Exit Function
Next
End Function

Function FstNEle(Ay, N)
FstNEle = AwFstUEle(Ay, N - 1)
End Function

Function ShfFstEle(OAy)
ShfFstEle = FstEle(OAy)
OAy = AeFstNEle(OAy)
End Function
