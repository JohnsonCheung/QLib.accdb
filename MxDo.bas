Attribute VB_Name = "MxDo"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxDo."
Sub ForEach(Ay, FunNm$)
If Si(Ay) = 0 Then Exit Sub
Dim I: For Each I In Ay
    Run FunNm, I
Next
End Sub

Sub ForEachABX(Ay, ABX$, A, b)
Dim X: For Each X In Itr(Ay)
    Run ABX, A, b, X
Next
End Sub

Sub ForEachAXB(Ay, AXB$, A, b)
Dim X: For Each X In Itr(Ay)
    Run AXB, A, X, b
Next
End Sub

Sub ForEachPPXP(A, PPXP$, P1, P2, P3)
Dim X
For Each X In Itr(A)
    Run PPXP, P1, P2, X, P3
Next
End Sub

Sub ForEachPX(A, PX$, P)
Dim X
For Each X In Itr(A)
    Run PX, P, X
Next
End Sub

Sub ForEachXP(A, Xp$, P)
Dim X
For Each X In Itr(A)
    Run Xp, X, P
Next
End Sub
