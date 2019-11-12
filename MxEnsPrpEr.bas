Attribute VB_Name = "MxEnsPrpEr"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxEnsPrpEr."
Const LinOfLblX$ = "X: Debug.Print CSub & "".PrpEr["" &  Err.Description & ""]"""
Function HasLinExitAndLblX(MthLy$(), LinOfExit) As Boolean
Dim U&: U = UB(MthLy): If U < 2 Then Exit Function
If MthLy(U - 1) <> LinOfLblX Then Exit Function
If MthLy(U - 2) <> MthExitLin(MthLy(0)) Then Exit Function
End Function

Function InsLinExitAndLblX(MthLy$(), LinOfExit$) As String()
InsLinExitAndLblX = InsAy(MthLy, Sy(LinOfExit, LinOfLblX), UB(MthLy))
End Function

Function EnsLinExitAndLblX(MthLy$(), LinOfExit$) As String()
If HasLinExitAndLblX(MthLy, LinOfExit) Then EnsLinExitAndLblX = MthLy: Exit Function
Dim O$():
O = AeEle(MthLy, LinOfExit)
O = AeEle(MthLy, LinOfLblX)
EnsLinExitAndLblX = InsLinExitAndLblX(O, LinOfExit)
End Function

Function RmvOnErGoNonX(MthLy$()) As String()
Dim J&, I&
For J = 0 To UB(MthLy)
    If HasPfx(MthLy(J), "On Error Goto") Then
        For I = J + 1 To UB(MthLy)
            PushI RmvOnErGoNonX, MthLy(I)
        Next
        Exit Function
    End If
    PushI RmvOnErGoNonX, MthLy(J)
Next
End Function

Function InsOnErGoX(MthLy$()) As String()
InsOnErGoX = InsEle(MthLy, "On Error Goto X", NxtIxzSrc(MthLy))
End Function

Function EnsLinzOnEr(MthLy$()) As String()
Dim O$()
O = RmvOnErGoNonX(MthLy)
EnsLinzOnEr = InsOnErGoX(O)
End Function

Function MthExitLin$(MthLin)
MthExitLin = MthXXXLin(MthLin, "Exit")
End Function
Function MthXXXLin(MthLin, XXX$)
Dim X$: X = MthKd(MthLin): If X = "" Then Thw CSub, "Given Lin is not MthLin", "Lin", MthLin
MthXXXLin = XXX & " " & X
End Function

Function IxOfExit&(MthLy$())
Dim J&, LinOfExit$
LinOfExit = "Exit " & MthKd(MthLy(0))
For J = 0 To UB(MthLy)
    If MthLy(J) = LinOfExit Then IxOfExit = J: Exit Function
Next
IxOfExit = -1
End Function

Function IxOfLblX&(MthLy$())
Dim J&, L$
For J = 0 To UB(MthLy)
    If HasPfx(MthLy(J), "X: Debug.Print") Then IxOfLblX = J: Exit Function
Next
IxOfLblX = -1
End Function

Function IxOfOnEr&(PurePrpLy$())
Dim J&
For J = 0 To UB(PurePrpLy)
    If HasPfx(PurePrpLy(J), "On Error Goto X") Then IxOfOnEr = J: Exit Function
Next
IxOfOnEr = -1
End Function

Sub Z_EnsPrpOnerzS()
Dim Src$()
Const TstId& = 2
GoSub Z
'GoSub T1
Exit Sub
T1:
    Src = TstLy(TstId, CSub, 1, "Pm-Src", IsEdt:=False)
    Ept = TstLy(TstId, CSub, 1, "Ept", IsEdt:=False)
    GoTo Tst
    Return
Tst:
    Act = EnsPrpOnEoS(Src)
    C
    Return
Z:
    Src = CSrc
    Vc EnsPrpOnEoS(Src)
    Return
End Sub

Sub Z_IsMthLinSngL()
Dim L
GoSub T1
Exit Sub
T1:
    L = Sy("Function AA():End Function")
    Ept = True
    GoTo Tst
Tst:
    Act = IsMthLinSngL(L)
    C
    Return
End Sub
Sub RmvPrpOnEoM(M As CodeModule, Optional Upd As EmUpd)
'Dim L&(): L = LngAp( _
IxOfExit(PurePrpLy), _
IxOfOnEr(PurePrpLy), _
IxOfLblX(PurePrpLy))
'RmvPrpOnEoPurePrpLy = CvSy(AeIxy(PurePrpLy, L))
End Sub

Sub RmvPrpOnErM()
RmvPrpOnEoM CMd
End Sub

Sub EnsPrpOnEoM(M As CodeModule)
'RplMd A, JnCrLf(EnsPrpOnEoS(Src(A)))
End Sub

Function EnsPrpOnEoS(Src$()) As String()
'RplMd A, JnCrLf(EnsPrpOnEoS(Src(A)))
End Function

Sub EnsPrpOnErM()
EnsPrpOnEoM CMd
End Sub

