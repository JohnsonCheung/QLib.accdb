Attribute VB_Name = "MxPthR"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxPthR."
Private O$(), A_Spec$ ' Used in PrpcAyR/FFnAyR
Private XX$()

Function EmpPthAyR(Pth) As String()
Dim I
For Each I In Itr(SubPthAyR(Pth))
    If IsEmpPth(I) Then PushI EmpPthAyR, I
Next
End Function

Function EntAyR(Pth, Optional FilSpec$ = "*.*") As String()
Erase O
A_Spec = FilSpec
EntAyR1 Pth
EntAyR = O
Erase O
End Function

Sub EntAyR1(Pth)
If Si(O) Mod 1000 = 0 Then Debug.Print "EntAyR1: (Each 1000): " & Pth
PushI O, Pth
PushIAy O, FfnAy(Pth, A_Spec)
Dim I, P$()
'P = SubPthAyR(Pth, A_Spec)
For Each I In Itr(P)
    EntAyR1 I
Next
End Sub
Sub Z_FfnAyR()
Dim Pth, Spec$, AtR As FileAttribute
GoSub T0
GoSub T1
Exit Sub
T0:
    Pth = "C:\Users\User\Documents\Projects\Vba"
    GoTo Tst
T1:
    Pth = "C:\Users\User\Documents\WindowsPowershell\"
    GoTo Tst
Tst:
    Act = FfnAyR(Pth, Spec)
    Brw Act
    Stop
    Return
End Sub
Function FfnAyR(Pth, Optional Spec$ = "*.*") As String()
Erase O
A_Spec = Spec
FfnAyR1 Pth
FfnAyR = O
End Function

Sub FfnAyR1(Pth)
PushIAy O, FfnAy(Pth, A_Spec)
If Si(O) Mod 1000 = 0 Then InfLin CSub, "...Reading", "#Ffn-read", Si(O)
Dim P$(): P = SubPthAy(Pth)
If Si(P) = 0 Then Exit Sub
Dim I: For Each I In P
    FfnAyR1 I
Next
End Sub

Sub Z_EntAyR()
Dim A$(): A = EntAyR("C:\users\user\documents\")
Debug.Print Si(A)
Stop
DmpAy A
End Sub

Sub Z_EmpPthAyR()
Brw EmpPthAyR(TmpRoot)
End Sub

Sub Z_EntAy()
BrwPth EntAy(TmpRoot)
End Sub

Sub Z_DltEmpPthR()
Z:
    DltEmpPthR TmpRoot
    Return
Z1:
    Debug.Print "Before-----"
    D EmpPthAyR(TmpRoot)
    DltEmpPthR TmpRoot
    Debug.Print "After-----"
    D EmpPthAyR(TmpRoot)
    Return
End Sub

Function SubPthAyR(Pth) As String()
Erase XX
X Pth
SubPthAyR = XX
Erase XX
End Function

Private Sub X(Pth)
Dim O$(), P$, I
O = SubPthAy(Pth)
PushIAy XX, O
For Each I In Itr(O)
    P = I
    X P
Next
End Sub
