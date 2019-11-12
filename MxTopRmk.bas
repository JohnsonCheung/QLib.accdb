Attribute VB_Name = "MxTopRmk"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxTopRmk."
Sub Z_MthFeiszSrcMth()
Dim Src$(), Mthn
Dim Ept As Feis, Act As Feis

Src = SrczMdn("IdeMthFei")
PushFei Ept, Fei(2, 11)
GoSub Tst

Exit Sub
Tst:
    Act = MthFeiszSN(Src, Mthn)
    If Not IsEqFeis(Act, Ept) Then Stop
    Return
End Sub

Function RmvBlnkLin(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushNB RmvBlnkLin, I
Next
End Function

Function TopRmkMdes$(Src$(), MthIx)
TopRmkMdes = JnCrLf(TopRmkLy(Src, MthIx))
End Function
Function TopRmkLyzSIW(Src$(), MthIx) As String()
TopRmkLyzSIW = TopRmkLy(Src, MthIx)
End Function
Function TopRmkLy(Src$(), MthIx) As String()
Dim Fm&: Fm = TopRmkIx(Src, MthIx): If Fm = -1 Then Exit Function
TopRmkLy = RmvBlnkLin(AwFT(Src, Fm, MthIx - 1))
End Function

Function TopRmkIx&(Src$(), MthIx)
If MthIx <= 0 Then Exit Function
Dim J&, L$
TopRmkIx = MthIx
For J = MthIx - 1 To 0 Step -1
    L = LTrim(Src(J))
    Select Case True
    Case L = ""
    Case FstChr(L) = "'": TopRmkIx = J
    Case Else: Exit Function
    End Select
Next
End Function

Function TopRmkLno(Md As CodeModule, MthLno)
Dim J&, L$
TopRmkLno = MthLno
If MthLno = 0 Then Exit Function
For J = MthLno - 1 To 1 Step -1
    L = LTrim(Md.Lines(J, 1))
    Select Case True
    Case L = ""
    Case FstChr(L) = "'": TopRmkLno = J
    Case Else: Exit Function
    End Select
Next
End Function
