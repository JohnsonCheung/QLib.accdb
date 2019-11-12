Attribute VB_Name = "MxFun"
Option Compare Text
Option Explicit
Public Const CLib$ = "?."
Public Const CMod$ = CLib & "MxFun."
Public Const CSub$ = CMod & "?"

Sub Swap(OA, OB)
Dim X: X = OA
OA = OB
OB = X
End Sub

Function CUsr$()
CUsr = Environ$("USERNAME")
End Function

Sub Asg(Fm, OTo)
Select Case True
Case IsObject(Fm): Set OTo = Fm
Case Else: OTo = Fm
End Select
End Sub

Sub ShellHid(Fcmd$)
Shell Fcmd, vbHide
End Sub

Sub ShellMax(Fcmd$)
Shell Fcmd, vbMaximizedFocus
End Sub

Sub ThwIf_NBet(Fun$, V, FmV, ToV)
If FmV > V Then Thw Fun, "FmV > V", "V FmV ToV", V, FmV, ToV
If ToV < V Then Thw Fun, "ToV < V", "V FmV ToV", V, FmV, ToV
End Sub

Function InStrWiIthSubStr&(S, SubStr, Optional Ith% = 1)
Dim P&, J%
If Ith < 1 Then Thw CSub, "Ith cannot be <1", "Ith", Ith
For J = 1 To Ith
    P = InStr(P + 1, SubStr, SubStr)
    If P = 0 Then Exit Function
Next
InStrWiIthSubStr = P
End Function

Function InStrN&(S, SubStr, Optional N% = 1)
InStrN = InStrWiIthSubStr(S, SubStr, N)
End Function

Function CvNothing(A)
If IsEmpty(A) Then Set CvNothing = Nothing: Exit Function
Set CvNothing = A
End Function

Sub Z_InStrN()
Dim Act&, Exp&, S, SubStr, N%

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 1
Exp = 1
Act = InStrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 2
Exp = 6
Act = InStrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 3
Exp = 11
Act = InStrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 4
Exp = 0
Act = InStrN(S, SubStr, N)
Ass Exp = Act
End Sub

Function Max(A, b, ParamArray Ap())
Dim O: O = IIf(A > b, A, b)
Dim J%: For J = 0 To UBound(Ap)
   If Ap(J) > O Then O = Ap(J)
Next
Max = O
End Function


Function CvVbTy(A) As VbVarType
CvVbTy = A
End Function

Function MaxVbTy(A As VbVarType, b As VbVarType) As VbVarType
Dim O As VbVarType
If A = vbString Or b = vbString Then O = A: Exit Function
If A = vbEmpty Then O = b: Exit Function
If b = vbEmpty Then O = A: Exit Function
If A = b Then O = A: Exit Function
Dim AIsNum As Boolean, BIsNum As Boolean
AIsNum = IsVbTyNum(A)
BIsNum = IsVbTyNum(b)
Select Case True
Case A = vbBoolean And BIsNum: O = b
Case AIsNum And b = vbBoolean: O = A
Case A = vbDate Or b = vbDate: O = vbString
Case AIsNum And BIsNum:
    Select Case True
    Case A = vbByte: O = b
    Case b = vbByte: O = A
    Case A = vbInteger: O = b
    Case b = vbInteger: O = A
    Case A = vbLong: O = b
    Case b = vbLong: O = A
    Case A = vbSingle: O = b
    Case b = vbSingle: O = A
    Case A = vbDouble: O = b
    Case b = vbDouble: O = A
    Case A = vbCurrency Or b = vbCurrency: O = A
    Case Else: Stop
    End Select
Case Else: Stop
End Select
End Function

Function CprMth(IgnCas As Boolean) As VbCompareMethod
If IgnCas Then
    CprMth = vbBinaryCompare
Else
    CprMth = vbTextCompare
End If
End Function
Function CanCvLng(V) As Boolean
On Error GoTo X
Dim L&: L = CLng(V)
CanCvLng = True
X:
End Function

Function MinUB(Ay1, Ay2)
MinUB = Min(UB(Ay1), UB(Ay2))
End Function

Function Min(ParamArray A())
Dim O, J&, Av()
Av = A
Min = AyMin(Av)
End Function

Sub SndKeys(A$)
DoEvents
SendKeys A, True
End Sub

Function NDig&(N&)
Dim A$: A = N
NDig = Len(A)
End Function

Sub Vc(V, Optional Fnn$)
Brw V, Fnn, OupTy:=EiOtVc
End Sub

Sub b(A, Optional Fnn$, Optional OupTy As EmOupTy = EmOupTy.EiOtBrw)
Brw A, Fnn, OupTy
End Sub

Sub Brw(V, Optional Fnn$, Optional OupTy As EmOupTy = EmOupTy.EiOtBrw)
BrwAy Fmt(V), Fnn, OupTy
End Sub

Function Fmt(V) As String()
Select Case True
Case IsStr(V):     Fmt = Sy(V)
Case IsLinesAy(V): Fmt = FmtLinesAy(V)
Case IsArray(V):   Fmt = SyzAy(V)
Case IsAset(V):    Fmt = CvAset(V).Sy
Case IsDic(V):     Fmt = FmtDic(CvDic(V), InclValTy:=True)
Case IsEmpty(V):   Fmt = Sy("#Empty")
Case IsNothing(V): Fmt = Sy("#Nothing")
Case Else:         Fmt = Sy("#TypeName:" & TypeName(V))
End Select
End Function

Function MchStr(Patn$, S)
'Ret : either str or empty
Dim M As Match: Set M = MchzPatnF(Patn, S)
If IsNothing(M) Then Exit Function
MchStr = M.Value
End Function


Function SumLngAy@(A&())
Dim L, O@
For Each L In Itr(A)
    O = O + L
Next
SumLngAy = O
End Function

Sub Z_SumLngAy()
Dim S$: S = LineszFt(PjfP)
Debug.Assert SumLngAy(AscCntAy(S)) = Len(S)
End Sub

Function AscCntAy(S) As Long()
Dim O&(255), A As Byte
Dim J&
For J = 1 To Len(S)
    A = Asc(Mid(S, J, 1))
    O(A) = O(A) + 1
Next
AscCntAy = O
End Function
Function RemBlkSz&(N&, BlkSz%)
RemBlkSz = (N Mod BlkSz)
End Function

Function NBlk&(N&, BlkSz%)
NBlk = ((N - 1) \ BlkSz) + 1
End Function
