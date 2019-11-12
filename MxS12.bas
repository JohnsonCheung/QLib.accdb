Attribute VB_Name = "MxS12"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxS12."
Type S12: S1 As String: S2 As String: End Type
Type S12Opt: Som As Boolean: S12 As S12: End Type
Type S12s: N As Long: Ay() As S12: End Type
Type S3: A As String: B As String: C As String: End Type
Function MapS1(A As S12s, Dic As Dictionary) As S12s
Dim J&: For J = 0 To A.N - 1
    Dim M As S12: M = A.Ay(J)
    If Not Dic.Exists(M.S1) Then
        Thw CSub, "Som S1 in [S12s] not found in [Dic]", "S1-not-found S12s Dic", M.S1, FmtS12s(A), FmtDic(Dic)
    End If
    M.S1 = Dic(M.S1)
    PushS12 MapS1, M
Next
End Function

Sub WrtS12s(A As S12s, Ft$, Optional OvrWrt As Boolean)
WrtStr S12sStr(A), Ft, OvrWrt
End Sub

Function S12Str$(A As S12)
S12Str = A.S1 & Chr(5) & A.S2
End Function

Function S12zStr(S12Str$) As S12
S12zStr = Brk(S12Str, Chr(5), NoTrim:=True)
End Function

Function S12szStr(S12sStr$) As S12s
Dim S: For Each S In Itr(Split(S12sStr, Chr(&H14)))
    PushS12 S12szStr, Brk(S, Chr(5), NoTrim:=True)
Next
End Function

Function S12szRes(ResFn$, Optional ResPseg$) As S12s
S12szRes = S12szStr(Resl(ResFn, ResPseg))
End Function

Function S12sStr$(A As S12s)
Dim J&, O$(): For J = 0 To A.N - 1
    PushI O, S12Str(A.Ay(J))
Next
S12sStr = Jn(O, Chr(&H14))
End Function

Function SwapS12s(A As S12s) As S12s
Dim J&, Ay() As S12: Ay = A.Ay
Dim O As S12s: O = A
For J = 1 To A.N - 1
    O.Ay(J) = SwapS12(Ay(J))
Next
SwapS12s = O
End Function

Function SwapS12(A As S12) As S12
With SwapS12
    .S1 = A.S2
    .S2 = A.S1
End With
End Function

Sub PushS12(O As S12s, M As S12)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function AddS12(A As S12, B As S12) As S12s
Dim O As S12s: O = S12szU(1)
O.Ay(0) = A
O.Ay(1) = B
AddS12 = O
End Function
Function Y_S12s() As S12s
Dim O As S12s
PushS12 O, S12("sldjflsdkjf", "lksdjf")
PushS12 O, S12("sldjflsdkjf", "lksdjf")
PushS12 O, S12("sldjf", "lksdjf")
PushS12 O, S12("sldjdkjf", "lksdjf")
Y_S12s = O
End Function

Function S12szU(U&) As S12s
S12szU.N = U + 1
ReDim S12szU.Ay(U)
End Function

Function S12szAyab(A, B, Optional NoTrim As Boolean) As S12s
ThwIf_DifSi A, B, CSub
Dim U&, O As S12s
U = UB(A)
O = S12szU(U)
Dim J&
For J = 0 To U
    O.Ay(J) = S12(A(J), B(J), NoTrim)
Next
S12szAyab = O
End Function

Function FstS2(S1, A As S12s) As StrOpt
'Ret : Lookup S1 in A return S2 @@
Dim Ay() As S12: Ay = A.Ay
Dim J&: For J = 0 To A.N - 1
    With Ay(J)
        If .S1 = S1 Then FstS2 = SomStr(.S2): Exit Function
    End With
Next
End Function
Function AddS2Sfx(A As S12s, S2Sfx$) As S12s
Dim O As S12s: O = A
Dim J&: For J = 0 To O.N - 1
    O.Ay(J).S2 = O.Ay(J).S2 & S2Sfx
Next
AddS2Sfx = O
End Function

Function S12szDrs(D As Drs, Optional CC$) As S12s
'Fm D  : ..@CC.. ! A drs with col-@CC.  At least has 2 col
'Fm CC :         ! if isBlnk, use fst 2 col
'Ret   :         ! fst col will be S1 and snd col will be S2 join with vbCrLf
Dim S1$(), S2() ' S2 is ay of sy
Dim I1%, I2%
    If CC = "" Then I1 = 0: I2 = 1 Else AsgIx D, CC, I1, I2
Dim Dr: For Each Dr In Itr(D.Dy)
    Dim A$, B$: A = Dr(I1): B = Dr(I2)
    Dim R&: R = IxzAy(S1, A, ThwEr:=EiNoThw)
    If R = -1 Then
        PushI S1, A
        PushI S2, Sy(B)
    Else
        PushI S2(R), B
    End If
Next
Dim J&: For J = 0 To UB(S1)
    PushS12 S12szDrs, S12(S1(J), JnCrLf(S2(J)))
Next
End Function

Function IsEqS12(A As S12, B As S12) As Boolean
With A
    If .S1 <> B.S1 Then Exit Function
    If .S2 <> B.S2 Then Exit Function
End With
IsEqS12 = True
End Function

Function HasS12(A As S12s, B As S12) As Boolean
Dim Ay() As S12: Ay = A.Ay
Dim J&: For J = 0 To A.N - 1
    If IsEqS12(Ay(J), B) Then HasS12 = True: Exit Function
Next
End Function
Function S12szDif(A As S12s, B As S12s) As S12s
'Ret : Subset of @A.  Those itm in @A also in @B will be exl.
Dim Ay() As S12: Ay = A.Ay
Dim J&: For J = 0 To A.N - 1
    If Not HasS12(B, Ay(J)) Then
        PushS12 S12szDif, Ay(J)
    End If
Next
End Function
Function CvStr$(V)
If IsNull(V) Then Exit Function
CvStr = V
End Function

Function S12(Optional S1, Optional S2, Optional NoTrim As Boolean) As S12
If NoTrim Then
    S12.S1 = S1
    S12.S2 = S2
Else
    S12.S1 = Trim(CvStr(S1))  ' Need CvStr, because S1 maybe null (due to reading Fx Sheet: DrszFxw, calling S12 in RduDrs)
    S12.S2 = Trim(CvStr(S2))
End If
End Function

Sub AsgS12(A As S12, O1, O2)
O1 = A.S1
O2 = A.S2
End Sub

Sub BrwS12s(A As S12s)
BrwAy FmtS12s(A)
End Sub

Function S12szDic(A As Dictionary) As S12s
Dim K
For Each K In A.Keys
    PushS12 S12szDic, S12(K, A(K))
Next
End Function

Function DiczS12s(A As S12s, Optional Sep$ = " ") As Dictionary
Dim J&, O As New Dictionary
For J = 0 To A.N - 1
    With A.Ay(J)
        If O.Exists(.S1) Then
            O(.S1) = O(.S1) & " " & O(.S2)
        Else
            O.Add .S1, .S2
        End If
    End With
Next
Set DiczS12s = O
End Function

Function S1Ay(A As S12s) As String()
Dim J&: For J = 0 To A.N - 1
   PushI S1Ay, A.Ay(J).S1
Next
End Function

Function S2Ay(A As S12s) As String()
Dim J&: For J = 0 To A.N - 1
   Push S2Ay, A.Ay(J).S2
Next
End Function

Function SqzS12s(A As S12s, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Variant()
If A.N = 0 Then Exit Function
Dim O(), I, R&, J&
ReDim O(1 To A.N + 1, 1 To 2)
R = 2
O(1, 1) = Nm1
O(1, 2) = Nm2
For J = 0 To A.N - 1
    With A.Ay(J)
        O(R, 1) = .S1
        O(R, 2) = .S2
        R = R + 1
    End With
Next
SqzS12s = O
End Function
Function S12szColonVbl(ColonVbl$) As S12s
Dim I
For Each I In SplitVBar(ColonVbl)
    PushS12 S12szColonVbl, BrkBoth(I, ":")
Next
End Function

Function S12szSySep(Sy$(), Sep$, Optional NoTrim As Boolean) As S12s
Dim O As S12s, J%
Dim U&: U = UB(Sy)
O = S12szU(U)
For J = 0 To U
    O.Ay(J) = Brk1(Sy(J), Sep, NoTrim)
Next
S12szSySep = O
End Function
Sub Z_S12szDic()
Dim A As New Dictionary
A.Add "A", "BB"
A.Add "B", "CCC"
Dim Act As S12s
Act = S12szDic(A)
Stop
End Sub



Function AddS1Pfx(A As S12s, S1Pfx$) As S12s
Dim J&: For J = 0 To A.N - 1
    Dim M As S12: M = A.Ay(J)
    M.S1 = S1Pfx & M.S1
    PushS12 AddS1Pfx, M
Next
End Function

Sub PushS12s(O As S12s, A As S12s)
Dim J&
For J = 0 To A.N - 1
    PushS12 O, A.Ay(J)
Next
End Sub
