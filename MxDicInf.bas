Attribute VB_Name = "MxDicInf"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxDicInf."

Function VyzDicKK(Dic As Dictionary, KK$) As Variant()
VyzDicKK = VyzDicKy(Dic, SyzSS(KK))
End Function

Function LineszLinesDic(LinesDic As Dictionary, Optional LinesSep$ = vbCrLf) ' Return the joined Lines from LinesDic
Dim O$(), I, Lines$
For Each I In LinesDic.Items
    PushI O, I
Next
LineszLinesDic = Jn(O, LinesSep)
End Function

Function AddPfxToKey(Pfx$, A As Dictionary) As Dictionary
Dim K
Set AddPfxToKey = New Dictionary
For Each K In A.Keys
    AddPfxToKey.Add Pfx & K, A(K)
Next
End Function



Function DicHasBlnkKey(A As Dictionary) As Boolean
If A.Count = 0 Then Exit Function
Dim K
For Each K In A.Keys
   If Trim(K) = "" Then DicHasBlnkKey = True: Exit Function
Next
End Function

Function DicHasK(A As Dictionary, K$) As Boolean
DicHasK = A.Exists(K)
End Function

Function DicHasKeyLvs(A As Dictionary, KeyLvs) As Boolean
DicHasKeyLvs = DicHasKy(A, SyzSS(KeyLvs))
End Function

Sub DicHasKeyssAss(A As Dictionary, KeySS$)
DicHasKyAss A, SyzSS(KeySS)
End Sub

Function DicHasKeySsl(A As Dictionary, KeySsl) As Boolean
DicHasKeySsl = A.Exists(SyzSS(KeySsl))
End Function

Function DicHasKy(A As Dictionary, Ky) As Boolean
Ass IsArray(Ky)
If Si(Ky) = 0 Then Stop
Dim K
For Each K In Ky
   If Not A.Exists(K) Then
       Debug.Print FmtQQ("Dix.HasKy: Key(?) is Missing", K)
       Exit Function
   End If
Next
DicHasKy = True
End Function

Sub DicHasKyAss(A As Dictionary, Ky)
Dim K
For Each K In Ky
   If Not A.Exists(K) Then Debug.Print K: Stop
Next
End Sub


Sub Z_IsDicKeyStr()
Dim A As Dictionary
GoSub T1
Exit Sub
T1:
    Set A = New Dictionary
    Dim J&
    For J = 1 To 10000
        A.Add J, J
    Next
    Ept = True
    GoSub Tst
    '
    A.Add 10001, "X"
    Ept = False
    GoTo Tst
Tst:
    Act = IsDicKeyStr(A)
    C
    Return
End Sub


Function TyNmAy(Ay) As String()
Dim V
For Each V In Itr(Ay)
    PushI TyNmAy, TypeName(V)
Next
End Function

Function VyzDicKy(D As Dictionary, Ky) As Variant()
Dim K
For Each K In Itr(Ky)
    If Not D.Exists(K) Then Thw CSub, "Some K in given Ky not found in given Dic keys", "[K with error] [given Ky] [given dic keys]", K, AvzItr(D.Keys), Ky
    Push VyzDicKy, D(K)
Next
End Function
Function DicwKy(D As Dictionary, Ky) As Dictionary
Set DicwKy = New Dictionary
Dim Vy(): Vy = VyzDicKy(D, Ky)
Dim K, J&
For Each K In Itr(Ky)
    DicwKy.Add K, Vy(J)
    J = J + 1
Next
End Function

Function Vy(A As Dictionary) As Variant()
Vy = IntozItr(EmpAv, A.Items)
End Function
Function TyNmAyzDic(A As Dictionary) As String()
TyNmAyzDic = TyNmAy(Vy(A))
End Function



Function DicTy$(A As Dictionary)
Dim O$
Select Case True
Case IsDicEmp(A):   O = "EmpDic"
Case IsStrDic(A):   O = "StrDic"
Case IsDicLines(A): O = "LineszDic"
Case IsDicSy(A):    O = "DiT1qLy"
Case Else:           O = "Dic"
End Select
End Function

Sub AddDicLin(ODic As Dictionary, DicLin$)
With BrkSpc(DicLin)
    ODic.Add .S1, .S2
End With
End Sub
Function AddDic(A As Dictionary, b As Dictionary) As Dictionary
Set AddDic = New Dictionary
PushDic AddDic, A
PushDic AddDic, b
End Function

Function DicAyzAp(ParamArray DicAp()) As Dictionary()
Dim Av(): Av = DicAp: If Si(Av) = 0 Then Exit Function
Dim I
For Each I In Av
    If Not IsDic(I) Then Thw CSub, "Some itm is not Dic", "TypeName-Ay", VbTyNyzAy(Av)
    PushObj DicAyzAp, CvDic(I)
Next
End Function
Function DefDic(Ly$(), KK) As Dictionary
Dim L, S As Aset, T1$, Rst$, O As New Dictionary
Set S = TermAset(KK)
If S.Has("*Er") Then Thw CSub, "KK cannot have Term-*Er", "KK Ly", KK, Ly
For Each L In Ly
    AsgTRst L, T1, Rst
    If S.Has(T1) Then
        PushItmzDiT1qLy O, T1, Rst
    Else
'        PushItmzDiT1qLy , O, L
    End If
    Set DefDic = O
Next
End Function

Function DiceKeySet(A As Dictionary, ExlKeySet As Aset) As Dictionary
Dim K
Set DiceKeySet = New Dictionary
For Each K In A.Keys
    If Not ExlKeySet.Has(K) Then
        DiceKeySet.Add K, A(K)
    End If
Next
End Function



Function DicwKK(A As Dictionary, KK) As Dictionary
Set DicwKK = New Dictionary
Dim K
For Each K In TermAy(KK)
    If A.Exists(K) Then
        DicwKK.Add K, A(K)
    End If
Next
End Function

Function KeyToLikAyDic_TkssLy(TkssLy$()) As Dictionary
Dim O As Dictionary
    Set O = Dic(TkssLy)
Dim K
For Each K In O.Keys
    O(K) = SyzSS(O(K))
Next
Set KeyToLikAyDic_TkssLy = O
End Function
