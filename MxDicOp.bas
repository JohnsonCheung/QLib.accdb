Attribute VB_Name = "MxDicOp"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxDicOp."

Function CvDic(A) As Dictionary
Set CvDic = A
End Function

Function AsetzDicKey(A As Dictionary) As Aset
Set AsetzDicKey = AsetzItr(A.Keys)
End Function

Function CvDicAy(A) As Dictionary()
CvDicAy = A
End Function

Function AddDicAy(A As Dictionary, Dy() As Dictionary) As Dictionary
Set AddDicAy = CloneDic(A)
Dim J%
For J = 0 To UB(Dy)
   PushDic AddDicAy, Dy(J)
Next
End Function

Function IupDic(A As Dictionary, By As Dictionary) As Dictionary 'Return New dictionary from A-Dic by Ins-or-upd By-Dic.  Ins: if By-Dic has key and A-Dic. _
Upd: K fnd in both, A-Dic-Val will be replaced by By-Dic-Val
Dim O As New Dictionary, K
For Each K In A.Keys
    If By.Exists(K) Then
        O.Add K, By(K)
    Else
        O(K) = By(K)
    End If
Next
Set IupDic = O
End Function

Function AddDicKeyPfx(A As Dictionary, Pfx) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add Pfx & K, A(K)
Next
Set AddDicKeyPfx = O
End Function

Sub DicAddOrUpd(A As Dictionary, K$, V, Sep$)
If A.Exists(K) Then
    A(K) = A(K) & Sep & V
Else
    A.Add K, V
End If
End Sub

Function DicAyKy(A() As Dictionary) As Variant()
Dim I
For Each I In Itr(A)
   PushNDupAy DicAyKy, CvDic(I).Keys
Next
End Function

Function CloneDic(A As Dictionary) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add K, A(K)
Next
Set CloneDic = O
End Function

Function DrDicKy(A As Dictionary, Ky$()) As Variant()
Dim O(), I, J&
ReDim O(UB(Ky))
For Each I In Ky
    If A.Exists(I) Then
        O(J) = A(I)
    End If
    J = J + 1
Next
DrDicKy = O
End Function

Function DicFny(InclValTy As Boolean) As String()
DicFny = SplitSpc("Key Val"): If InclValTy Then PushI DicFny, "ValTy"
End Function
Function DyoDotAy(DotAy$()) As Variant()
Dim I, Lin
For Each I In Itr(DotAy)
    Lin = I
    PushI DyoDotAy, SplitDot(Lin)
Next
End Function
Function DyzDi(A As Dictionary, Optional InclValTy As Boolean) As Variant()
Dim I, Dr
If A.Count = 0 Then Exit Function
Dim K(): K = A.Keys
If Si(K) = 0 Then Exit Function
For Each I In K
    If InclValTy Then
        Dr = Array(I, A(I), TypeName(A(I)))
    Else
        Dr = Array(I, A(I))
    End If
    Push DyzDi, Dr
Next
End Function

Function DicAyIntersect(A As Dictionary, B As Dictionary) As Dictionary
Dim O As New Dictionary
If A.Count = 0 Then GoTo X
If B.Count = 0 Then GoTo X
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            O.Add K, A(K)
        End If
    End If
Next
X: Set DicAyIntersect = O
End Function
Sub ThwIf_DifDic(A As Dictionary, B As Dictionary, Fun$, Optional N1$ = "A", Optional N2$ = "B")
If Not IsEqDic(A, B) Then Thw Fun, "2 given dic are diff", FmtQQ("[?] [?]", N1, N2), FmtDic(A), FmtDic(B)
End Sub

Function KeySet(A As Dictionary) As Aset
Set KeySet = AsetzItr(A.Keys)
End Function

Function KeySyzDic(A As Dictionary) As String()
KeySyzDic = SyzAy(A.Keys)
End Function

Function VzDicIfKyJn$(A As Dictionary, Ky, Optional Sep$ = vb2CrLf)
Dim O$(), K
For Each K In Itr(Ky)
    If A.Exists(K) Then
        PushI O, A(K)
    End If
Next
VzDicIfKyJn = Join(O, Sep)
End Function

Function SyzDicKy(Dic As Dictionary, Ky$()) As String()
Const CSub$ = CMod & "SyzDicKy"
Dim K
For Each K In Itr(Ky)
    If Dic.Exists(K) Then Thw CSub, "K of Ky not in Dic", "K Ky Dic", K, Ky, Dic
    PushI SyzDicKy, Dic(K)
Next
End Function

Function LineszDic$(A As Dictionary)
LineszDic = JnCrLf(FmtDic2(A))
End Function

Function FmtDic2(A As Dictionary) As String()
Dim K: For Each K In A.Keys
    Push FmtDic2, LyzKLines(K, A(K))
Next
End Function

Function LyzKLines(K, Lines$) As String()
Dim Ly$(): Ly = SplitCrLf(Lines)
Dim J&: For J = 0 To UB(Ly)
    Dim Lin
        Lin = Ly(J)
        If FstChr(Lin) = " " Then Lin = "~" & RmvFstChr(Lin)
    Push LyzKLines, K & " " & Lin
Next
End Function

Function MaxSizAyDic%(AyDic As Dictionary) ' DiMthnqLines is DicOf_Mthn_zz_MthlAy
'MaxCntgMth is max-of-#-of-method per Mthn
Dim O%, K
For Each K In AyDic.Items
    O = Max(O, Si(AyDic(K)))
Next
MaxSizAyDic = O
End Function

Function MgeDic(A As Dictionary, PfxSsl$, ParamArray DicAp()) As Dictionary
Dim Av(): Av = DicAp
Dim Ny$()
   Ny = SyzSS(PfxSsl)
   Ny = AmAddSfx(Ny, "@")
If Si(Av) <> Si(Ny) Then Stop
Dim Dy() As Dictionary
Dim D As Dictionary
   Dim J%
   For J = 0 To UB(Ny)
       Set D = Av(J)
       Push Dy, AddDicKeyPfx(A, Ny(J))
   Next
Set MgeDic = AddDicAy(A, Dy)
End Function

Sub BrwKSet(KSet As Dictionary)
BrwDrs DrszKSet(KSet)
End Sub

Function DrszKSet(KSet As Dictionary) As Drs
Dim K, Dy(), S As Aset, V
For Each K In KSet.Keys
    Set S = KSet(K)
    If S.Cnt = 0 Then
        PushI Dy, Array(K, "#EmpSet#")
    Else
        For Each V In S.Itms
            PushI Dy, Array(K, V)
        Next
    End If
Next
DrszKSet = DrszFF("K V", Dy)
End Function
Function HasKSet(KSet As Dictionary, K, S As Aset) As Boolean
'Fm KSet : KSet if a dictionary with value is Aset.
'Ret     : True if KSet has such Key-K and Val-Set-S  @@
If KSet.Exists(K) Then
    Dim ISet As Aset: Set ISet = KSet(K)
    HasKSet = ISet.IsEq(S)
End If
End Function
Function KSetzDif(KSet1 As Dictionary, KSet2 As Dictionary)
'Ret : KSet from KSet1 where not found in KSet2 (Not found means K is not found or K is found but V is dif @@
Set KSetzDif = New Dictionary
Dim K: For Each K In KSet1.Keys
    Dim V As Aset: Set V = KSet1(K)
    Dim Has As Boolean: Has = HasKSet(KSet2, K, V)
    If Not Has Then
        KSetzDif.Add K, V
    End If
Next

End Function
Function MinusDic(A As Dictionary, B As Dictionary) As Dictionary
'Ret those Ele in A and not in B
If B.Count = 0 Then Set MinusDic = CloneDic(A): Exit Function
Dim O As New Dictionary, K
For Each K In A.Keys
   If Not B.Exists(K) Then O.Add K, A(K)
Next
Set MinusDic = O
End Function

Function DicSelIntozAy(A As Dictionary, Ky$()) As Variant()
Dim O()
Dim U&: U = UB(Ky)
ReDim O(U)
Dim J&
For J = 0 To U
   If Not A.Exists(Ky(J)) Then Stop
   O(J) = A(Ky(J))
Next
DicSelIntozAy = O
End Function

Function DicSelIntoSy(A As Dictionary, Ky$()) As String()
DicSelIntoSy = SyzAy(DicSelIntozAy(A, Ky))
End Function

Function SyzDicKey(A As Dictionary) As String()
SyzDicKey = SyzItr(A.Keys)
End Function

Function DiczSwapKV(A As Dictionary) As Dictionary
Dim K
Set DiczSwapKV = New Dictionary
For Each K In A.Keys
    DiczSwapKV.Add A(K), K
Next
End Function

Function DicValOpt(A As Dictionary, K)
If IsNothing(A) Then Exit Function
If A.Exists(K) Then Asg A(K), DicValOpt
End Function

Function KeyzLikAyDic_Itm$(Dic As Dictionary, Itm$)
Dim K, LikAy$()
For Each K In Dic.Keys
    LikAy = Dic(K)
    If HitLikAy(Itm, LikAy) Then
        KeyzLikAyDic_Itm = K
        Exit Function
    End If
Next
End Function

Function KeyzKssDic_Itm$(A As Dictionary, Itm$)
Dim Kss$, K
For Each K In A
    Kss = A(K)
    If HitKss(Itm, Kss) Then KeyzKssDic_Itm = K: Exit Function
Next
End Function

Sub Z_MaxSizAyDic()
Dim D As Dictionary, M%
'Set D = PjDiMthnqLines(CPj)
M = MaxSizAyDic(D)
Stop
End Sub


Function WbzNmzDiLines(NmzDiLines As Dictionary) As Workbook 'Assume each dic keys is name and each value is lines. _
create a new Wb with worksheet as the dic key and the lines are break to each cell of the sheet
Dim A As Dictionary: Set A = NmzDiLines
Ass IsItrNm(A.Keys)
Ass IsItrStr(A.Items)
Dim K, ThereIsSheet1 As Boolean
Dim O As Workbook: Set O = NewWb
Dim Ws As Worksheet
For Each K In A.Keys
    If K = "Sheet1" Then
        ThereIsSheet1 = True
    Else
        Set Ws = O.Sheets.Add
        Ws.Name = K
    End If
    Ws.Range("A1").Value = SqvzLines(A(K))
Next
X: Set WbzNmzDiLines = O
End Function

Function DicACzOuter(DicAB As Dictionary, DicBC As Dictionary) As Dictionary
Dim A, B, C
Set DicACzOuter = New Dictionary
For Each A In DicAB.Keys
    B = DicAB(A)
    If DicBC.Exists(B) Then
        DicACzOuter.Add A, C
    Else
        DicACzOuter.Add A, Empty
    End If
Next
End Function
Function DicAC(DicAB As Dictionary, DicBC As Dictionary) As Dictionary
Dim A, B, C
Set DicAC = New Dictionary
For Each A In DicAB.Keys
    B = DicAB(A)
    If DicBC.Exists(B) Then
        DicAC.Add A, DicBC(B)
    End If
Next
End Function

Function DicAzDifVal(A As Dictionary, B As Dictionary) As Dictionary
Set DicAzDifVal = New Dictionary
Dim K, V
For Each K In A.Keys
    If B.Exists(K) Then
        V = A(K)
        If V <> B(K) Then DicAzDifVal.Add K, V
    End If
Next
End Function
Sub SetKv(O As Dictionary, K, V)
If O.Exists(K) Then
    Asg V, O(K)
Else
    O.Add K, V
End If
End Sub

Sub PushDic(O As Dictionary, A As Dictionary)
If IsNothing(O) Then
    Set O = CloneDic(A)
    Exit Sub
End If
Dim K
For Each K In A.Keys
    If O.Exists(A) Then Thw CSub, "O already has K.  Cannot push Dic-A to Dic-O", "K Dic-O Dic-A", K, O, A
    O.Add K, A(K)
Next
End Sub

Function ChnDic(DiAqB As Dictionary, DiBqC As Dictionary) As Dictionary
':Ret :DiAqC  ! Thw Er if A->B and B not in fnd in DiBqC @@
Dim A As Dictionary: Set A = DiAqB
Dim B As Dictionary: Set B = DiBqC
Set ChnDic = New Dictionary
Dim KA: For Each KA In A.Keys
    If Not B.Exists(A(KA)) Then
        Thw CSub, "KeyA has ValB which not found in DiBqC", "KeyA ValB DiAqB DiBqC", KA, A(KA), FmtDic(DiAqB), FmtDic(DiBqC)
    End If
    ChnDic.Add KA, B(A(KA))
Next
End Function

Sub PushItmzDiT1qLy(A As Dictionary, K, Itm)
Dim M$()
If A.Exists(K) Then
    M = A(K)
    PushI M, Itm
    A(K) = M
Else
    A.Add K, Sy(Itm)
End If
End Sub

Sub ThwNotDiT1qLy(A As Dictionary, Fun$)
If Not IsDicSy(A) Then Thw Fun, "Given dictionary is not DiT1qLy, all key is string and val is Sy", "Give-Dictionary", FmtDic(A)
End Sub

Function AddSfxzDic(D As Dictionary, Sfx$) As Dictionary
Dim O As New Dictionary
Dim K: For Each K In D.Keys
    Dim V$: V = D(K) & Sfx
    O.Add K, V
Next
Set AddSfxzDic = O
End Function

Sub PushNBlnkzDi(ODic As Dictionary, K, S$)
If S = "" Then Exit Sub
ODic.Add K, S
End Sub

Sub SetDicValAsItr(O As Dictionary)
Dim K: For Each K In O.Keys
    O(K) = Itr(O(K))
Next
End Sub

