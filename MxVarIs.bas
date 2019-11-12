Attribute VB_Name = "MxVarIs"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVarIs."
Const C_UCasA% = 65
Const C_UCasZ% = 90
Const C_LCasA% = 97
Const C_LCasZ% = 122

Function IsAv(V) As Boolean
IsAv = VarType(V) = vbArray + vbVariant
End Function

Function IsAyDic(V As Dictionary) As Boolean
If Not IsSy(V.Keys) Then Exit Function
If Not IsAyOfAy(V.Items) Then Exit Function
IsAyDic = True
End Function

Function IsAyOfAy(V) As Boolean
If Not IsAv(V) Then Exit Function
Dim X
For Each X In Itr(V)
    If Not IsArray(X) Then Exit Function
Next
IsAyOfAy = True
End Function

Function IsBool(V) As Boolean
IsBool = VarType(V) = vbBoolean
End Function
Function IsBoolAy(V) As Boolean
IsBoolAy = VarType(V) = vbArray + vbBoolean
End Function

Function IsByt(V) As Boolean
IsByt = VarType(V) = vbByte
End Function

Function IsBytAy(V) As Boolean
IsBytAy = VarType(V) = vbByte + vbArray
End Function

Function IsDigit(V) As Boolean
IsDigit = "0" <= V And V <= "9"
End Function

Function IsDte(V) As Boolean
IsDte = VarType(V) = vbDate
End Function

Function IsEq(A, B) As Boolean
If Not IsEqTy(A, B) Then Exit Function
Select Case True
Case IsArray(A): IsEq = IsEqAy(A, B)
Case IsDic(A): IsEq = IsEqDic(CvDic(A), CvDic(B))
Case IsObject(A): IsEq = ObjPtr(A) = ObjPtr(B)
Case Else: IsEq = A = B
End Select
End Function

Function IsEqDic(V As Dictionary, B As Dictionary) As Boolean
If V.Count <> B.Count Then Exit Function
If V.Count = 0 Then IsEqDic = True: Exit Function
Dim K1, k2
K1 = AySrtQ(V.Keys)
k2 = AySrtQ(B.Keys)
If Not IsEqAy(K1, k2) Then Exit Function
Dim K
For Each K In K1
   If B(K) <> V(K) Then Exit Function
Next
IsEqDic = True
End Function

Function IsEqTy(V, B) As Boolean
IsEqTy = VarType(V) = VarType(B)
End Function
Function IsInt(V) As Boolean
IsInt = VarType(V) = vbInteger
End Function
Function IsIntAy(V) As Boolean
IsIntAy = VarType(V) = vbArray + vbInteger
End Function

Function IsItr(V) As Boolean
IsItr = TypeName(V) = "Collection"
End Function

Function IsLetter(V$) As Boolean
IsLetter = IsAscLetter(Asc(V))
End Function

Function IsLines(V) As Boolean
If Not IsStr(V) Then Exit Function
IsLines = HasSubStr(V, vbLf)
End Function

Function IsLinesAy(V) As Boolean
If Not IsArray(V) Then Exit Function
If Not IsSy(V) Then Exit Function
Dim L
For Each L In Itr(V)
    If IsLines(L) Then IsLinesAy = True: Exit Function
Next
End Function

Function IsLng(V) As Boolean
IsLng = VarType(V) = vbLong
End Function

Function IsLngAy(V) As Boolean
IsLngAy = VarType(V) = vbArray + vbLong
End Function

Function IsNE(V, B) As Boolean
IsNE = Not IsEq(V, B)
End Function

Function IsNoLinMd(V As CodeModule) As Boolean
IsNoLinMd = V.CountOfLines = 0
End Function

Function IsNBStr(V) As Boolean
If Not IsStr(V) Then Exit Function
IsNBStr = V <> ""
End Function

Function IsNothing(V) As Boolean
If Not IsObject(V) Then Exit Function
IsNothing = ObjPtr(V) = 0
End Function

Function IsObjAy(V) As Boolean
IsObjAy = VarType(V) = vbArray + vbObject
End Function

Function IsPrimTy(Ty$) As Boolean
Select Case Ty
Case _
   "Boolean", _
   "Byte", _
   "Currency", _
   "Date", _
   "Decimal", _
   "Double", _
   "Integer", _
   "Long", _
   "Single", _
   "String"
   IsPrimTy = True
End Select
End Function
Function IsPrim(V) As Boolean
Select Case VarType(V)
Case _
   VbVarType.vbBoolean, _
   VbVarType.vbByte, _
   VbVarType.vbCurrency, _
   VbVarType.vbDate, _
   VbVarType.vbDecimal, _
   VbVarType.vbDouble, _
   VbVarType.vbInteger, _
   VbVarType.vbLong, _
   VbVarType.vbSingle, _
   VbVarType.vbString
   IsPrim = True
End Select
End Function
' sdlf _
sdfsdf
Function IsPun(V$) As Boolean
If IsLetter(V) Then Exit Function
If IsDigit(V) Then Exit Function
If V = vbCr Then Exit Function
If V = vbLf Then Exit Function
If V = "_" Then Exit Function
IsPun = True
End Function

Function IsQted(S, Q1$, Optional ByVal Q2$) As Boolean
If Q2 = "" Then Q2 = Q1
If FstChr(S) <> Q1 Then Exit Function
IsQted = LasChr(S) = Q2
End Function

Function IsSngQRmk(S) As Boolean
IsSngQRmk = FstChr(LTrim(S)) = "'"
End Function

Function IsSngQted(S) As Boolean
IsSngQted = IsQted(S, "'")
End Function

Function IsSomFething(V) As Boolean
IsSomFething = Not IsNothing(V)
End Function
Function IsNeedQte(S) As Boolean
If IsSqBktQted(S) Then Exit Function
Select Case True
Case IsAscDig(Asc(FstChr(S))), HasSpc(S), HasDot(S), HasHyphen(S), HasPound(S): IsNeedQte = True
End Select
End Function
Function IsDtezS(S$) As Boolean
On Error GoTo X
Dim D As Date: D = S
IsDtezS = True
X:
End Function

Function IsStr(V) As Boolean
IsStr = VarType(V) = vbString
End Function

Function IsStrAy(V) As Boolean
IsStrAy = VarType(V) = vbArray + vbString
End Function

Function IsEmpSy(V) As Boolean
If Not IsSy(V) Then Exit Function
IsEmpSy = Si(V) = 0
End Function

Function IsLy(V) As Boolean
If Not IsLy(V) Then Exit Function
Dim L: For Each L In Itr(V)
    If IsLin(L) Then Exit Function
Next
End Function

Function IsLin(V) As Boolean
If HasSubStr(V, vbCr) Then Exit Function
If HasSubStr(V, vbLf) Then Exit Function
IsLin = True
End Function

Function IsSy(V) As Boolean
IsSy = IsStrAy(V)
End Function

Function IsTglBtn(V) As Boolean
IsTglBtn = TypeName(V) = "ToggleButton"
End Function

Function IsVbTyNum(V As VbVarType) As Boolean
Select Case V
Case vbInteger, vbLong, vbDouble, vbSingle, vbDouble: IsVbTyNum = True: Exit Function
End Select
End Function

Function IsVdtLyDicStr(LyDicStr$) As Boolean
If Left(LyDicStr, 3) <> "***" Then Exit Function
Dim I, K$(), Key$
For Each I In SplitCrLf(LyDicStr$)
   If Left(I, 3) = "***" Then
       Key = Mid(I, 4)
       If HasEle(K, Key) Then Exit Function
       Push K, Key
   End If
Next
IsVdtLyDicStr = True
End Function

Function IsWhiteChr(V) As Boolean
Select Case Left(V, 1)
Case " ", vbCr, vbLf, vbTab: IsWhiteChr = True
End Select
End Function

Sub ZIsSy()
Dim V$()
Dim B: B = V
Dim C()
Dim D
Ass IsSy(V) = True
Ass IsSy(B) = True
Ass IsSy(C) = False
Ass IsSy(D) = False
End Sub

Sub Z_IsStrAy()
Dim V$()
Dim B: B = V
Dim C()
Dim D
Ass IsStrAy(V) = True
Ass IsStrAy(B) = True
Ass IsStrAy(C) = False
Ass IsStrAy(D) = False
End Sub

Sub Z_IsVdtLyDicStr()
Ass IsVdtLyDicStr(LineszVbl("***ksdf|***ksdf1")) = True
Ass IsVdtLyDicStr(LineszVbl("***ksdf|***ksdf")) = False
Ass IsVdtLyDicStr(LineszVbl("**ksdf|***ksdf")) = False
Ass IsVdtLyDicStr(LineszVbl("***")) = True
Ass IsVdtLyDicStr("**") = False
End Sub


Function IsAllBlnkSy(V$()) As Boolean
Dim I
For Each I In V
    If Trim(I) <> "" Then Exit Function
Next
IsAllBlnkSy = True
End Function

Function IsBlnk(V) As Boolean
Select Case True
Case IsBlnkStr(V), IsNull(V), IsEmpty(V), IsMissing(V): IsBlnk = True
End Select
End Function

Function IsBlnkStr(V) As Boolean
If IsStr(V) Then
    If Trim(V) = "" Then IsBlnkStr = True
End If
End Function
Function IsBet(V, A, B) As Boolean
If A > V Then Exit Function
If V > B Then Exit Function
IsBet = True
End Function
Function IsErObj(A) As Boolean
IsErObj = TypeName(A) = "Error"
End Function
Function IsEmp(A) As Boolean
Select Case True
Case IsStr(A):    IsEmp = Trim(A) = ""
Case IsArray(A):  IsEmp = Si(A) = 0
Case IsEmpty(A), IsNothing(A), IsMissing(A), IsNull(A): IsEmp = True
End Select
End Function

Function IsNBet(V, A, B) As Boolean
IsNBet = Not IsBet(V, A, B)
End Function

Function IsSqBktQted(S) As Boolean
IsSqBktQted = IsQted(S, "[", "]")
End Function

Function IsIxyOut(Ixy, U&) As Boolean
Dim Ix
For Each Ix In Itr(Ixy)
    If 0 > Ix Then IsIxyOut = True: Exit Function
    If Ix > U Then IsIxyOut = True: Exit Function
Next
End Function


Function IsEqStr(A, B, Optional C As VbCompareMethod = vbBinaryCompare) As Boolean
IsEqStr = StrComp(A, B, C) = 0
End Function

Function IsDteStr(S) As Boolean
On Error GoTo X
Dim A As Date: A = S
IsDteStr = True
Exit Function
X:
End Function

Function IsDblStr(S) As Boolean
On Error GoTo X
Dim A#: A = S
IsDblStr = True
Exit Function
X:
End Function

Function IsAySam(A, B) As Boolean
IsAySam = IsEqDic(DiKqCnt(A), DiKqCnt(B))
End Function

Function IsEqzAllEle(Ay) As Boolean
If Si(Ay) <= 1 Then IsEqzAllEle = True: Exit Function
Dim A0, J&
A0 = Ay(0)
For J = 1 To UB(Ay)
    If A0 <> Ay(J) Then Exit Function
Next
IsEqzAllEle = True
End Function

Function IsEqSi(A, B) As Boolean
IsEqSi = Si(A) = Si(B)
End Function

Function IsNeAy(A, B) As Boolean
IsNeAy = Not IsEqAy(A, B)
End Function

Function IsEqDy(A, B) As Boolean
IsEqDy = IsEqAy(A, B)
End Function
