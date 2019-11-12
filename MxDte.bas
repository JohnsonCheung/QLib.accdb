Attribute VB_Name = "MxDte"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxDte."

Property Get CurMon() As Byte
CurMon = Month(Now)
End Property

Function NxtMonzM(M As Byte) As Byte
NxtMonzM = IIf(M = 12, 1, M + 1)
End Function

Function PrvMonzM(M As Byte) As Byte
PrvMonzM = IIf(M = 1, 12, M - 1)
End Function

Function FstDteOfMon(A As Date) As Date
FstDteOfMon = DateSerial(Year(A), Month(A), 1)
End Function

Function IsVdtDte(A) As Boolean
On Error Resume Next
IsVdtDte = Format(CDate(A), "YYYY-MM-DD") = A
End Function

Function LasDteOfMon(A As Date) As Date
LasDteOfMon = PrvDte(FstDteOfMon(NxtMon(A)))
End Function

Function NxtMon(A As Date) As Date
NxtMon = DateTime.DateAdd("M", 1, A)
End Function

Function IsHHMMDD(S) As Boolean
Select Case True
Case _
    Len(S) <> 6, _
    Not IsHH(Left(S, 2)), _
    Not Is0059(Mid(S, 3, 2)), _
    Not Is0059(Right(S, 2))
Case Else: IsHHMMDD = True
End Select
End Function

Function IsHH(S) As Boolean
Select Case True
Case _
    Len(S) <> 2, _
    Not IsAllDig(S), _
    "00" > S, S > "23"
Case Else: IsHH = True
End Select
End Function

Function Is0059(S) As Boolean
Select Case True
Case _
    Len(S) <> 2, _
    Not IsAllDig(S), _
    "00" > S, S > "59"
Case Else: Is0059 = True
End Select
End Function

Function IsYYYYMMDD(S) As Boolean
If Len(S) <> 8 Then Exit Function
If Not IsYYYY(Left(S, 4)) Then Exit Function
If Not IsMM(Mid(S, 5, 2)) Then Exit Function
If Not IsDD(Right(S, 2)) Then Exit Function
IsYYYYMMDD = True
End Function
Function IsAllDig(S) As Boolean
Dim J%
For J = 1 To Len(S)
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then Exit Function
Next
IsAllDig = True
End Function
Function IsMM(S) As Boolean
If Len(S) <> 2 Then Exit Function
If Not IsAllDig(S) Then Exit Function
If S < "00" Then Exit Function
If S > "12" Then Exit Function
IsMM = True
End Function
Function IsYYYY(S) As Boolean
Select Case True
Case Len(S) <> 4, Not IsAllDig(S), S < "2000"
Case Else: IsYYYY = True
End Select
End Function
Function IsDD(S) As Boolean
Select Case True
Case Len(S) <> 2, Not IsAllDig(S), S < "00", "31" < S
Case Else: IsDD = True
End Select
End Function

Function PrvDte(A As Date) As Date
PrvDte = DateAdd("D", -1, A)
End Function

Function YYMM$(A As Date)
YYMM = Right(Year(A), 2) & Format(Month(A), "00")
End Function

Function FstDtezYYMM(YYMM) As Date
FstDtezYYMM = DateSerial(Left(YYMM, 2), Mid(YYMM, 3, 2), 1)
End Function

Function IsVdtYYYYMMDD(A) As Boolean
On Error Resume Next
IsVdtYYYYMMDD = Format(CDate(A), "YYYY-MM-DD") = A
End Function

Function FstDtezYM(Y As Byte, M As Byte) As Date
FstDtezYM = DateSerial(2000 + Y, M, 1)
End Function

Function LasDtezYM(Y As Byte, M As Byte) As Date
LasDtezYM = NxtMon(FstDtezYM(Y, M))
End Function

Function YofNxtMzYM(Y As Byte, M As Byte) As Byte
YofNxtMzYM = IIf(M = 12, Y + 1, Y)
End Function

Function YofPrvMzYM(Y As Byte, M As Byte) As Byte
YofPrvMzYM = IIf(M = 1, Y - 1, Y)
End Function

Property Get CurY() As Byte
CurY = CurYY - 2000
End Property

Property Get CurYY%()
CurYY = Year(Now)
End Property

Function IsTimStr(Str) As Boolean
If Len(Str) <> 17 Then Exit Function
Select Case True
Case IsHHMMSS(Right(Str, 6)), IsYYYYDashMMDashMM(Left(Str, 10)): IsTimStr = True
End Select
End Function
Function IsHHMMSS(HHMMSS$) As Boolean
On Error GoTo X
Dim T As Date: T = CDate(Format(HHMMSS, "00:00:00"))
IsHHMMSS = Format(T, "HHMMSS")
Exit Function
X:
End Function
Function TimId$(A As Date)
TimId = Format(A, "YYYYMMDD_HHMMSS")
End Function
Function IsYYYYDashMMDashMM(A$) As Boolean
Select Case True
Case Len(A) <> 10, Mid(A, 5, 1) <> "-", Mid(A, 8, 1) <> "-": Exit Function
End Select
On Error GoTo X
Dim T As Date: T = CDate(A)
IsYYYYDashMMDashMM = Format(T, "YYYY-MM-DD")
Exit Function
X:
End Function

Function TimNm$(A As Date, Optional Pfx$ = "N")
TimNm = Pfx & TimId(A)
End Function

Function TimStr$(A As Date)
TimStr = Format(A, "YYYY-MM-DD HHMMSS")
End Function
Function NowId$()
NowId = TimId(Now)
End Function

Property Get NowStr$()
NowStr = TimStr(Now)
End Property



Function CvDbl(S, Optional Fun$)
'Ret : a dbl of @S if can be converted, otherwise empty and debug.print S$
On Error GoTo X
CvDbl = CDbl(S)
Exit Function
X: If Fun <> "" Then Inf CSub, "str[" & S & "] cannot cv to dbl, emp is ret"
End Function
