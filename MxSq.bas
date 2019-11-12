Attribute VB_Name = "MxSq"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxSq."
Sub SetSqr(OSq(), Dr, Optional R = 1, Optional NoTxtSngQ As Boolean)
Dim J&
If NoTxtSngQ Then
    For J = 0 To UB(Dr)
        If IsStr(Dr(J)) Then
            OSq(R, J + 1) = QteSng(CStr(Dr(J)))
        Else
            OSq(R, J + 1) = Dr(J)
        End If
    Next
Else
    For J = 0 To UB(Dr)
        OSq(R, J + 1) = Dr(J)
    Next
End If
End Sub

Sub PushSq(OSq(), Sq())
Dim NR&: NR = UBound(OSq, 1) + UBound(Sq, 1)
Dim NC&: NC = UBound(OSq, 2)
Dim NC2&: NC2 = UBound(Sq, 2)
If NC <> NC2 Then Thw CSub, "NC of { OSq, Sq } are dif", "OSq-NC Sq-NC", NC, NC2
ReDim Preserve OSq(1 To NR, 1 To NC)
Dim R&, C&
For R = 1 To NC2
    For C = 1 To NC
        OSq(R + NR, C) = Sq(R, C)
    Next
Next
End Sub

Function Sq(R&, C&) As Variant()
'Ret : a Sq(1 to @R, 1 to @C)
Dim O()
ReDim O(1 To R, 1 To C)
Sq = O
End Function

Function AddSngQtezSq(Sq())
Dim NC%, C%, R&, O
O = Sq
NC = UBound(Sq, 2)
For R = 1 To UBound(Sq, 1)
    For C = 1 To NC
        If IsStr(O(R, C)) Then
            O(R, C) = "'" & O(R, C)
        End If
    Next
Next
AddSngQtezSq = O
End Function
Function JnSq(Sq(), SepChr$) As String()
Dim NC&: NC = UBound(Sq, 2)
Dim R&
For R = 1 To UBound(Sq, 1)
    PushI JnSq, JnSqr(Sq, R, SepChr)
Next
End Function

Function JnSqr$(Sq(), R&, SepChr$)
JnSqr = Join(SyzSqr(Sq, R), SepChr)
End Function

Function SyzSqr(Sq(), R&) As String()
Dim J&
For J = 1 To UBound(Sq, 2)
    PushI SyzSqr, Sq(R, J)
Next
End Function

Function FmtSq(Sq(), Optional SepChr$ = " ") As String()
FmtSq = JnSq(AlignSq(Sq), SepChr)
End Function

Sub BrwSq(Sq())
Brw FmtSq(Sq)
End Sub

Function WdtzSqc%(Sq(), C&)
Dim R&, O%
For R = 1 To UBound(Sq, 1)
    O = Max(O, Len(Sq(R, C)))
Next
WdtzSqc = O
End Function
Sub AlignColzSq(OSq(), Sq(), C&, W%)
Dim R&
For R = 1 To UBound(Sq, 1)
    OSq(R, C) = Align(Sq(R, C), W)
Next
End Sub

Function ColzSq(Sq(), Optional C = 1) As Variant()
ColzSq = IntozSqc(EmpAv, Sq, C)
End Function

Function DrzSq(Sq(), Optional R = 1) As Variant()
DrzSq = IntozSqr(EmpAv, Sq, R)
End Function

Function IntozSqc(Into, Sq(), C)
Dim NR&: NR = UBound(Sq, 1)
Dim O:    O = ResiN(Into, NR)
Dim R&: For R = 1 To NR
    O(R - 1) = Sq(R, C)
Next
IntozSqc = O
End Function

Function F_Into_SelSq_ByR_AndCny(Into, Sq(), R, Cny%())
Dim NCol&:    NCol = UBound(Cny)
Dim O: O = Into: ReDim O(NCol - 1)
Dim C%: For C = 0 To NCol - 1
    O(C) = Sq(R, Cny(C))
Next
F_Into_SelSq_ByR_AndCny = O
End Function

Function IntozSqr(Into, Sq(), R)
Dim NCol&:    NCol = UBound(Sq, 2)
Dim O: O = Into: ReDim O(NCol - 1)
Dim C%: For C = 1 To NCol
    O(C - 1) = Sq(R, C)
Next
IntozSqr = O
End Function

Function IntozSqrCny(Into, Sq(), R, Cny)
Dim UCol%:    UCol = UBound(Cny)
Dim O: O = Into: ReDim O(UCol)
Dim C%: For C = 1 To UCol + 1
    O(C - 1) = Sq(R, C)
Next
IntozSqrCny = O
End Function

Function SyzSq(Sq(), Optional C& = 1) As String()
SyzSq = IntozSqc(EmpSy, Sq(), C)
End Function

Function DrzSqr(Sq(), Optional R = 1) As Variant()
DrzSqr = IntozSqr(EmpAv, Sq, R)
End Function

Function DrzSqrCny(Sq(), R, Cny) As Variant()
DrzSqrCny = IntozSqrCny(EmpAv, Sq, R, Cny)
End Function

Function F_Dr_SelSq_ByR_AndCny(Sq(), R, Cny%()) As Variant()
F_Dr_SelSq_ByR_AndCny = F_Into_SelSq_ByR_AndCny(EmpAv, Sq, R, Cny)
End Function

Function InsSqr(Sq(), Dr(), Optional Row& = 1)
Dim O(), C&, R&, NC&, NR&
NC = NColzSq(Sq)
NR = NRowzSq(Sq)
ReDim O(1 To NR + 1, 1 To NC)
For R = 1 To Row - 1
    For C = 1 To NC
        O(R, C) = Sq(R, C)
    Next
Next
For C = 1 To NC
    O(Row, C) = Dr(C - 1)
Next
For R = NR To Row Step -1
    For C = 1 To NC
        O(R + 1, C) = Sq(R, C)
    Next
Next
InsSqr = O
End Function

Function IsEqSq(A, B) As Boolean
Dim NR&, NC&
NR = UBound(A, 1)
NC = UBound(A, 2)
If NR <> UBound(B, 1) Then Exit Function
If NC <> UBound(B, 2) Then Exit Function
Dim R&, C&
For R = 1 To NR
    For C = 1 To NC
        If A(R, C) <> B(R, C) Then
            Exit Function
        End If
    Next
Next
IsEqSq = True
End Function

Function TermLinAyzSq(Sq()) As String()
Dim R&
For R = 1 To UBound(Sq(), 1)
    Push TermLinAyzSq, TermLin(DrzSqr(Sq, R))
Next
End Function

Function NColzSq&(Sq())
On Error Resume Next
NColzSq = UBound(Sq, 2)
End Function
Function NewLoSqAt(Sq(), At As Range) As ListObject
Set NewLoSqAt = CrtLo(RgzSq(Sq(), At))
End Function

Function NewLoSq(Sq(), Optional Wsn$ = "Data") As ListObject
Set NewLoSq = NewLoSqAt(Sq(), NewA1(Wsn))
End Function

Function WszSq(Sq(), Optional Wsn$) As Worksheet
Set WszSq = CrtLo(RgzSq(Sq(), NewA1(Wsn)))
End Function

Function NRowzSq&(Sq())
On Error Resume Next
NRowzSq = UBound(Sq, 1)
End Function


Function Transpose(Sq()) As Variant()
Dim NR&, NC&
NR = NRowzSq(Sq): If NR = 0 Then Exit Function
NC = NColzSq(Sq): If NC = 0 Then Exit Function
Dim O(), J&, I&
ReDim O(1 To NC, 1 To NR)
For J = 1 To NR
    For I = 1 To NC
        O(I, J) = Sq(J, I)
    Next
Next
Transpose = O
End Function


Function SampSq() As Variant()
Const NR% = 10
Const NC% = 10
Dim O(), R%, C%
ReDim O(1 To NR, 1 To NC)
SampSq = O
For R = 1 To NR
    For C = 1 To NC
        O(R, C) = R * 1000 + C
    Next
Next
SampSq = O
End Function


Function CvDte(S, Optional Fun$)
'Ret : a date fm @S if can be converted, otherwise empty and debug.print @S
On Error GoTo X
Dim O As Date: O = S
If SubStrCnt(S, "/") <> 2 Then GoTo X ' ! one [/]-str is cv to yyyy/mm, which is not consider as a dte.
'                                       ! so use 2-[/] to treat as a dte str.
If Year(O) < 2000 Then GoTo X         ' ! year < 2000, treat it as str or not
CvDte = O
Exit Function
X: If Fun <> "" Then Inf CSub, "str[" & S & "] cannot cv to dte, emp is ret"
End Function
Sub Z_SqStr()
Brw SqStrzDrs(DoMthP)
End Sub
Function SqStrzDy$(Dy())
SqStrzDy = SqStr(SqzDy(Dy))
End Function
Function SqStrzDrs$(D As Drs)
SqStrzDrs = SqStrzDy(D.Dy)
End Function
Sub Z_SqStrzWs()
Dim S As Worksheet: Set S = WsoMthcP
Dim Bef$: Bef = SqStrzWs(S)
EnsSprp S, "A", Bef
Dim Aft$: Aft = Sprp(S, "A")
If Bef <> Aft Then Stop
VcStr Bef
End Sub
Function SqStrzWs$(S As Worksheet)
SqStrzWs = SqStrzLo(FstLo(S))
End Function
Function SqStrzLo$(L As ListObject)
SqStrzLo = SqStrzRg(L.DataBodyRange)
End Function

Function CellStr$(V, Optional Fun$)
':CellStr: :S #Xls-Cell-Str# ! A str coming fm xls cell
Dim T$: T = TypeName(V)
Dim O$
Select Case T
Case "Boolean", "Long", "Integer", "Date", "Currency", "Single", "Double": CellStr = V
Case "String": If IsDblStr(V) Then CellStr = "'" & V Else CellStr = SlashCrLfTab(V)
Case Else: If Fun <> "" Then Inf Fun, "Val-of-TypeName[" & T & "] cannot cv to :CellStr"
End Select
End Function


Function SqStrzRg$(R As Range)
SqStrzRg = SqStr(SqzRg(R))
End Function

Function IsSqEmp(Sq()) As Boolean
Dim R&: For R = 1 To UBound(Sq, 1)
    Dim C%: For C = 1 To UBound(Sq, 2)
        If Not IsEmpty(Sq(R, C)) Then Exit Function
    Next
Next
IsSqEmp = True
End Function

Function DrszSq(SqWiHdr()) As Drs
Dim Fny$(): Fny = SyzSqr(SqWiHdr, 1)
Dim Dy()
    Dim R&: For R = 2 To UBound(SqWiHdr, 1)
        PushI Dy, DrzSqr(SqWiHdr, R)
    Next
DrszSq = Drs(Fny, Dy)
End Function

