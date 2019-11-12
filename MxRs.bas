Attribute VB_Name = "MxRs"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxRs."

Function AsetzRs(Rs As DAO.Recordset, F$) As Aset
Set AsetzRs = New Aset
With Rs
    While Not .EOF
        AsetzRs.PushItm .Fields(F).Value
        .MoveNext
    Wend
End With
End Function

Sub AsgRs(A As DAO.Recordset, ParamArray OAp())
Dim F As DAO.Field, J%, U%
Dim Av(): Av = OAp
U = UB(Av)
For Each F In A.Fields
    OAp(J) = F.Value
    If J = U Then Exit Sub
    J = J + 1
Next
End Sub

Function AvRsCol(A As DAO.Recordset, Optional Fld = 0) As Variant()
AvRsCol = IntozRs(EmpAv, A, Fld)
End Function

Sub BrwRs(A As DAO.Recordset)
BrwDrs DrszRs(A)
End Sub

Sub BrwSngRec(A As DAO.Recordset)
BrwAy FmtRec(A)
End Sub

Function ColSetzRs(A As DAO.Recordset, Optional Fld = 0) As Aset
Set ColSetzRs = New Aset
With A
    While Not .EOF
        ColSetzRs.PushItm .Fields(Fld).Value
        .MoveNext
    Wend
End With
End Function

Function CsvLinzRs$(A As DAO.Recordset)
CsvLinzRs = CsvLinzFds(A.Fields)
End Function

Function CsvLyzRs(A As DAO.Recordset, Optional FF$) As String()
Dim Fny$(), Flds As Fields, F
Dim O$(), J&, I%, UFld%, Dr()
Fny = TermAy(FF)
UFld = A.Fields.Count - 1
While Not A.EOF
    J = J + 1
    If J Mod 5000 = 0 Then Debug.Print "CsvLinzRsLy: " & J
    If J > 100000 Then Stop
    ReDim Dr(UFld)
    I = 0
    Set Flds = A.Fields
    For Each F In Fny
        Dr(I) = Csv(Flds(F).Value)
        I = I + 1
    Next
    Push O, Join(Dr, ",")
    A.MoveNext
Wend
CsvLyzRs = O
End Function

Function CsvLyzRs1(A As DAO.Recordset) As String()
Dim O$(), J&, I%, UFld%, Dr(), F As DAO.Field
UFld = A.Fields.Count - 1
While Not A.EOF
    J = J + 1
    If J Mod 5000 = 0 Then Debug.Print "CsvLinzRsLy: " & J
    If J > 100000 Then Stop
    ReDim Dr(UFld)
    PushI CsvLyzRs1, CsvLinzRs(A)
    A.MoveNext
Wend
End Function

Function CvRs2(A) As DAO.Recordset2
Set CvRs2 = A
End Function
Function CvRs(A) As DAO.Recordset
Set CvRs = A
End Function

Sub DltRs(A As DAO.Recordset)
With A
    While Not .EOF
        .Delete
        .MoveNext
    Wend
End With
End Sub

Sub DmpRs(A As Recordset, FF$)
DmpAy CsvLyzRs(A, FF)
A.MoveFirst
End Sub

Function DrszRs(A As DAO.Recordset) As Drs
DrszRs = Drs(FnyzRs(A), DyoRs(A))
End Function

Function DrzRs(A As DAO.Recordset, Optional FF$) As Variant()
DrzRs = DrzFds(A.Fields, FF)
End Function

Function DrzRsFny(A As DAO.Recordset, Fny$()) As Variant()
Dim F
For Each F In Fny
    PushI DrzRsFny, EmptyIfNull(A.Fields(F).Value)
Next
End Function

Function DyoRs(A As DAO.Recordset, Optional IsIncFldn As Boolean) As Variant()
If IsIncFldn Then
    PushI DyoRs, FnyzRs(A)
End If
If Not HasRec(A) Then Exit Function
With A
    .MoveFirst
    While Not .EOF
        PushI DyoRs, DrzFds(.Fields)
        .MoveNext
    Wend
    .MoveFirst
End With
End Function

Function EmptyIfNull(V)
EmptyIfNull = IIf(IsNull(V), Empty, V)
End Function

Function FmtRec(A As DAO.Recordset)
FmtRec = LyzNyAv(FnyzRs(A), DrzRs(A))
End Function

Function FnyzRs(A As DAO.Recordset) As String()
FnyzRs = Itn(A.Fields)
End Function

Function HasRec(A As DAO.Recordset) As Boolean
HasRec = Not NoRec(A)
End Function

Function HasReczFEv(Rs As DAO.Recordset, F, Ev) As Boolean
With Rs
    If .BOF Then
        If .EOF Then Exit Function
    End If
    .MoveFirst
    While Not .EOF
        If .Fields(F) = Ev Then HasReczFEv = True: Exit Function
        .MoveNext
    Wend
End With
End Function

Sub InsRs(Rs As DAO.Recordset, Dr)
Rs.AddNew
SetRs Rs, Dr
Rs.Update
End Sub

Sub InsRszAp(Rs As DAO.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
InsRs Rs, Dr
End Sub

Sub InsRszDy(A As DAO.Recordset, Dy())
Dim Dr
With A
    For Each Dr In Itr(Dy)
        InsRs A, Dr
    Next
End With
End Sub

Function IntAyzRs(A As DAO.Recordset, Optional Fld = 0) As Integer()
IntAyzRs = IntozRs(IntAyzRs, A, Fld)
End Function

Function IntozRs(Into, Rs As Recordset, Optional Fld = 0)
IntozRs = ResiU(Into)
While Not Rs.EOF
    PushI IntozRs, Nz(Rs(Fld).Value, Empty)
    Rs.MoveNext
Wend
End Function

Function LngAyzRs(A As DAO.Recordset, Optional Fld = 0) As Long()
LngAyzRs = IntozRs(LngAyzRs, A, Fld)
End Function

Function NoRec(A As DAO.Recordset) As Boolean
If Not A.EOF Then Exit Function
If Not A.BOF Then Exit Function
NoRec = True
End Function

Function NReczRs&(A As DAO.Recordset)
Dim O&
With A
    .MoveFirst
    While Not .EOF
        O = O + 1
        .MoveNext
    Wend
    .MoveFirst
End With
NReczRs = O
End Function

Sub RsDlt(A As DAO.Recordset)
With A
    If .EOF Then Exit Sub
    If .BOF Then Exit Sub
    .Delete
End With
End Sub

Function RsLin(A As DAO.Recordset, Optional Sep$ = " ")
RsLin = Join(DrzRs(A), Sep)
End Function

Function RsMapJn(A As DAO.Recordset, Optional Sep$ = " ") As String()
Dim O$()
With A
    Push O, Join(FnyzRs(A), Sep)
    While Not .EOF
        Push RsMapJn, RsLin(A, Sep)
        .MoveNext
    Wend
End With
RsMapJn = O
End Function

Sub SetRs(Rs As DAO.Recordset, Dr)
If Si(Dr) <> Rs.Fields.Count Then
    Thw CSub, "Si of Rs & Dr are diff", _
        "Si-Rs and Si-Dr Rs-Fny Dr", Rs.Fields.Count, Si(Dr), Itn(Rs.Fields), Dr
End If
Dim V, J%
For Each V In Dr
    If IsEmpty(V) Then
        Rs(J).Value = Rs(J).DefaultValue
    Else
        Rs(J).Value = V
    End If
    J = J + 1
Next
End Sub

Function SqzRs(A As DAO.Recordset, Optional ExlFldNm As Boolean) As Variant()
SqzRs = SqzDy(DyoRs(A, ExlFldNm))
End Function


Function SyzRs(A As DAO.Recordset, Optional F = 0) As String()
Dim O$()
With A
    While Not .EOF
        Push O, .Fields(F).Value
        .MoveNext
    Wend
End With
SyzRs = O
End Function

Sub UpdRs(Rs As DAO.Recordset, Dr)
Rs.Edit
SetRs Rs, Dr
Rs.Update
End Sub

Sub UpdRszAp(Rs As DAO.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
UpdRs Rs, Dr
End Sub

Sub Y_AsgRs()
Dim Y As Byte, M As Byte
'AsgRs TblRs("YM"), Y, M
Stop
End Sub
