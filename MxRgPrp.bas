Attribute VB_Name = "MxRgPrp"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxRgPrp."

Function RgCRR(A As Range, C, R1, R2) As Range
Set RgCRR = RgRCRC(A, R1, C, R2, C)
End Function

Function EntRgC(A As Range, C) As Range
Set EntRgC = RgC(A, C).EntireColumn
End Function

Function EntRgRR(A As Range, R1, R2) As Range
Set EntRgRR = RgRR(A, R1, R2).EntireRow
End Function

Function FstColRg(A As Range) As Range
Set FstColRg = RgC(A, 1)
End Function

Function FstRowRg(A As Range) As Range
Set FstRowRg = RgR(A, 1)
End Function

Function IsRgHBar(A As Range) As Boolean
IsRgHBar = A.Rows.Count = 1
End Function

Function IsRgVBar(A As Range) As Boolean
IsRgVBar = A.Columns.Count = 1
End Function


Function NColzRg%(A As Range)
NColzRg = A.Columns.Count
End Function



Function RgMoreBelow(A As Range, Optional N% = 1)
Set RgMoreBelow = RgRR(A, 1, A.Rows.Count + N)
End Function
Function RgMoreTop(A As Range, Optional N = 1)
Dim O As Range
Set O = RgRR(A, 1 - N, A.Rows.Count)
Set RgMoreTop = O
End Function

Function NRowzRg&(A As Range)
NRowzRg = A.Rows.Count
End Function

Function RgAy(Rg As Range, RC As Drs) As Range()
'Fm RC : .. R C .. @@
Dim IxR%, IxC%: AsgIx RC, "R C", IxR, IxC
Dim Dr: For Each Dr In Itr(RC.Dy)
    Dim R&: R = Dr(IxR)
    Dim C%: C = Dr(IxC)
    PushObj RgAy, RgRC(Rg, R, C)
Next
End Function

Function RgR(A As Range, Optional R = 1)
Set RgR = RgRR(A, R, R)
End Function

Function SqzRg(A As Range) As Variant()
If A.Columns.Count = 1 Then
    If A.Rows.Count = 1 Then
        Dim O()
        ReDim O(1 To 1, 1 To 1)
        O(1, 1) = A.Value
        SqzRg = O
        Exit Function
    End If
End If
SqzRg = A.Value
End Function

Function RgRC(A As Range, R, C) As Range
Set RgRC = A.Cells(R, C)
End Function

Function RgRCC(A As Range, R, C1, C2) As Range
Set RgRCC = RgRCRC(A, R, C1, R, C2)
End Function

Function RgRCRC(A As Range, R1, C1, R2, C2) As Range
Set RgRCRC = WszRg(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2))
End Function

Function RgRR(A As Range, R1, R2) As Range
Set RgRR = RgRCRC(A, R1, 1, R2, NColzRg(A))
End Function

Function IsA1(A As Range) As Boolean
If A.Row <> 1 Then Exit Function
If A.Column <> 1 Then Exit Function
IsA1 = True
End Function

Function WsRgAdr$(A As Range)
WsRgAdr = "'" & WszRg(A).Name & "'!" & A.Address
End Function

Function RCzRg(A As Range) As RC
With RCzRg
.R = A.Row
.C = A.Column
End With
End Function
Function RRCCzRg(A As Range) As RRCC
With RRCCzRg
.R1 = A.Row
.R2 = .R1 + A.Rows.Count - 1
.C1 = A.Column
.C2 = .C1 + A.Columns.Count - 1
End With
End Function

Function DrzRg(Rg As Range, Optional R = 1) As Variant()
DrzRg = DrzSq(SqzRg(RgR(Rg, R)))
End Function

Function A1Adr$(A As Range)
A1Adr = A1zRg(A).Address(External:=True)
End Function

