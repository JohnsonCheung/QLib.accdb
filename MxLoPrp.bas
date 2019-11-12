Attribute VB_Name = "MxLoPrp"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxLoPrp."
Function EntColRgzLc(L As ListObject, C) As Range
Set EntColRgzLc = ColRgzLc(L, C).EntireColumn
End Function
Function ColRgzLc(L As ListObject, C) As Range
Set ColRgzLc = Lc(L, C).Range
End Function
Function CnozLc%(L As ListObject, C)
CnozLc = Lc(L, C).Index
End Function

Function LcAft(L As ListObject, AftC) As ListColumn
Dim C%: C = CnozLc(L, AftC) + 1
Set LcAft = Lc(L, C)
End Function

Function Lc(L As ListObject, C) As ListColumn
Set Lc = L.ListColumns(C)
End Function

Function ColRgzAftLc(L As ListObject, Aft) As Range
Set ColRgzAftLc = Lc(L, CnozLc(L, Aft) + 1).Range
End Function

Function DtaColRgzLc(L As ListObject, C) As Range
Set DtaColRgzLc = Lc(L, C).DataBodyRange
End Function

Function LoRno&(L As ListObject, Rg As Range)
':LoRno :Row-No #Listobject-Row-No# ! Fm 1-L.ListRows.Count, 0 will ix not found
If Not HasRg(L, Rg) Then Exit Function
LoRno = Rg.Row - L.DataBodyRange.Row + 1
End Function
Function LozHasRg(Rg As Range) As ListObject
Dim RC As RC: RC = RCzRg(Rg)
Dim L As ListObject: For Each L In WszRg(Rg).ListObjects
    If HasRC(RRCCzLo(L), RC) Then Set LozHasRg = L: Exit Function
Next
End Function
Function HasRg(L As ListObject, Rg As Range) As Boolean
'Ret :True if @L-data-range has the @Rg-A1
HasRg = HasRC(RRCCzLo(L), RCzRg(Rg))
End Function

Function RRCCzLo(L As ListObject) As RRCC
RRCCzLo = RRCCzRg(L.DataBodyRange)
End Function

Function DyzLo(L As ListObject) As Variant()
DyzLo = DyzSq(CvSq(L.DataBodyRange.Value))
End Function

Function LasRow(L As ListObject) As Range
If L.ListRows.Count = 0 Then Thw CSub, "There is no LasRow in L", "Lon Wsn", L.Name, WszLo(L).Name
Set LasRow = L.ListRows(L.ListRows.Count).Range
End Function

Function LasRowCell(L As ListObject, C) As Range
Dim Ix%: Ix = L.ListColumns(C).Index
Set LasRowCell = RgRC(LasRow(L), 1, L.ListColumns(C).Index)
End Function

Function LoCC(L As ListObject, C1$, C2$) As Range
Dim A%, b%
With L
    A = .ListColumns(C1).Index
    b = .ListColumns(C2).Index
    Set LoCC = RgCC(.DataBodyRange, A, b)
End With
End Function

Function WsnzLo$(A As ListObject)
WsnzLo = WszLo(A).Name
End Function

