Attribute VB_Name = "MxCrtLo"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxCrtLo."

Function CrtLozSq(Sq(), At As Range, Optional Lon$) As ListObject
Set CrtLozSq = CrtLo(RgzSq(Sq(), At), Lon)
End Function

Function CrtLo(Rg As Range, Optional Lon$) As ListObject
Dim S As Worksheet: Set S = WszRg(Rg)
Dim O As ListObject: Set O = S.ListObjects.Add(xlSrcRange, Rg, , xlYes)
BdrAround Rg
Rg.EntireColumn.AutoFit
SetLon O, Lon
Set CrtLo = O
End Function

Function CrtLozDrs(D As Drs, At As Range, Optional Lon$) As ListObject
Set CrtLozDrs = CrtLo(RgzDrs(D, At), Lon)
End Function

Function CrtEmpLo(At As Range, FF$, Optional Lon$) As ListObject
Set CrtEmpLo = CrtLo(RgzAyH(SyzSS(FF), At), Lon)
End Function

Function RgzSq(Sq(), At As Range) As Range
If Si(Sq) = 0 Then
    Set RgzSq = A1zRg(At)
    Exit Function
End If
Dim O As Range
Set O = ResiRg(At, Sq)
O.MergeCells = False
O.Value = Sq
Set RgzSq = O
End Function

Sub CrtLozDbt(D As Database, T, At As Range, Optional AddgWay As EmAddgWay)
Select Case AddgWay
Case EmAddgWay.EiSqWay: CrtLozSq SqzT(D, T), At
Case EmAddgWay.EiWcWay: CrtLozFbt At, D.Name, T
Case Else: Thw CSub, "Invalid AddgWay"
End Select
End Sub

Sub CrtLozDbtWs(D As Database, T, Ws As Worksheet)
CrtLozDbt D, T, A1zWs(Ws)
End Sub

