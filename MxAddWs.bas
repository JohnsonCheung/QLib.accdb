Attribute VB_Name = "MxAddWs"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxAddWs."
Function AddWszSq(B As Workbook, Sq(), Optional Wsn$) As Worksheet
Dim A1 As Range: Set A1 = A1zWs(AddWs(B, Wsn))
CrtLozSq Sq, A1
Set AddWszSq = WszRg(A1)
End Function

Function AddWszDbt1(B As Workbook, Db As Database, T, Optional Wsn$, Optional AddgWay As EmAddgWay) As Worksheet
If AddgWay = EiSqWay Then
    Set AddWszDbt1 = AddWszDbt(B, Db, T, Wsn, AddgWay)
Else
    Set AddWszDbt1 = AddWszSq(B, SqzT(Db, T), Wsn)
End If
End Function

Function AddWszDrs(B As Workbook, D As Drs, Optional Wsn$) As Worksheet
Set AddWszDrs = AddWszSq(B, SqzDrs(D), Wsn)
End Function

Sub AddWszDbTny(B As Workbook, Db As Database, Tny$(), Optional AddgWay As EmAddgWay)
Dim T$, I
For Each I In Tny
    T = I
    AddWszDbt B, Db, T, , AddgWay
Next
End Sub

Function AddWszDt(B As Workbook, Dt As Dt) As Worksheet
Dim O As Worksheet
Set O = AddWs(B, Dt.DtNm)
LozDrs DrszDt(Dt), A1zWs(O)
Set AddWszDt = O
End Function

Function AddWs(A As Workbook, Optional Wsn$, Optional Pos As EmWsPos, Optional Aft$, Optional Bef$) As Worksheet
Dim O As Worksheet
DltWsIf A, Wsn
Select Case True
Case Pos = EiBeg:  Set O = A.Sheets.Add(FstWs(A))
Case Pos = EiEnd:  Set O = A.Sheets.Add(, LasWs(A))
Case Pos = EiRfWs And Aft <> "": Set O = A.Sheets.Add(, A.Sheets(Aft))
Case Pos = EiRfWs And Bef <> "": Set O = A.Sheets.Add(A.Sheets(Bef))
Case Else: Stop
End Select
SetWsn O, Wsn
Set AddWs = O
End Function

