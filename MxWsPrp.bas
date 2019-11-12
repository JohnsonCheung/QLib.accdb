Attribute VB_Name = "MxWsPrp"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxWsPrp."

Function HasLo(S As Worksheet, Lon$, Optional IsInf As Boolean) As Boolean
Dim O As Boolean: O = HasItn(S.ListObjects, Lon)
If Not O Then
    InfLin CSub, "No Lo in Ws", "Lon Wsn", Lon, S.Name
End If
End Function

Function LasCell(S As Worksheet) As Range
Set LasCell = S.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function LasRno&(S As Worksheet)
LasRno = LasCell(S).Row
End Function

Function LasCno%(S As Worksheet)
LasCno = LasCell(S).Column
End Function

Function PtNyzWs(S As Worksheet) As String()
PtNyzWs = Itn(S.PivotTables)
End Function

