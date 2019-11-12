Attribute VB_Name = "MxXlsRfh"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsRfh."
Sub RfhPc(A As PivotCache)
A.MissingItemsLimit = xlMissingItemsNone
A.Refresh
End Sub

Sub RfhFx(Fx, Fb)
RfhWb(WbzFx(Fx), Fb).Close SaveChanges:=True
End Sub

Sub RfhWs(A As Worksheet)
Dim Q As QueryTable: For Each Q In A.QueryTables: Q.BackgroundQuery = False: Q.Refresh: Next
Dim P As PivotTable: For Each P In A.PivotTables: P.Update: Next
Dim L As ListObject: For Each L In A.ListObjects: L.Refresh: Next
End Sub

Function RfhWb(Wb As Workbook, Fb) As Workbook
RplLozFb Wb, Fb
Dim C As WorkbookConnection
Dim P As PivotCache, W As Worksheet
'For Each C In Wb.Connections: RfhWc C, Fb:                                          Next
For Each P In Wb.PivotCaches: P.MissingItemsLimit = xlMissingItemsNone: P.Refresh:  Next
For Each W In Wb.Sheets:      RfhWs W:                                              Next
FmtLoBStd Wb
ClsWczWb Wb
DltWc Wb
Set RfhWb = Wb
End Function

