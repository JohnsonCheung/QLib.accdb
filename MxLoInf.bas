Attribute VB_Name = "MxLoInf"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxLoInf."
Public Const FFoLoInf$ = "Wsn Lon R C NR NC"
Function DyoLoInf(Wb As Workbook) As Variant()
Dim O()
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    PushI O, DyoLoInfzWs(Ws)
Next
DyoLoInf = O
End Function

Function DyoLoInfzWs(Ws As Worksheet) As Variant()
Dim Lo As ListObject: For Each Lo In Ws.ListObjects
    PushI DyoLoInfzWs, DroLoInf(Lo)
Next
End Function

Sub Z_DroLoInf()
Dim Lo As ListObject: Set Lo = SampLo
D DroLoInf(Lo)
ClsWbNoSav WbzLo(Lo)
End Sub

Function DroLoInf(L As ListObject) As Variant()
Dim Wsn: Wsn = WsnzLo(L)
Dim Lon$:: Lon = L.Name
Dim NR&: NR = NRowzLo(L)
Dim NC&: NC = L.ListColumns.Count
DroLoInf = Array(WsnzLo(L), L.Name, L.Range.Row, L.Range.Column, NR, NC)
End Function

Sub CrtLoInf(At As Range)
CrtLozDrs DoLoInf(WbzRg(At)), At
End Sub


Sub Z_DyoLoInf()
Dim Wb As Workbook: Set Wb = NewWb
AddWszSq Wb, SampSq
AddWszSq Wb, SampSq1
BrwSq DyoLoInf(Wb)
End Sub

Function DoLoInf(Wb As Workbook) As Drs
DoLoInf = DrszFF(FFoLoInf, DyoLoInf(Wb))
End Function

