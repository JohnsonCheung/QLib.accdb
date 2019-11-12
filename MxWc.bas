Attribute VB_Name = "MxWc"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxWc."
Function WszWc(Wc As WorkbookConnection) As Worksheet
Dim Wb As Workbook, Ws As Worksheet
Set Wb = Wc.Parent
Set Ws = AddWs(Wb, Wc.Name)
PutWc Wc, A1zWs(Ws)
Set WszWc = Ws
End Function

Sub Rpl_Wc_ByFb(Wc As WorkbookConnection, Fb)
CvCn(Wc.OLEDBConnection.ADOConnection).ConnectionString = AdoCnStrzFb(Fb)
End Sub


Function TxtWc(A As Workbook) As TextConnection
Dim C As WorkbookConnection
For Each C In A.Connections
    If Not IsNothing(TxtCnzWc(C)) Then
        Set TxtWc = C.TextConnection
        Exit Function
    End If
Next
Stop
'XHalt_Impossible CSub
End Function

Function TxtWcCnt%(A As Workbook)
Dim C As WorkbookConnection, Cnt%
For Each C In A.Connections
    If Not IsNothing(TxtCnzWc(C)) Then Cnt = Cnt + 1
Next
TxtWcCnt = Cnt
End Function

Function TxtWcStr$(A As Workbook)
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Dim T As TextConnection: Set T = TxtWc(A)
If IsNothing(T) Then Exit Function
TxtWcStr = T.Connection
End Function


Function WcnyzWb(A As Workbook) As String()
WcnyzWb = Itn(A.Connections)
End Function

Function WcStrAyzWbOLE(A As Workbook) As String()
Dim O() As OLEDBConnection
Dim Wc As WorkbookConnection: For Each Wc In A.Connections
    If Not IsNothing(Wc.OLEDBConnection) Then
        PushI WcStrAyzWbOLE, Wc.OLEDBConnection.Connection
    End If
Next
End Function

Function AddWc(A As Workbook, Fb, T) As WorkbookConnection
Set AddWc = A.Connections.Add2(T, T, WcCnStrzFb(Fb), T, XlCmdType.xlCmdTable)
End Function

Sub DltWc(A As Workbook)
Dim Wc As Excel.WorkbookConnection
For Each Wc In A.Connections
    Wc.Delete
Next
End Sub
Sub ClsWc(A As WorkbookConnection)
If IsNothing(A.OLEDBConnection) Then Exit Sub
CvCn(A.ODBCConnection.Connection).Close
End Sub

Sub ClsWczWb(Wb As Workbook)
Dim Wc As WorkbookConnection
For Each Wc In Wb.Connections
    ClsWc Wc
Next
End Sub


Sub SetWczFb(A As WorkbookConnection, ToUseFb$)
If IsNothing(A.OLEDBConnection) Then Exit Sub
Dim Cn$
#Const A = 2
#If A = 1 Then
    Dim S$
    S = A.OLEDBConnection.Connection
    Cn = RplBet(S, ToUseFb, "Data Source=", ";")
#End If
#If A = 2 Then
    Cn = OleCnStrzFb(ToUseFb)
#End If
A.OLEDBConnection.Connection = Cn
End Sub

Sub RfhWc(A As WorkbookConnection, ToUseFb$)
If IsNothing(A.OLEDBConnection) Then Exit Sub
SetWczFb A, ToUseFb
A.OLEDBConnection.BackgroundQuery = False
A.OLEDBConnection.Refresh
End Sub


Sub PutWc(Wc As WorkbookConnection, At As Range)
Dim Lo As ListObject
Set Lo = WszRg(At).ListObjects.Add(SourceType:=0, Source:=Wc.OLEDBConnection.Connection, Destination:=At)
With Lo.QueryTable
    .CommandType = xlCmdTable
    .CommandText = Wc.Name
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = True
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnINF = True
    .ListObject.DisplayName = Lon(Wc.Name)
    .Refresh BackgroundQuery:=False
End With
End Sub

Sub AddWczTT(ToWb As Workbook, FmFb, TT$)
Dim T$, I
For Each I In Ny(TT)
    T = I
    AddWc ToWb, FmFb, T
Next
End Sub
Function WcStrAyzFx(Fx) As String()
Dim B As Workbook: Set B = Xls.Workbooks.Open(Fx)
WcStrAyzFx = WcStrAyzWbOLE(B)
B.Close False
End Function

