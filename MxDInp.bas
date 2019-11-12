Attribute VB_Name = "MxDInp"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxDInp."

Function SampDoInp() As Drs
Erase XX
X "MB52 C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\MB52 2018-07-30.xls"
X "UOM  C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\sales text.xlsx"
X "ZHT1 C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\ZHT1.XLSX"
SampDoInp = DrszTRst("Inpn Ffn", XX)
End Function

Function EoDrsMsg(A As Drs, Msg$) As String()
If NoReczDrs(A) Then Exit Function
Erase XX
XLin Msg
XDrs A
XLin
EoDrsMsg = XX
End Function

Sub Z_EoDInp()
Brw EoDInp(SampDoInp)
End Sub

Function EoDInp(DInp As Drs) As String()
'@DInp
Dim E1$(), E2$(), E3$()
E1 = EoColDup(DInp, "Ffn")
E2 = EoColDup(DInp, "Inpn")
E3 = EoFfnMiszD(DInp)
EoDInp = Sy(E1, E2, E3)
End Function

Function EoFfnMiszD(Wi_Ffn As Drs) As String()
Dim I%: I = IxzAy(Wi_Ffn.Fny, "Ffn")
Dim Dr, Dy(): For Each Dr In Itr(Wi_Ffn.Dy)
    If NoFfn(Dr(I)) Then PushI Dy, Dr
Next
Dim b As Drs: b = Drs(Wi_Ffn.Fny, Dy)
EoFfnMiszD = EoDrsMsg(b, "File not exist")
End Function

