Attribute VB_Name = "MxRegApp"
Option Compare Text
Option Explicit
Const CLib$ = "QApp."
Const CMod$ = CLib & "MxRegApp."
Const H$ = "C:\Users\user\Desktop\MHD\SAPAccessReports\"
Public Const AppStkShpRateFb$ = H & "StockShipRate\StockShipRate\StockShipRate (ver 1.0).accdb"
Public Const AppTaxExpCmpFb$ = H & "TaxExpCmp\TaxExpCmp\TaxExpCmp v1.3.accdb"
Public Const AppStkShpCstFb$ = H & "StockShipCost\StockShipCost (ver 1.0).accdb"
Public Const AppTaxRateAlertFb$ = H & "TaxRateAlert\TaxRateAlert\TaxRateAlert (ver 1.3).accdb"
Public Const AppJJFb$ = H & "TaxExpCmp\TaxExpCmp\PgmObj\Lib\jj.accdb"

Function AppFbAy() As String()
PushI AppFbAy, AppJJFb
PushI AppFbAy, AppStkShpCstFb
PushI AppFbAy, AppStkShpRateFb
PushI AppFbAy, AppTaxExpCmpFb
PushI AppFbAy, AppTaxRateAlertFb
End Function

Property Get MHDAppFbDic() As Dictionary
Const A$ = "N:\SAPAccessReports\"
Dim X As New Bfr
X.Var "Duty     " & A & "DutyPrepay\.accdb"
X.Var "SkHld    " & A & "StkHld\.accdb"
X.Var "ShpRate  " & A & "DutyPrepay\StockShipRate_Data.accdb"
X.Var "ShpCst   " & A & "StockShipCost\.accdb"
X.Var "TaxCmp   " & A & "TaxExpCmp\.accdb"
X.Var "TaxAlert " & A & "TaxRateAlert\.accdb"
Set MHDAppFbDic = Dic(X.Ly)
End Property
