VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaxCmpOupGenr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Implements IOupGenr
Const CLib$ = "QTaxCmp."
Const CMod$ = CLib & "TaxCmpOupGenr."
Const A$ = "A"
Sub GenOupTbl(D As Database)
IOupGenr_GenOupTbl D
End Sub
Sub IOupGenr_GenOupTbl(D As Database)
End Sub
