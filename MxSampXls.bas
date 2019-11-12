Attribute VB_Name = "MxSampXls"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxSampXls."
Property Get SampLo() As ListObject
Set SampLo = CrtLo(RgzSq(SampSqWithHdr, NewA1), "Sample")
End Property
Property Get SampLoVis() As ListObject
Set SampLoVis = ShwLo(SampLo)
End Property

Function SampPt() As PivotTable
Set SampPt = PtzRg(SampRg)
End Function
Function SampRg() As Range
Set SampRg = ShwRg(CrtLozSq(SampSq, NewA1))
End Function

Property Get SampLofTp() As String()
Dim O$()
PushI O, "Lo  Nm     *Nm"
PushI O, "Lo  Fld    *Fld.."
PushI O, "Align Left   *Fld.."
PushI O, "Align Right  *Fld.."
PushI O, "Align Center *Fld.."
PushI O, "Bdr Left   *Fld.."
PushI O, "Bdr Right  *Fld.."
PushI O, "Bdr Col    *Fld.."
PushI O, "Tot Sum    *Fld.."
PushI O, "Tot Avg    *Fld.."
PushI O, "Tot Cnt    *Fld.."
PushI O, "Fmt *Fmt   *Fld.."
PushI O, "Wdt *Wdt   *Fld.."
PushI O, "Lvl *Lvl   *Fld.."
PushI O, "Cor *Cor   *Fld.."
PushI O, "Fml *Fld   *Formula"
PushI O, "Bet *Fld   *Fld1 *Fld2"
PushI O, "Tit *Fld   *Tit"
PushI O, "Lbl *Fld   *Lbl"
SampLofTp = O
End Property



Property Get SampWs() As Worksheet
Dim O As Worksheet
Set O = NewWs
LozDrs SampDrs, WsRC(O, 2, 2)
Set SampWs = O
ShwWs O
End Property
