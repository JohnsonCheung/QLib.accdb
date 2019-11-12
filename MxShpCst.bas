Attribute VB_Name = "MxShpCst"
Option Compare Text
Option Explicit
Const CNs$ = "App.ShpCst"
Const CLib$ = "QShpCst."
Const CMod$ = CLib & "MxShpCst."

Sub UomDoc()
#If False Then
InpX: [>UOM]     Material [Base Unit of Measure] [Material Description] [Unit per case]
Oup : UOM        Sku      SkuUOM                 Des                    Sc_U

Note on [Sales text.xls]
Col  Xls Title            FldName     Means
F    Base Unit of Measure SkuUOM      either COL (bottle) or PCE (set)
J    Unit per case        Sc_U        how many unit per AC
K    SC                   SC_U        how many unit per SC   ('no need)
L    COL per case         AC_B        how many bottle per AC
-----
Letter meaning
B = Bottle
AC = act case
SC = standard case
U = Unit  (Bottle(COL) or Set (PCE))

 "SC              as SC_U," & _  no need
 "[COL per case]  as AC_B," & _ no need
#End If
End Sub

Function EoMisWh8687(FxMB52$, Wsn$) As String()
If NReczFxw(FxMB52, Wsn, "Plant in ('8601','8701')") = 0 Then
    EoMisWh8687 = MoMisWh8687(FxMB52, Wsn)
End If
End Function

Function MoMisWh8687(FxMB52$, Wsn$) As String()
Const CSub$ = CMod & "EoMB52MissingWhs8601Or8701"
Const M$ = "Column-[Plant] must have value 8601 or 8701"
MoMisWh8687 = LyzFunMsgNap(CSub, M, "MB52-File Worksheet", FxMB52, Wsn)
End Function

Function GenOMain$(D As Database, IMB52$, IUom$)
'@IMB52 :Drs-Whs-Sku-QUnRes-QBlk-QInsp
'@IUom  :Sku-Sc_U-Des-StkUom
'Ret      : @@
DrpT D, "@Main"

'== Crt @Main fm #IMB52
'   Whs Sku OH Des StkUom Sc_U OH
Rq D, "Select Distinct Whs,Sku,Sum(QUnRes+QBlk+QInsp) As OH into [@Main] from [#IMB52] Group by Whs,Sku"
Rq D, "Alter Table [@Main] Add Column Des Text(255), StkUom Text(10),Sc_U Int, OH_Sc Double"
Rq D, "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.Sc_U = a.Sc_U,x.Des=a.Des,x.StkUom=a.StkUom"
Rq D, "Update [@Main] set OH_Sc=OH/Sc_U where Sc_U>0"

'== Add Col Stream ProdH F2 M32 M35 M37 Topaz ZHT1 RateSc Z2 Z5 Z7
'   Upd Col ProdH Topaz
'   Upd Col F2 M32 M35 M37
Rq D, "Alter Table [@Main] add column Stream Text(10), Topaz Text(20), ProdH text(7), F2 Text(2), M32 text(2), M35 text(5), M37 text(7), ZHT1 Text(7), Z2 text(2), Z5 text(5), Z7 text(7), RateSc Currency, Amt Currency"
Rq D, "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.ProdH=a.ProdH,x.Topaz=a.Topaz"
Rq D, "Update [@Main] set F2=Left(ProdH,2),M32=Mid(ProdH,3,2),M35=Mid(ProdH,3,5),M37=Mid(ProdH,3,7)"

'== Upd Col ZHT1 RateSc
Rq D, "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M37=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
Rq D, "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M35=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
Rq D, "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M32=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"

'Stream
'Z2 Z5 Z7
'Amt
Rq D, "Update [@Main] set Stream=IIf(Left(Topaz,3)='UDV','Diageo','MH')"
Rq D, "Update [@Main] Set Z2=Left(ZHT1,2), Z5=Left(ZHT1,5), Z7=Left(ZHT1,7) where not ZHT1 is null"
Rq D, "Update [@Main] Set Amt = RateSc * OH_Sc where RateSc is not null"
End Function

Function GenORate$(D As Database, IZHT187$, IZHT186$, IUom$)
'VdtFm & VdtTo format DD.MM.YYYY
'1: #IZHT1 VdtFm VdtTo L3 RateSc
'2: #IUom     SKu Sc_U
'O: @Rate  ZHT1 RateSc
DrpTT D, "#Cpy1 #Cpy2 #Cpy @Rate"
Rq D, "Select '8701' as Whs,x.* into [#Cpy1] from [#IZHT18701] x"
Rq D, "Select '8601' as Whs,x.* into [#Cpy2] from [#IZHT18601] x"

Rq D, "Select * into [#Cpy] from [#Cpy1] where False"
Rq D, "Insert into [#Cpy] select * from [#Cpy1]"
Rq D, "Insert into [#Cpy] select * from [#Cpy2]"

Rq D, "Alter Table [#Cpy] Add Column VdtFmDte Date,VdtToDte Date,IsCur YesNo"
Rq D, "Update [#Cpy] Set" & _
" VdtFmDte = DateSerial(RIGHT(VdtFm,4),MID(VdtFm,4,2),LEFT(VdtFm,2))," & _
" VdtToDte = DateSerial(RIGHT(VdtTo,4),MID(VdtTo,4,2),LEFT(VdtTo,2))"
Rq D, "Update [#Cpy] set IsCur = true where Now between VdtFmDte and VdtToDte"

Rq D, "Select Whs,ZHT1,RateSc into [@Rate] from [#Cpy]"
DrpTT D, "#Cpy #Cpy1 #Cpy2"
End Function

Sub GenOupTbl(D As Database)
Dim IZHT187$, IZHT186$, IUom$, IMB52$
GenORate D, IZHT187, IZHT186, IUom
GenOMain D, IMB52, IUom
End Sub
