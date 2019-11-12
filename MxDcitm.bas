Attribute VB_Name = "MxDcitm"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDcitm."
':ShtDclSfx: :Sfx
':DclSfx:    :Sfx
':Dcn:   :Nm #Dcl-Itm-Nm#
':Dcitm: :S  #Dcl-Itm# ! Is from [Dim] | [Arg]-aft-rmv-[=]-Optional-Paramarray | End-Type
Function Dcn$(Dcitm)
If HasSubStr(Dcitm, " As ") Then
    Dcn = DcnoAs(Dcitm)
Else
    Dcn = DcnoTyChr(Dcitm)
End If
End Function

Function DcnoTyChr$(DimShtItm)
DcnoTyChr = RmvLasChrzzLis(RmvSfxzBkt(DimShtItm), MthTyChrLis)
End Function

Function DcnoAs$(DimAsItm)
DcnoAs = RmvSfxzBkt(Bef(DimAsItm, " As"))
End Function

Function DclSfx$(Dcitm$)
DclSfx = ShtDclSfx(RmvNm(Dcitm))
End Function

Function ShtDclSfx$(DclSfx$)
If DclSfx = "" Then Exit Function
Dim L$: L = DclSfx
Select Case True
Case L = " As Boolean":: ShtDclSfx = "^"
Case L = " As Boolean()": ShtDclSfx = "^()"
Case Else
    ShfPfx L, " As "
    ShtDclSfx = L
End Select
End Function
