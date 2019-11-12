Attribute VB_Name = "MxDoDim"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoDim."
Public Const FFoDim$ = "DimItm V Vsf"

Sub Z_DoDimP()
GoSub Z1
Exit Sub
Z1:
    Brw AySrtQ(AwDist(StrCol(DoDimP, "Vsf")))
    Return
Z:  BrwDrs DoDimP
    Return
End Sub

Function DoDimP() As Drs
DoDimP = DoDimzP(CPj)
End Function
Function DoDimzP(P As VBProject) As Drs
DoDimzP = DoDim(DimItmAyzS(SrczP(P)))
End Function
Function DoDim(DimItmAy$()) As Drs
DoDim = Drs(FoDim, DyoDim(DimItmAy))
End Function

Function FoDim() As String()
FoDim = SyzSS(FFoDim)
End Function

Function DyoDim(DimItmAy$()) As Variant()
Dim I: For Each I In Itr(DimItmAy)
    PushI DyoDim, DroDim(I)
Next
End Function

Function DroDim(DimItm) As Variant()
With S12oDimnqVsfx(DimItm)
DroDim = Array(DimItm, .S1, .S2)
End With
End Function
