Attribute VB_Name = "MxDoMthCml"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoMthCml."
#Const Sav = True
':MthCml$ = "NewType:Sy."

Function DoMthCmlP() As Drs
Dim A$()
A = MthNyP
A = AeEle(A, "Z")
A = AeLik(A, "Z_*")
A = AwDist(A)
A = AySrtQ(A)
DoMthCmlP = DoCmlnss(A)
End Function

Function WsoMthCmlP() As Worksheet
Set WsoMthCmlP = WszDrs(DoMthCmlP)
End Function

