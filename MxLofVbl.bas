Attribute VB_Name = "MxLofVbl"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxLofVbl."

Function LofVblzQt$(A As Excel.QueryTable)
LofVblzQt = LofVblzFbtStr(FbtStrzQt(A))
End Function

Function LofVblzT$(D As Database, T)
LofVblzT = DbtPrp(D, T, "LofVbl")
End Function

Sub SetLofVblzT(D As Database, T, V$)
DbtPrp(D, T, "LofVbl") = V
End Sub

Function LofVblzLo$(A As ListObject)
LofVblzLo = LofVblzQt(LoQt(A))
End Function

Function LofVblzFbt$(Fb, T)
LofVblzFbt = LofVblzT(Db(Fb), T)
End Function

Sub SetLofVblzFbt(Fb, T, LofVblzVbl$)
SetLofVblzT Db(Fb), T, LofVblzVbl
End Sub

Function LofVblzFbtStr$(FbtStr$)
Dim Fb$, T$
AsgFbtStr FbtStr, Fb, T
LofVblzFbtStr = LofVblzFbt(Fb, T)
End Function
