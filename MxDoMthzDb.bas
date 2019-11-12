Attribute VB_Name = "MxDoMthzDb"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoMthzDb."

Function DoFun() As Drs
Stop
End Function

Function DoPubSub() As Drs
DoPubSub = DwEq(DoFun, "Ty", "Sub")
End Function

Function DoPubFun() As Drs
DoPubFun = DwEqExl(DoFun, "Ty", "Fun")
End Function

Function DoPubFunwPatn(PatnSS3$) As Drs
DoPubFunwPatn = DwPatnSS3(DoPubFun, "Mthn", PatnSS3)
End Function

Function DoPubPrp() As Drs
DoPubPrp = DwIn(DoPubFun, "Ty", SyzSS("Get Let Set"))
End Function

Function DoPubPrpWiPm() As Drs
Dim A As Drs: A = AddMthColMthPm(DoPubPrpWiPm)
DoPubPrpWiPm = DwEqExl(A, "MthPm", "")
End Function

Function IsPurePrp(MdTy$, MthPm$) As Boolean
End Function

