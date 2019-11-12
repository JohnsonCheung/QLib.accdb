Attribute VB_Name = "Mx3Cnstv"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "Mx3Cnstv."

Function CNsLin$(Ns$)
':CLibLin: :PrvCnstLin ! Is a `Const CLib$ = "${Clibv}."`
CNsLin = FmtQQ("Const CNs$ = ""?""", Ns)
End Function

Function CLibLin$(CLibv$)
':CLibLin: :PrvCnstLin ! Is a `Const CLib$ = "${Clibv}."`
CLibLin = FmtQQ("Const CLib$ = ""?.""", CLibv)
End Function

Function CModLin$(M As CodeModule)
':CModLin: :CnstLin ! Is a Const CMod$ = CLib & "xxxx."
CModLin = FmtQQ("Const CMod$ = CLib & ""?.""", Mdn(M))
End Function

Function CNsvzM$(M As CodeModule)
CNsvzM = CNsv(DclzM(M))
End Function

Function CNsvM$()
CNsvM = CNsvzM(CMd)
End Function

Function CNsv$(Dcl$())
CNsv = CnstStrvzDcl(Dcl, "CNs")
End Function


Function CModv$(Dcl$())
CModv = RmvSfxDot(CnstStrvzDcl(Dcl, "CMod"))
End Function

Function CLibv$(Dcl$())
CLibv = RmvSfxDot(CnstStrvzDcl(Dcl, "CLib"))
End Function

