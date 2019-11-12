Attribute VB_Name = "MxCntsi"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxCntsi."
Function CntsiStrzLines$(Lines$)
CntsiStrzLines = CntsiStr(LinCnt(Lines), Len(Lines))
End Function

Function CntsiStr$(Cnt&, Si&)
CntsiStr = FmtQQ("Cnt-Size(? ?)", Cnt, Si)
End Function
Function CntsiStrzLy$(Ly$())
CntsiStrzLy = CntsiStr(Si(Ly), Len(JnCrLf(Ly)))
End Function
