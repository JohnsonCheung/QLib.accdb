Attribute VB_Name = "MxEleStr"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxEleStr."
Function ErzEleStr$(EleStr$)

End Function
Function FdzEleStr(F, EleStr$) As DAO.Field2
Stop '
End Function
Function FdzE(F, StdEle$) As DAO.Field2
Dim O As DAO.Field2
Set O = FdzTnnn(F, StdEle): If Not IsNothing(O) Then Set FdzE = O: Exit Function
Select Case StdEle
Case "Nm":  Set FdzE = FdzNm(F)
Case "Amt": Set FdzE = FdzCur(F): FdzE.DefaultValue = 0
Case "Txt": Set FdzE = FdzTxt(F, dbText, True): FdzE.DefaultValue = """""": FdzE.AllowZeroLength = True
Case "Dte": Set FdzE = FdzDte(F)
Case "Int": Set FdzE = FdzInt(F)
Case "Lng": Set FdzE = FdzLng(F)
Case "Dbl": Set FdzE = FdzDbl(F)
Case "Sng": Set FdzE = FdzSng(F)
Case "Lgc": Set FdzE = FdzBool(F)
Case "Mem": Set FdzE = FdzMem(F)
End Select
End Function
