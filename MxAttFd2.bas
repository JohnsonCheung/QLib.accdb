Attribute VB_Name = "MxAttFd2"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxAttFd2."
Function AttFldR2(R As DAO.Recordset, AttFld$) As DAO.Recordset2
Set AttFldR2 = R.Fields(AttFld).Value
End Function

Function Fd2zRsAttFld(R As DAO.Recordset, AttFld$) As DAO.Field2
Set Fd2zRsAttFld = AttFldR2(R, AttFld).Fields("FileData")
Stop
End Function
Function FileDataF2(R As DAO.Recordset, AttFld) As DAO.Field2

End Function
