Attribute VB_Name = "MxTbsMthOp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxTbsMthOp."
Sub RfhTbsMthD()
RfhTbsMth CurrentDb
End Sub
Sub RfhTbsMth(D As Database)
Dim Pj As VBProject: Set Pj = CPj
Dim RfhId&: RfhId = NewRfhId(D)
RfhTbMd D, RfhId, Pj
RfhTbMth D, RfhId, Pj

'Upd $$Lib from $$Md
DoCmd.RunSql "Delete * from Lib"
DoCmd.RunSql "Insert Into Lib Select Distinct Lib from Md"

'Upd $$Pj from $$Md
DoCmd.RunSql "Delete * from Pj"
DoCmd.RunSql "Insert Into Lib Select Distinct Pj from Md"
End Sub

Function LasId&(D As Database, T)
'@T ! Assume it has a field <T>Id and a "PrimaryKey", using the field as Key
Dim R As DAO.Recordset
Set R = D.TableDefs(T).OpenRecordset
R.Index = "PrimaryKey"
R.MoveLast
LasId = R.Fields(0).Value
End Function

Function LasRfhId&(D As Database)
LasRfhId = LasId(D, "RfhHis")
End Function

Function NewRfhId&(D As Database)
With D.TableDefs("RfhHis").OpenRecordset
    .AddNew
    NewRfhId = !RfhId
    .Update
    .Close
End With
End Function

