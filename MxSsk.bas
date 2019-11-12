Attribute VB_Name = "MxSsk"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxSsk."
Public Const Skn$ = "SecondaryKey"
Public Const Pkn$ = "PrimaryKey"

Function AsetzDbtf(D As Database, T, F$) As Aset
Set AsetzDbtf = AsetzRs(Rs(D, SqlSel_F_T(F, T)), F)
End Function

Sub DltRecNotInSskv(D As Database, SskTbl$, NotInSSskv As Aset) _
'Delete Db-T record for those record's Sk not in NotInSSskv, _
'Assume T has single-fld-sk
Const CSub$ = CMod & "DltRecNotInSskv"
'If NotInSSskv.IsEmp Then Thw CSub, "Given NotInSSskv cannot be empty", "Db SskTbl SskFld", Dbn(A), SskTbl, SskFld_Dbt(Db, SskTbl)
Dim Q$, Excess As Aset
Set Excess = SskVset(D, SskTbl).Minus(NotInSSskv)
If Excess.IsEmp Then Exit Sub
'RunSqy Db, SqyDlt_Fm_WhFld_InAset(SskTbl, SskFld_Dbt(Db, SskTbl), Excess)
End Sub

Sub InsReczSskv(D As Database, T, ToInsSSskv As Aset) _
'Insert Single-Field-Secondary-Key-Aset-A into Dbt
'Assume T has single-fld-sk and can be inserted by just giving such SSk-value
Dim ShouldIns As Aset
    Set ShouldIns = ToInsSSskv.Minus(SskVset(D, T))
If ShouldIns.IsEmp Then Exit Sub
Dim I, F$
'F = SskFld_Dbt(Fv, T)
'With RsDbt(Fv, T)
'    For Each I In ShouldIns.Itms
'        .AddNew
'        .Fields(F).Value = I
'        .Update
'    Next
'    .Close
'End With
End Sub

Function SkFny(D As Database, T) As String()
SkFny = SkFnyzTd(D.TableDefs(T))
End Function

Function SkFnyzTd(T As DAO.TableDef) As String()
SkFnyzTd = FnyzIdx(SkIdxzTd(T))
End Function

Function SkIdx(D As Database, T) As DAO.Index
Set SkIdx = Idx(D, T, Skn)
End Function

Function SkIdxzTd(T As DAO.TableDef) As DAO.Index
Set SkIdxzTd = IdxzTd(T, Skn)
End Function

Function SkSqlQPfxSy(D As Database, T) As String()
Stop '
End Function

Function SskFld$(D As Database, T)
Dim Sk$(): Sk = SkFny(D, T): If Si(Sk) = 1 Then SskFld = Sk(0): Exit Function
Thw CSub, "SkFny-Sz<>1", "Db T, SkFny-Si SkFny", D.Name, T, Si(Sk), Sk
End Function

Function Sskv(D As Database, T) As Aset
'SSskv is [S]ingleFielded [S]econdKey [K]ey [V]alue [Aset], which is always a Value-Aset.
'and Ssk is a field-name from , which assume there is a Unique-Index with name "SecordaryKey" which is unique and and have only one field
'Set Sskv = ColSet(SskFld)
End Function

Function SskVset(D As Database, T) As Aset
Set SskVset = AsetzDbtf(D, T, SskFld(D, T))
End Function

