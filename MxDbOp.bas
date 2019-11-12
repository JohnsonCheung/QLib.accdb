Attribute VB_Name = "MxDbOp"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDbOp."
Sub RunSqy(D As Database, Sqy$())
Dim Q$, I
For Each I In Sqy
    Q = I
    Rq D, Q
Next
End Sub
Sub AppTdAy(D As Database, TdAy() As DAO.TableDef)
Dim T
For Each T In Itr(TdAy)
    D.TableDefs.Append T
Next
End Sub

