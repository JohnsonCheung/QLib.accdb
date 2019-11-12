Attribute VB_Name = "MxSetFv"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxSetFv."

Sub SetFvzQ(D As Database, Q, V)
FvzRs(D.OpenRecordset(Q)) = V
End Sub

Sub SetFvzRs(A As DAO.Recordset, V)
If NoRec(A) Then
    A.AddNew
Else
    A.Edit
End If
A.Fields(0).Value = V
A.Update
End Sub

Sub SetFvzRsF(Rs As DAO.Recordset, Fld, V)
With Rs
    .Edit
    .Fields(Fld).Value = V
    .Update
End With
End Sub

Sub SetFvzSsk(D As Database, T, F$, Sskv(), V)
FvzRs(Rs(D, SqlSel_F_T_F_Ev(F, T, SskFld(D, T), Sskv))) = V
End Sub
