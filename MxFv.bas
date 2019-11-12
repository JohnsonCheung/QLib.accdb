Attribute VB_Name = "MxFv"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxFv."
Function FvzQ(D As Database, Q)
FvzQ = FvzRs(D.OpenRecordset(Q))
End Function

Function FvzQQ(D As Database, QQSql$, ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
FvzQQ = FvzQ(D, FmtQQAv(QQSql, Av))
End Function

Function FvzRs(A As DAO.Recordset)
If NoRec(A) Then Exit Function
Dim V: V = A.Fields(0).Value
If IsNull(V) Then Exit Function
FvzRs = V
End Function

Function FvzRsF(Rs As DAO.Recordset, Fld)
With Rs
    If .EOF Then Exit Function
    If .BOF Then Exit Function
    FvzRsF = .Fields(Fld).Value
End With
End Function

Function FvzSsk(D As Database, T, F$, Sskv())
Dim Ssk$: Ssk = SskFld(D, T)
FvzSsk = FvzRs(Rs(D, SqlSel_F_T_F_Ev(F, T, Ssk, Sskv)))
End Function

Function FvzTF(D As Database, T, F)
FvzTF = D.TableDefs(T).OpenRecordset.Fields(F).Value
End Function

Function FvzTFW(D As Database, T, F, Bexpr$)
FvzTFW = FvzRs(Rs(D, SqlSel_T_F_Wh(T, F, Bexpr)))
End Function

Sub Z_FvzQ()
Dim D As Database
Ept = CByte(18)
Act = FvzQ(D, "Select Y from [^YM]")
C
End Sub

Function FvzArs(A As ADODB.Recordset)
If NoReczArs(A) Then Exit Function
Dim V: V = A.Fields(0).Value
If IsNull(V) Then Exit Function
FvzArs = V
End Function

Function FvzCnq(A As ADODB.Connection, Q)
FvzCnq = FvzArs(A.Execute(Q))
End Function
