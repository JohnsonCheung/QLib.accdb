Attribute VB_Name = "MxCache"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCache."

Function CacheDtezPjf(Pjf) As Date
CacheDtezPjf = FvzQ(MthDbP, FmtQQ("Select PjDte from Mth where Pjf='?'", Pjf))
End Function

Function DoMthzPjfzFmCache(Pjf) As Drs
Dim Sql$: Sql = FmtQQ("Select * from MthCache where Pjf='?'", Pjf)
DoMthzPjfzFmCache = DrszFbq(MthDbP, Sql)
End Function

Function SkFnyWiSqlQPfx(D As Database, T) As String()
Dim F
For Each F In Itr(SkFny(D, T))
    PushI SkFnyWiSqlQPfx, SqlQteChrzT(DaoTyzTF(D, T, F)) & F
Next
End Function

Sub IupTbl(D As Database, T, Drs As Drs)
Dim Dy(): Dy = Drs.Dy
If Si(Dy) = 0 Then Exit Sub
Dim R As DAO.Recordset, Q$, Sql$, Dr
'Sql = SqlSel_T_Wh(T, BexpzFnyzSqlQPfxSy(SkFny(D, T), SkSqlQPfxSy(D, T)))
For Each Dr In Dy
    Q = FmtQQAv(Sql, CvAv(Dr))
    Set R = Rs(D, Q)
    If HasRec(R) Then
        UpdRs R, Dr
    Else
        InsRs R, Dr
    End If
Next
End Sub
