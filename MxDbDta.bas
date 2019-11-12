Attribute VB_Name = "MxDbDta"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDbDta."

Function LngAyzQ(D As Database, Q) As Long()
LngAyzQ = LngAyzRs(Rs(D, Q))
End Function

Function SyzQ(D As Database, Q) As String()
SyzQ = SyzRs(Rs(D, Q))
End Function

Sub Z_Rs()
Shell "Subst N: c:\subst\users\user\desktop", vbHide
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
BrwAy CsvLyzRs(Rs(TmpDb, S))
End Sub

Function StrColzTF(D As Database, T, F) As String()
StrColzTF = StrColzRs(RszT(D, T), F)
End Function

Function StrColzQ(D As Database, Q) As String()
StrColzQ = StrColzRs(Rs(D, Q))
End Function

Function StrColzTFW(D As Database, T, F, Bexpr$) As String()
Dim Q$: Q = FmtQQ("Select [?] from [?] where ?", F, T, Bexpr)
StrColzTFW = StrColzRs(Rs(D, Q), 0)
End Function

Sub TwoStrColzTFW(T, F12$, Bexpr$, OStrCol1$(), OStrCol2$())
':F12: :SS #Fld1-Fld2-Separated-Spc#
Dim F1$, F2$: AsgS12 BrkSpc(F12), F1, F2
Dim Q$: Q = FmtQQ("Select [?],[?] from [?] where ?", F1, F2, T, Bexpr)
TwoStrColzRs Rs(CurrentDb, Q), OStrCol1, OStrCol2
End Sub

Sub TwoStrColzTF(T, F12$, OStrCol1$(), OStrCol2$())
':F12: :SS #Fld1-Fld2-Separated-Spc#
Dim F1$, F2$: AsgS12 BrkSpc(F12), F1, F2
Dim Q$: Q = FmtQQ("Select [?],[?] from [?]", F1, F2, T)
TwoStrColzRs Rs(CurrentDb, Q), OStrCol1, OStrCol2
End Sub

Sub TwoStrColzRs(R As DAO.Recordset, OStrCol1$(), OStrCol2$())
Erase OStrCol1, OStrCol2
With R
    If Not .EOF Then .MoveFirst
    While Not .EOF
        PushI OStrCol1, Nz(.Fields(0).Value, "")
        PushI OStrCol2, Nz(.Fields(1).Value, "")
        .MoveNext
    Wend
End With
End Sub


Function StrColzRs(R As DAO.Recordset, Optional F = 0) As String()
With R
    If Not .EOF Then .MoveFirst
    While Not .EOF
        PushI StrColzRs, .Fields(F).Value
        .MoveNext
    Wend
End With
End Function

Sub BrwQ(D As Database, Q)
BrwDrs DrszQ(D, Q)
End Sub

Function IntAyzQ(D As Database, Q) As Integer()
End Function

Function SyzTF(D As Database, T, F$) As String()
SyzTF = SyzRs(RszTF(D, T, F))
End Function

Function IntozTF(Into, D As Database, T, F$)
IntozTF = IntozRs(Into, RszTF(D, T, F))
End Function



Function FnyzQ(D As Database, Q) As String()
FnyzQ = FnyzRs(Rs(D, Q))
End Function

Sub Z_FnyzQ()
Dim Db As Database
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
DmpAy FnyzQ(Db, S)
End Sub






Function DrzQ(D As Database, Q) As Variant()
DrzQ = DrzRs(Rs(D, Q))
End Function

Function DyoQ(D As Database, Q) As Variant()
DyoQ = DyoRs(Rs(D, Q))
End Function
