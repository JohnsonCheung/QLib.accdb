Attribute VB_Name = "MxTdPrp"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxTdPrp."
Sub SetTdDes(D As Database, T, Des$)
SetTdPrp D.TableDefs(T), C_Des, Des
End Sub

Sub SetTdDeszDic(D As Database, DiTnqDes As Dictionary)
Dim T: For Each T In DiTnqDes.Keys
    SetTdDes D, T, DiTnqDes(T)
Next
End Sub

Function TdDes$(D As Database, T)
TdDes = DaoPv(D.TableDefs(T), C_Des)
End Function

Function DiTnqDes(D As Database) As Dictionary
Dim T, O As New Dictionary
For Each T In Tni(D)
    PushNBlnkzDi O, T, TdDes(D, T)
Next
Set DiTnqDes = O
End Function

Sub SetTdPrp(T As TableDef, Prp$, V)

End Sub

Function DaoPvzP(P As DAO.Properties, Pn$)
If HasDaoPrp(P, Pn) Then DaoPvzP = P(Pn).Value
End Function

Function DaoPrps(DaoPrpsObj) As DAO.Properties
Set DaoPrps = DaoPrpsObj.Properties
End Function

Function DaoPv(DaoPrpObj, P$)
DaoPv = DaoPvzP(DaoPrpObj.Properties, P)
End Function

Function HasDaoPrp(P As DAO.Properties, Pn$) As Boolean
Dim I As DAO.Property: For Each I In P
    If I.Name = Pn Then HasDaoPrp = True: Exit Function
Next
End Function

