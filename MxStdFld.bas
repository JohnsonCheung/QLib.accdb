Attribute VB_Name = "MxStdFld"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxStdFld."
Public Const SSoStdEle$ = "CrtDte Pk Fk Ty Nm Dte Amt Att"
Function FdzStdFld(F, Optional T) As DAO.Field2
Dim E$: E = ElezStdFld(F, T)
Set FdzStdFld = FdzStdEle(F, E)
End Function

Function IsStdFld(F) As Boolean
IsStdFld = ElezStdFld(F) <> ""
End Function

Function FdzStdEle(F, E) As DAO.Field2
Set FdzStdEle = F & " " & StdEleStr(E)
End Function

Function StdEleStr$(E)
Dim O$
Select Case E
Case "CrtDte"
Case "Dte"
Case "Pk"
Case "Fk"
Case "Ty"
Case "Nm"
Case "Dte"
Case "Amt"
Case "Att"
Case Else: Thw CSub, "Given Ele is not std", "E", E
End Select
StdEleStr = O
End Function

Function ElezStdFld$(F, Optional T)
Dim R2$, R3$
R2 = Right(F, 2)
R3 = Right(F, 3)
Dim O$
Select Case True
Case F = "CrtDte":  O = "CrtDte"
Case T & "Id" = F:  O = "Pk"
Case R2 = "Id":     O = "Fk"
Case R2 = "Ty":     O = "Ty"
Case R2 = "Nm":     O = "Nm"
Case R3 = "Dte":    O = "Dte"
Case R3 = "Amt":    O = "Amt"
Case R3 = "Att":    O = "Att"
End Select
ElezStdFld = O
End Function

Function StdEleAy() As String()
StdEleAy = SyzDicKey(DiStdEleqEleStr)
End Function

Function DiStdEleqEleStr() As Dictionary
Static X As Boolean, Y As Dictionary
If Not X Then
    X = True
    Set Y = New Dictionary
    Y.Add "Id", ""
    Y.Add "*Id", ""
    Y.Add "*Id", ""
    Y.Add "*Id", ""
    Y.Add "*Id", ""
    Y.Add "*Id", ""
    Y.Add "*Id", ""
End If
Set DiStdEleqEleStr = Y
End Function
