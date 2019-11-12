Attribute VB_Name = "MxQuote"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxQuote."
Function BrkQte(QteStr$) As S12
Dim L%: L = Len(QteStr)
Dim S1$, S2$
Select Case L
Case 0:
Case 1
    S1 = QteStr
    S2 = QteStr
Case 2
    S1 = Left(QteStr, 1)
    S2 = Right(QteStr, 1)
Case Else
    If InStr(QteStr, "*") > 0 Then
        BrkQte = Brk(QteStr, "*", NoTrim:=True)
        Exit Function
    End If
    Stop
End Select
BrkQte = S12(S1, S2)
End Function
Sub AsgQte(OQ1$, OQ2$, QteStr$)
With BrkQte(QteStr)
    OQ1 = .S1
    OQ2 = .S2
End With
End Sub
Function QteBigBkt$(S)
QteBigBkt = "{" & S & "}"
End Function

Function QteBkt$(S)
QteBkt = "(" & S & ")"
End Function
Function QteDot$(S)
QteDot = "." & S & "."
End Function
Function QteAy(Ay, QteStr$) As String()
Dim P$, S$
With BrkQte(QteStr)
    P = .S1
    S = .S2
End With
QteAy = AmAddPfxS(Ay, P, S)
End Function

Function UnQteVb$(QteVb)
UnQteVb = Replace(RmvFstLasChr(QteVb), vb2DblQ, vbDblQ)
End Function

Function QteVb$(S)
':QteVb: #Quoted-Vb-Str# ! a str with fst and lst chr is vbDblQ and inside each vbDblQ is in pair, which will cv to one vbDblQ  @@
QteVb = vbDblQ & Replace(S, vbDblQ, vb2DblQ) & vbDblQ
End Function

Function Qte$(S, QteStr$)
With BrkQte(QteStr)
    Qte = .S1 & S & .S2
End With
End Function

Function QteDbl$(S)
QteDbl = vbDblQ & S & vbDblQ
End Function

Function QteSng$(S)
QteSng = "'" & S & "'"
End Function

Function QteSq$(S)
QteSq = "[" & S & "]"
End Function
Function QteSqIf$(S)
If IsNeedQte(S) Then QteSqIf = QteSq(S) Else QteSqIf = S
End Function
Function QteSqAv(Av()) As String()
Dim I
For Each I In Av
    PushI QteSqAv, QteSq(I)
Next
End Function
