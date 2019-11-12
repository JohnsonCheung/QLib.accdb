Attribute VB_Name = "MxIdeRes"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeRes."

Function ReszMthn(M As CodeModule, Mthn$) As String()
Dim L$(): L = MthLyzM(M, Mthn): If Si(L) Then Exit Function
If Not IsResMth(L) Then Exit Function
ReszMthn = AeFstNEle(AeLasNEle(L, 2), 2)
End Function

Function ReslzMthn$(M As CodeModule, Mthn$)
':MthQn: :Dn|Nm ! if :Dn, Mdn.Mthn, If :Nm Mthn
':Dn:    :Nm.Nm #Dot-Nm#
':DDn:   :Nm{.Nm} #Dot-Dot-Nm#
ReslzMthn = JnCrLf(ReszMthn(M, Mthn))
End Function

Function IsResMth(MthLy$()) As Boolean
Dim N%: N = Si(MthLy)
If N < 4 Then Exit Function
Dim L$: L = MthLy(0)

Select Case True
Case _
    Not ShfPrv(L), _
    Not ShfSub(L), _
    MthLy(1) <> "#If False Then", _
    MthLy(N - 1) <> "End Sub", _
    MthLy(N - 2) <> "#End If"
    Exit Function
End Select
IsResMth = True
End Function
