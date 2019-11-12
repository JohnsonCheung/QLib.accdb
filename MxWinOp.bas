Attribute VB_Name = "MxWinOp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxWinOp."
Sub ClsAllWin()
Dim W As VBIDE.Window: For Each W In CVbe.Windows
    If W.Visible Then W.Close
Next
End Sub
Sub ClsWin(W As VBIDE.Window)
W.Visible = False
End Sub

Sub ShwWin(W As VBIDE.Window)
W.Visible = True
End Sub


Sub ClsWinExlAp(ParamArray ExlWinAp())
Dim I, W As VBIDE.Window, Av(): Av = ExlWinAp
For Each I In Itr(VisWiny)
    Set W = I
    If Not HasObj(Av, W) Then
        ClsWin W
    Else
        ShwWin W
    End If
Next
End Sub

Sub ShwDbg()
ClsWinExlAp ImmWin, LclWin, CWin
DoEvents
TileV
End Sub

Sub ClrImm()
Dim W As VBIDE.Window
DoEvents
With ImmWin
    .SetFocus
    .Visible = True
End With
SndKeys "^{HOME}^+{END}"
DoEvents
'SndKeys "{DEL}" '<-- it does not work?
'DoEvents
End Sub

Sub ClsWinE(Optional Mdn$)
' ! Cls win ept cur md @@
Dim W1 As VBIDE.Window: Set W1 = CWin
Dim W2 As VBIDE.Window: Set W2 = WinzMdn(Mdn)
Dim W As VBIDE.Window: For Each W In CVbe.Windows
    If Not IsEqObj(W1, W) Then
        If Not IsEqObj(W2, W) Then
            If W.Visible Then W.Close
        End If
    End If
Next
ImmWin.Close
BoTileV.Execute
End Sub

Sub ClrWin(A As VBIDE.Window)
DoEvents
BoSelAll.Execute
DoEvents
SendKeys " "
BoEdtClr.Execute
End Sub

