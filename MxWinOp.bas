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

Sub JmpCmpn(Cmpn$)
Dim C As VBIDE.CodePane: Set C = PnezCmpn(Cmpn)
If IsNothing(C) Then Debug.Print "No such WinOfCmpNm": Exit Sub
C.Show
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

