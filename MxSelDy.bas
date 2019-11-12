Attribute VB_Name = "MxSelDy"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxSelDy."
Function DywKeyDr(FmDy(), KeyDr(), KeyDy()) As Variant()
Dim SubIxy&():    SubIxy = F_SubIxy_ByDr_InDy(KeyDr, KeyDy)
         DywKeyDr = AwIxy(FmDy, SubIxy)
End Function

Function F_SubIxy_ByDr_InDy(ByDr(), InDy()) As Long()
Dim Dr, I&: For Each Dr In Itr(InDy)
    If IsEqDr(Dr, ByDr) Then
        #If False Then
        Stop
        Debug.Print I
        Debug.Print JnSpc(Dr)
        Debug.Print JnSpc(ByDr)
        Debug.Print
        #End If
        PushI F_SubIxy_ByDr_InDy, I
    End If
    I = I + 1
Next
End Function

Sub Z_F_SubIxy_ByDr_InDy()
Dim KeyDr()
Dim KeyDy()
    KeyDy = SelDrs(DoMdP, "Clibv CNsv").Dy
Dim SubIxy1&()
    KeyDr = Array("QGit", Empty)
    SubIxy1 = F_SubIxy_ByDr_InDy(KeyDr, KeyDy)
Dim SubIxy2&()
    KeyDr = Array("QAct", Empty)
    SubIxy2 = F_SubIxy_ByDr_InDy(KeyDr, KeyDy)
Dmp SubIxy1
Debug.Print
Dmp SubIxy2

11
82
150
176
181
327
351
352
QGit:
QAcs:
End Sub
