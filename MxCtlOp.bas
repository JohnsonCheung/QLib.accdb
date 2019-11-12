Attribute VB_Name = "MxCtlOp"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCtlOp."

Function RmvRmkzVb$(Lin)
Stop
RmvRmkzVb = LeftIf(Lin, RmkPoszVb(Lin))
End Function

Sub Z_RmkPoszVb()
Dim I, O$(), L$, P%
For Each I In AwSubStr(AwSubStr(SrczP(CPj), "'"), """")
    P = RmkPoszVb(I)
    If P = 0 Then
        PushI O, I
    Else
        PushI O, I & vbCrLf & Dup(" ", P - 1) & "^"
    End If
Next
Vc O
End Sub

Function RmkPoszVb%(Lin)
Dim P%: P = InStr(Lin, "'"): If P = 0 Then Exit Function
Dim P1%: P1 = InStr(Left(Lin, P - 1), """"): If P1 > 0 Then Exit Function
RmkPoszVb = P
End Function
Sub TileH()
BoTileH.Execute
End Sub

Sub TileV()
BoTileV.Execute
End Sub
Sub Compile(Pjn$)
JmpPj Pj(Pjn)
BoCompile.Execute
End Sub
Sub CompilezP(P As VBProject)
JmpPj P
ThwIf_BoCompile P.Name
With BoCompile
    If .Enabled Then
        .Execute
        Debug.Print P.Name, "<--- Compiled"
    Else
        Debug.Print P.Name, "already Compiled"
    End If
End With
BoTileV.Execute
BoSav.Execute
End Sub

Sub CompilezV(A As Vbe)
ForItrFun A.VBProjects, "CompilezP"
End Sub

Sub ThwIf_BoCompile(NEPjn$)
Dim Act$, Ept$
Act = BoCompile.Caption
Ept = "Compi&le " & NEPjn
If Act <> Ept Then Thw CSub, "Cur BoCompile.Caption <> Compi&le {Pjn}", "Compile-Btn-Caption Pjn Ept-Btn-Caption", Act, NEPjn, Ept
End Sub

Sub Z_PjCompile()
CompilezP CPj
End Sub

Sub DltClr(A As CommandBar)
Dim I
For Each I In Itr(OyzItr(A.Controls))
    CvCtl(I).Delete
Next
End Sub


Sub DltBar(BarNm$)
If Not HasBar(BarNm) Then Debug.Print "Bar[" & BarNm & "] not found": Exit Sub
Bars(BarNm).Delete
End Sub



Sub EnsBtns(BarBtnccAy$())
Dim I
For Each I In Itr(BarBtnccAy)
    EnsBarBtncc I
Next
End Sub

Sub EnsBarBtncc(BarBtncc)
Dim L$
L = BarBtncc
EnsBtnzCC EnsBar(ShfT1(L)), L
End Sub

Sub RmvBarByNy(BarNy$())
Dim IBar: For Each IBar In BarNy
    If HasBar(IBar) Then
        If Not Bar(IBar).BuiltIn Then
            Bar(IBar).Delete
        End If
    End If
Next
End Sub

Function EnsBar(BarNm$) As CommandBar
If HasBar(BarNm) Then
    Set EnsBar = Bars(BarNm)
Else
    Set EnsBar = Bars.Add(BarNm)
End If
EnsBar.Visible = True
End Function

Sub EnsBtnzCC(Bar As CommandBar, BtnCapcc$)
Dim BtnCap
For Each BtnCap In TermAy(BtnCapcc)
    EnsBtnzC Bar, BtnCap
Next
End Sub

Function HasBtn(Bar As CommandBar, BtnCap) As Boolean
Dim C As CommandBarControl
For Each C In Bar.Controls
    If C.Type = msoControlButton Then
        If C.Caption = BtnCap Then HasBtn = True: Exit Function
    End If
Next
End Function

Sub EnsBtnzC(Bar As CommandBar, BtnCap)
If HasBtn(Bar, BtnCap) Then Exit Sub
Dim B As CommandBarButton
Set B = Bar.Controls.Add(MsoControlType.msoControlButton)
B.Caption = BtnCap
B.Style = msoButtonCaption
End Sub

Sub AddBtn(Bar As CommandBar, BtnCap)
Dim B As CommandBarButton
Set B = Bar.Controls.Add(MsoControlType.msoControlButton)
B.Caption = BtnCap
B.Style = msoButtonCaption
End Sub

Sub ABtn_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
Stop
End Sub

Property Get Y_BtnSpec() As String()
Erase XX
X "Bars"
X " AA A1 A2 A3"
X " BB B1 B2 B3"
X "Btns"
X " A1"
Y_BtnSpec = XX
End Property
