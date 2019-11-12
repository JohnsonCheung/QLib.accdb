Attribute VB_Name = "MxSelPth"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxSelPth."

Function SelFfn$(Optional Ffn$, Optional FSpec$ = "*.*", Optional Tit$ = "Select a file", Optional BtnNm$ = "Use the File Name")
With Application.FileDialog(msoFileDialogFilePicker)
    .Filters.Clear
    .Title = Tit
    .AllowMultiSelect = False
    .Filters.Add "", FSpec
    .InitialFileName = Ffn
    .ButtonName = BtnNm
    .Show
    If .SelectedItems.Count = 1 Then
        SelFfn = .SelectedItems(1)
    End If
End With
End Function

Sub SetTxtbSelPth(A As Access.TextBox)
Dim R$
R = SelPth(A.Value)
If R = "" Then Exit Sub
A.Value = R
End Sub

Function SelPth$(Optional Pth$, Optional Tit$ = "Select a Path", Optional BtnNm$ = "Use this path")
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .InitialFileName = IIf(IsNull(Pth), "", Pth)
    .Show
    If .SelectedItems.Count = 1 Then
        SelPth = EnsPthSfx(.SelectedItems(1))
    End If
End With
End Function

Sub Z_SelPth()
GoTo Z
Z:
MsgBox SelFfn("C:\")
End Sub

