VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppPm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "AppPm."
Private Type A
    Db As Database
End Type
Private A As A
Const TnPm$ = "Pm"
Const SchmoPm$ = ""
Sub Ini(Db As Database)
Set A.Db = Db
End Sub
Friend Sub EnsTnPm()

End Sub
Property Get PmNy() As String()

End Property
Function PthzPm$(D As Database, PmNm$)
PthzPm = EnsPthSfx(VzPm(D, PmNm & "Pth"))
End Function

Function Pjfnm$(D As Database, PmNm$)
Pjfnm = VzPm(D, PmNm & "Fn")
End Function

Function VzPm(D As Database, PmNm$)
Dim Q$: Q = FmtQQ("Select ? From Pm where CUsr='?'", PmNm, CUsr)
VzPm = FvzQ(D, Q)
End Function

Sub SetVzPm(D As Database, PmNm$, V)
With D.TableDefs("Pm").OpenRecordset
    .Edit
    .Fields(PmNm).Value = V
    .Update
End With
End Sub

Sub BrwPm(D As Database)
BrwT D, "Pm"
End Sub




