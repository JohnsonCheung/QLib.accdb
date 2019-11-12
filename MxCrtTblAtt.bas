Attribute VB_Name = "MxCrtTblAtt"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxCrtTblAtt."
Sub CrtTblAtt(D As Database)
Const S1$ = "Tbl"
Const S2$ = " AttKey Att FilSi FilTim"
Const S3$ = "Sk"
Const S4$ = " AttKey"
'Attk Text(255), Att Attachment, FilSi Long,FilTim Date  ! The fld spec of create table sql inside the bkt.
CrtSchm D, SyzAp(S1, S2, S3, S4)
End Sub

Sub EnsTblAtt(D As Database)
If HasTbl(D, "Att") Then
    Dim FF$: FF = FFzT(D, "Att")
    If FF <> "Attk Att FilSi FilTim" Then Thw CSub, "Db has :Tbl:Att, but its FF is not [Attk Att FilSi FilTim", "Dbn Tbl-Att-FF", D.Name, FF
End If
CrtTblAtt D
End Sub

Sub Z_EnsTblAtt()
Dim D As Database: Set D = TmpDb
EnsTblAtt D
BrwDb D
End Sub

