Attribute VB_Name = "MxZipPth"
Option Compare Text
Option Explicit
Const CLib$ = "QApp."
Const CMod$ = CLib & "MxZipPth."
Function ZipPth$(Z7Db As Database, FmPth, ToPth)
If IsPmEr(FmPth, ToPth) Then Exit Function
Dim P$:             P = AddFdrEns(Pth(Z7Db.Name), "ZipPthWrking")
Dim Fpgm$:       Fpgm = P & "z7.exe"
Dim Fcmd$:       Fcmd = P & "Zip.Cmd"
Dim Foup$:       Foup = Fdr(FmPth) & "(" & Format(Now, "YYYY-MM-DD HH-MM") & " " & CUsr & ").zip"
Dim FcmdStr$: FcmdStr = Fnd_FcmdStr(P, FmPth, ToPth, Foup)
Dim ShellStr$: ShellStr = FmtQQ("Cmd.Exe /C ""?""", Fcmd)
                        Expz7 Z7Db, Fpgm
                        WrtStr FcmdStr, Fcmd
                        Shell Fcmd, vbMaximizedFocus
               ZipPth = EnsPthSfx(ToPth) & Foup
End Function

Sub Expz7(Z7Db As Database, Fpgm$)
'Private Sub Exp7z(): Dim Z7Db As Database, Fpgm$: Set Z7Db = CurrentDb: Fpgm = "C:\users\user\documents\projects\vba\backuppth\7z.exe"
If HasFfn(Fpgm) Then Exit Sub
Dim R As DAO.Recordset: Set R = Z7Db.TableDefs("7z").OpenRecordset
Dim R2 As DAO.Recordset2: Set R2 = R.Fields("7z").Value
Dim F2 As DAO.Field2: Set F2 = R2.Fields("FileData")
Dim NoExt$: NoExt = RmvExt(Fpgm)
DltFfnIf NoExt
F2.SaveToFile NoExt
Name NoExt As Fpgm
End Sub

Private Function Fnd_FcmdStr$(WrkPth, FmPth, ToPth, Foup$)
Dim O$()
Push O, FmtQQ("Cd ""?""", WrkPth)
Push O, FmtQQ("z7 a ""?"" ""?""", Foup, FmPth)
Push O, FmtQQ("Copy ""?"" ""?""", Foup, ToPth)
Push O, FmtQQ("Del ""?""", Foup)
Push O, "Pause"
Fnd_FcmdStr = JnCrLf(O)
End Function

Private Function IsPmEr(FmPth, ToPth) As Boolean
IsPmEr = True
If NoPth(ToPth) Then SetMainMsg "To Path not found: " & ToPth: Exit Function
If NoPth(FmPth) Then SetMainMsg "From Path not found: " & FmPth: Exit Function
IsPmEr = False
End Function

Sub Z_ZipPth()
CrtResPthA
ZipPth CurrentDb, ResPthA, ResPth("B")
End Sub
