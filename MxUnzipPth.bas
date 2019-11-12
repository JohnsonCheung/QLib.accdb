Attribute VB_Name = "MxUnzipPth"
Option Explicit
Option Compare Text
Const CLib$ = "QZip."
Const CMod$ = CLib & "MxUnzipPth."

Sub UnzipPth(Z7Db As Database, FmZipFfn, ToPth)
If IsPmEr(FmZipFfn, ToPth) Then Exit Sub
If IsCfmAndClrPthR(ToPth) Then Exit Sub

Dim P$:             P = AddFdrEns(Pth(Z7Db.Name), "ZipPthWrking")
Dim Fpgm$:       Fpgm = P & "z7.exe"
Dim Fcmd$:       Fcmd = P & "Unzip.Cmd"
Dim FcmdStr$: FcmdStr = Fnd_FcmdStr(P, FmZipFfn, ToPth)
Dim ShellStr$: ShellStr = FmtQQ("Cmd.Exe /C ""?""", Fcmd)

                        Expz7 Z7Db, Fpgm
                        WrtStr FcmdStr, Fcmd
                        Shell Fcmd, vbMaximizedFocus
End Sub


Private Function Fnd_FcmdStr$(WrkPth, FmZipFfn, ToPth)
Dim O$()
Dim W$: W = EnsPthSfx(WrkPth)
Dim Z$: Z = W & Fn(FmZipFfn)
Dim T$: T = ParPth(ToPth)       ' Target-Path should the parent of @ToPth
Push O, FmtQQ("Cd ""?""", T)
Push O, ""
Push O, FmtQQ("Copy  ""?"" ""?""", FmZipFfn, W)
Push O, ""
Push O, FmtQQ("""?z7.exe"" x -r ""?""", W, Z)
Push O, ""
Push O, FmtQQ("Del ""?""", Z)
Push O, "Pause"
Fnd_FcmdStr = JnCrLf(O)
End Function
'--
'2 Vdt
Private Function IsPmEr(FmZipFfn, ToPth) As Boolean
IsPmEr = True
If NoPth(ParPth(ToPth)) Then SetMainMsg "To Path not found: " & ParPth(ToPth): Exit Function
If NoFfn(FmZipFfn) Then SetMainMsg "Zip file not found: " & FmZipFfn: Exit Function
IsPmEr = False
End Function

Sub Z_UnzipPth()
Dim B$: B = ResPthB
Dim A$: A = ResPthA
CrtResPthA
EnsPth B
Dim Z$: Z = ZipPth(CurrentDb, A, B)
Stop ' Wait the Dos to zip
UnzipPth CurrentDb, Z, A
BrwPth A '<== A should be restored
End Sub
