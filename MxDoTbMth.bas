Attribute VB_Name = "MxDoTbMth"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoTbMth."

Private Sub Z_TbMthP()
BrwDrs TbMthD
End Sub

Function TbMth(D As Database) As Drs
TbMth = DrszT(D, "Mth")
End Function

Function TbMthD() As Drs
TbMthD = TbMth(CurrentDb)
End Function

Function DoTbMthzN(D As Database, Pjn$, MdNy$()) As Drs
':$$Mth:  ! Pjn MdTy Mdn Mdy Ty Mthn L E MthLin TyChr RetAs ShtPm Mthl Rmk @@
'FFoMthc  = Pjn MdTy Mdn Mdy Ty Mthn L E MthLin Mthl
DoTbMthzN = Add5MthCol(DoMthczDbTbMd(D, Pjn$, MdNy))
End Function

Function DoMthczDbTbMd(D As Database, Pjn$, MdNy$()) As Drs
'Ret : :DoMthc ! Fm @D->TbMd @@
Dim O As Drs
Dim J%, N: For Each N In Itr(MdNy)
    Dim R As DAO.Recordset: Set R = Rs(D, FmtQQ("Select MdTy,Mdl from Md where Pjn='?' and Mdn='?'", Pjn, N))
    Dim Src$(): Src = SplitCrLf(CStr(R!Mdl))
    Dim T$: T = R!MdTy
    Dim Dr(): Dr = DroMdn(Pjn, T, N)
    O = AddDrs(O, DoMthc(Src, Dr))
    J = J + 1
Next
DoMthczDbTbMd = O
End Function

Function MdlzTbMdP$(Mdn)
MdlzTbMdP = MdlzTbMd(CurrentDb, CPjn, Mdn)
End Function

Function MdlzTbMd$(D As Database, Pjn$, Mdn)
Dim B$: B = FmtQQ("Pjn='?' and Mdn='?'", Pjn, Mdn)
MdlzTbMd = FvzTFW(D, "Md", "Mdl", B)
End Function

Sub Z_RfhTbMth()
RfhTbMth CurrentDb, LasRfhId(CurrentDb), CPj
End Sub
