Attribute VB_Name = "MxTbMthOp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxTbMthOp."

Sub RfhTbMdD()
RfhTbMd CurrentDb, NewRfhId(CurrentDb), CPj
End Sub

Sub RfhTbMth(D As Database, RfhId&, Pj As VBProject)
'Do : delete reocrds to @D.Mth for those record @RfhId by
'     insert records to @D.Mth from $$Md->Mdl
Dim Ny$(), Ty$(): TwoStrColzTFW "Md", "Mdn MdTy", "UpdId=" & RfhId, Ny, Ty
Dim Pjn$: Pjn = CPjn
Dim N: For Each N In Itr(Ny)
    D.Execute FmtQQ("Delete * from Mth where Mthn='?' and Pjn='?'", N, Pjn)
Next
InsTblzDrs D, "Mth", DoTbMthzN(D, Pjn, Ny)
End Sub

Private Sub InsTbMth(D As Database, Pjn$, MdNy$(), ShtMdTy$())
Dim N, J&: For Each N In Itr(MdNy)
    Dim Mdl$: Mdl = MdlzTbMd(D, Pjn, N)
    Dim DroMdn1(): DroMdn1 = DroMdn(Pjn, ShtMdTy(J), N)
    Dim DoMthc1 As Drs: DoMthc1 = DoMthc(SplitCrLf(Mdl), DroMdn1)
    InsTblzDrs D, "Mth", DoTbMthzN(D, Pjn, MdNy)
    J = J + 1
Next
End Sub

