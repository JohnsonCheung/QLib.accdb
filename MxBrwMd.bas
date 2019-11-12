Attribute VB_Name = "MxBrwMd"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxBrwMd."

Sub BrwMd(Optional MdPatnSS3$, Optional NsPatnSS3$, Optional Lib$, Optional SrtFF$)
LisMd MdPatnSS3, NsPatnSS3, Lib, SrtFF, OupTy:=EiOtBrw, Top:=0
End Sub

Sub LisMd(Optional MdPatnSS3$, Optional NsPatnSS3$, Optional Lib$, Optional SrtFF$, Optional OupTy As EmOupTy = EmOupTy.EiOtDmp, Optional Top% = 50)
Dim D1 As Drs: D1 = DwPatnSS3(DoMdP, "Mdn", MdPatnSS3)
Dim D2 As Drs: D2 = DwPatnSS3(D1, "CNsv", NsPatnSS3)
                    If Lib <> "" Then D2 = DwEQ(D2, "CLibv", Lib)
Dim D3 As Drs: D3 = SrtDrs(D2, SrtFF)
Dim Ly$():     Ly = FmtDrs(D3, Fmt:=EiSSFmt)
                    Brw Ly, OupTy:=OupTy
End Sub

Sub VcMd(Optional MdPatnSS3$, Optional NsPatnSS3$, Optional Lib$, Optional SrtFF$)
LisMd MdPatnSS3, NsPatnSS3, Lib, SrtFF, OupTy:=EiOtVc, Top:=0
End Sub
