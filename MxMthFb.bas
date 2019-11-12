Attribute VB_Name = "MxMthFb"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMthFb."
Function MthFbP$()
MthFbP = MthFbzP(CPj)
End Function

Function MthFb$()
MthFb = MthFbP
End Function

Function MthFbzP$(P As VBProject)
MthFbzP = MthPthzP(P) & Fn(Pjf(P)) & ".MthDb.accdb"
End Function

Function MthPthzP$(P As VBProject)
Dim F$: F = Pjf(P)
MthPthzP = AddFdrEns(Pth(F), ".MthDb")
End Function

Function EnsMthFb(MthFb$) As Database
EnsFb MthFb
Dim D As Database
Set EnsMthFb = Db(MthFb)
'EnsSchm D, LnoChm
End Function

Function MthDbzP(P As VBProject) As Database
Dim Fb$: Fb = MthFbzP(P)
EnsMthFb Fb
Set MthDbzP = Db(Fb)
End Function

Property Get MthDbP() As Database
Set MthDbP = MthDbzP(CPj)
End Property

Sub BrwMthFb()
BrwFb MthFb
End Sub

Property Get LnoChm() As String()
Erase XX
X "Fld Nm  Md Pj"
X "Fld T50 MchStr"
X "Fld T10 MthPfx"
X "Fld Txt Pjf Prm Ret LinRmk"
X "Fld T3  Ty Mdy"
X "Fld T4  MdTy"
X "Fld Lng Lno"
X "Fld Mem Lines TopRmk"
X "Tbl Pj  *Id Pjf | Pjn PjDte"
X "Tbl Md  *Id PjId Mdn | MdTy"
X "Tbl Mth *Id MdId Mthn ShtTy | ShtMdy Prm Ret LinRmk TopRmk Lines Lno"
LnoChm = XX
Erase XX
End Property
