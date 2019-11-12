Attribute VB_Name = "MxCcm"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxCcm."

Function CcmTny(D As Database) As String()
CcmTny = AwPfx(Tny(D), "^")
End Function

Sub Chk(D As Database, CcmTny$(), IsLcl As Boolean)
Const CSub$ = CMod & "Chk"
If Not IsLcl Then Chk2 D, CcmTny: Exit Sub ' Asserting for TarFb is stored in CcmTny's description

'Asserting for TarFb = Db
Dim Mis$(): Mis = Chk1(D, CcmTny)
If Si(Mis) = 0 Then Exit Sub
Thw CSub, "[Some-Missing-Tar-Tbl] in [Db] cannot be found according to given [CcmTny] in [Db]", "Mis Dbn CcmTny", Mis, D.Name, CcmTny
End Sub

Function Chk1(D As Database, CcmTny$()) As String()
Dim N1$(): N1 = Tny(D)
Dim N2$(): N2 = RmvFstChrzAy(CcmTny)
Chk1 = AyMinus(N2, N1)
End Function

Sub Chk2(D As Database, CcmTny$())
'Throw if any Corresponding-Table in TarFb is not found
Dim O$(), T$, I
For Each I In CcmTny
    T = I
    PushIAy O, Chk3(D, T)
Next
'ErThw O
End Sub

Function Chk3(D As Database, CcmTbl$) As String()
Dim TarFb$
'    TarFb = Dbt_Des(Db, CcmTbl)
Select Case True
Case TarFb = "":            Chk3 = LyzMsgNap("[CcmTbl] in [Db] should have 'Des' which is TarFb, but this TarFb is blank", CcmTbl, D.Name)
'Case NotHasFfn(TarFb):    Chk3 = LyzMsgNap("[CcmTbl] in [Db] should have [Des] which is TarFb, but this TarFb does not Has", CcmTbl, D.Name, TarFb)
Case Not HasFbt(TarFb, RmvFstChr(CcmTbl)):
    Chk3 = LyzMsgNap("[CcmTbl] in [Db] should have [Des] which is TarFb, but this TarFb does not Has [Tbl-RmvFstChr(CcmTbl)]", CcmTbl, D.Name, TarFb, RmvFstChr(CcmTbl))
End Select
End Function

Sub LnkCcm(D As Database, IsLcl As Boolean)
'Ccm stands for Space-[C]ir[c]umflex-accent
'CcmTbl is ^xxx table in Db (pgm-database),
'          which should be same stru as N:\..._Data.accdb @ xxx
'          and   data should be copied from N:\..._Data.accdb for development purpose
'At the same time, in Db, there will be xxx as linked table either
'  1. In production, linking to N:\..._Data.accdb @ xxx
'  2. In development, linking to Db @ ^xxx
'Notes:
'  The TarFb (N:\..._Data.accdb) of each CcmTbl may be diff
'      They are stored in Description of CcmTbl manual, it is edited manually during development.
'  those xxx table in Db will be used in the program.
'  and ^xxx is create manually in development and should be deployed to N:\..._Data.accdb
'  assume Db always have some ^xxx, otherwise throw
'This Sub is to re-link the xxx in given [Db] to
'  1. [Db] if [TarFb] is not given
'  2. [TarFb] if [TarFb] is given.
Const CSub$ = CMod & "LnkCcm"
Dim T$()  ' All ^xxx
    T = CcmTny(D)
    If Si(T) = 0 Then Thw CSub, "No ^xxx table in [Db]", D.Name 'Assume always
Chk D, T, IsLcl ' Chk if all T after rmv ^ is in TarFb
LnkCcmzTny D, T, IsLcl
End Sub

Sub LnkCcmzTny(D As Database, CcmTny$(), IsLcl As Boolean)
Const CSub$ = CMod & "LnkCcmzTny"
Dim CcmTbl$, I, TarFb$
TarFb = D.Name
For Each I In CcmTny
    CcmTbl = I
    If FstChr(CcmTbl) <> "^" Then Thw CSub, "All table in CcmTny must begin ^", "Tbl-without-^ CcmTny", CcmTbl, CcmTny
    LnkFb D, RmvFstChr(CcmTbl), TarFb, CcmTbl
Next
End Sub


Sub Z_CcmTny()
Dim D As Database
'
Set D = Db(SampFbzShpRate)
Ept = SyzSS("^CurYM ^IniRate ^IniRateH ^InvD ^InvH ^YM ^YMGR ^YMGRnoIR ^YMOH ^YMRate")
GoSub Tst
Exit Sub
Tst:
    Act = CcmTny(D)
    C
    Return
End Sub

Sub Z_LnkCcm()
Dim D As Database, IsLcl As Boolean
Set D = Db(SampFbzShpRate)
IsLcl = True
GoSub Tst
Exit Sub
Tst:
    LnkCcm D, IsLcl
    Return
End Sub
