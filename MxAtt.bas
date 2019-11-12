Attribute VB_Name = "MxAtt"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxAtt."
Type Attd
    TblRs As DAO.Recordset '..Att.. #Tbl-Rs ! It is the Tbl-Att recordset
    AttRs As DAO.Recordset '.       #Att-Rs2 !
End Type

Function AttFn$(D As Database, Att$)
'Ret : fst attachment fn in the att fld of att tbl, if no fn, return blnk @@
Const CSub$ = CMod & "AttFnzAttd"
Dim A As Attd: A = XAttd(D, Att) ' if @Att in exist in Tbl-Att, a rec will created
With A.AttRs
    If .EOF Then
        If .BOF Then
            Inf CSub, "[Attk] has no attachment files", "Attk", Attk(A)
            Exit Function
        End If
    End If
    .MoveFirst
    AttFn = !FileName
End With
End Function

Function AttFnAy(D As Database, Att$) As String()
Dim R As Attd: R = XAttd(D, Att)
AttFnAy = SyzRs(R.AttRs, "FileName")
End Function

Function Attk$(A As Attd)
Attk = A.TblRs!Attk
End Function

Function AttSi&(D As Database, Att$)
AttSi = FvzSsk(D, "Att", "FilSz", Av(Att))
End Function

Function AttTim(D As Database, Att$) As Date
AttTim = FvzSsk(D, "Att", "FilTim", Av(Att))
End Function

Function DAttFld(A As Attd) As Drs
'Ret : :Drs:Fldn DtaTy Si: of the Fld-Att of Tbl-Att.  The Fld-Att is Dao.Recordset2
DAttFld = DFldzRs(A.AttRs)
End Function

Function DAttFldzDb(D As Database) As Drs
'Ret : :Drs:Fldn DtaTy Si: from @D assume there is table-Att
DAttFldzDb = DAttFld(XAttd(D, "Sample"))
End Function

Function DFldzRs(R As DAO.Recordset2) As Drs
'Ret : :Drs:Fldn DtaTy Si: the @R
Dim Dy(), F As DAO.Field2: For Each F In R.Fields
    Dim N$: N = F.Name
    Dim T$: T = DtaTy(F.Type)
    Dim S%: S = F.Size
    PushI Dy, Array(N, T, S)
Next
DFldzRs = DrszFF("Fldn DtaTy Si", Dy)
End Function

Sub DltAtt(D As Database, Att$)
D.Execute FmtQQ("Delete * from Att where Attn='?'", Att)
End Sub

Function DoTblAtt(D As Database) As Drs
DoTblAtt = DrszT(D, "Att")
End Function


Function FnyzAttFld(Rs As DAO.Recordset2, AttFld$) As String()
FnyzAttFld = FnyzRs(CvRs(Rs.Fields(AttFld).Value))
End Function

Function FnyzAttTbl(D As Database) As String()
FnyzAttTbl = Fny(D, "Att")
End Function

Sub ImpAtt(D As Database, Att$, FmFfn$)
Dim F2 As DAO.Field2
'Msg CSub, "[Att] is going to import [Ffn] with [Si] and [Tim]", Fv(A.TblRs!Attk), Ffn, S, T
Dim A As Attd: A = XAttd(D, Att)
Dim T As DAO.Recordset2: Set T = A.TblRs ' The Tbl-Rs of Tbl-Att
    T.Edit
    With A.AttRs
        If HasReczFEv(A.AttRs, "FileName", Fn(FmFfn)) Then
            Dmp "Ffn is found in Att and it is replaced"
            .Edit
        Else
            Dmp "Ffn is not found in Att tbl and it is IMPORTED.  Ffn[" & FmFfn & "]"
            .AddNew
        End If
        Set F2 = !FileData
        F2.LoadFromFile FmFfn
        .Update
    End With
    A.TblRs.Fields!FilTim = DtezFfn(FmFfn)
    A.TblRs.Fields!FilSi = SizFfn(FmFfn)
    A.TblRs.Update
End Sub

Function IsAttOld(D As Database, Att$, Ffn$) As Boolean
Const CSub$ = CMod & "IsAttOld"
Dim ATim$:   ATim = AttTim(D, Att)
Dim FTim$:   FTim = DtezFfn(Ffn)
Dim AttIs$: AttIs = IIf(ATim > FTim, "new", "old")
Dim M$:         M = "Att is " & AttIs
Inf CSub, M, "Att Ffn AttTim FfnTim AttIs-Old-or-New?", Att, Ffn, ATim, FTim, AttIs
End Function

Function IsAttOneFil(D As Database, Att$) As Boolean
Debug.Print "DbAttHasOnlyFile: " & XAttd(D, Att).AttRs.RecordCount
IsAttOneFil = XAttd(D, Att).AttRs.RecordCount = 1
End Function

Function NAtt%(D As Database, Att$)
NAtt = NAttzAttd(XAttd(D, Att))
End Function

Function NAttzAttd%(D As Attd)
NAttzAttd = NReczRs(D.AttRs)
End Function

Function TmpAttDb() As Database
'Ret: a tmp db with tbl-att @@
Dim O As Database: Set O = TmpDb
EnsTblAtt O
Set TmpAttDb = O
End Function

Function XAtt$(A As Attd)
XAtt = A.TblRs!Att
End Function

Function XAttd(D As Database, Att$) As Attd
'Ret: :Attd ! which keeps :TblRs and :AttRs opened,
'           ! where :TblRs is poiting the rec in tbl-att, if fnd just point to it, if not fnd, add one rec with Attn=@Att
'           ! and   :AttRs is pointing to the :FileData of the fld-Att of the tbl-Att
Dim Q$: Q = FmtQQ("Select Att,FilTim,FilSi from Att where Attn='?'", Att)
If Not HasReczQ(D, Q) Then
    D.Execute FmtQQ("Insert into Att (Attk) values('?')", Att) ' add rec to tbl-att with Att=@Att
End If
With XAttd
    Set .TblRs = Rs(D, Q)
    Set .AttRs = .TblRs.Fields(0).Value ' there is always a rec of Att=@Att in .TblRs (Tbl-Att)
End With
End Function

Function XAttn$(A As Attd)
XAttn = A.TblRs!Attk
End Function

Function XAttNy(D As Database) As String()
XAttNy = SyzRs(Rs(D, "Select Attk from Att order by Attk"))
End Function

Function XF2(A As Attd) As DAO.Field2
Set XF2 = A.AttRs!FileData
End Function

Function XF2zFn(D As Database, Att$, AttFn$) As DAO.Field2
With XAttd(D, Att)
    With .AttRs
        .MoveFirst
        While Not .EOF
            If !FileName = AttFn Then
                Set XF2zFn = !FileData
            End If
            .MoveNext
        Wend
    End With
End With
End Function

Sub XThwIf_CntNe1(Fun$, D As Database, Att$)
Dim N%: N = NAtt(D, Att)
If N <> 1 Then
    Thw Fun, "Attk should have only one file, no export.", _
        "Attk FilCnt D", _
        Att, N, D.Name
End If
End Sub

Sub XThwIf_ExtDif(Fun$, A As Attd, ToFfn$)
With A.AttRs
    If Ext(!FileName) <> Ext(ToFfn) Then Thw Fun, "The Ext in the Att should be same", "Att-Ext ToFfn-Ext", Ext(!FileName), Ext(ToFfn)
End With
End Sub


Sub Z_AttFnAy()
D AttFnAy(SampDbShpCst, "AA")
End Sub

Sub Z_ExpAtt()
Dim T$, D As Database
T = TmpFx
ExpAttzFn D, "Tp", "TaxRateAlert(Template).xlsm", T
Debug.Assert HasFfn(T)
Kill T
End Sub

Sub Z_ImpAtt()
Dim T$, D As Database
T = TmpFt
WrtStr "sdfdf", T
ImpAtt D, "AA", T
Kill T
'T = TmpFt
'ExpAttToFfn "AA", T
'BrwFt T
End Sub

