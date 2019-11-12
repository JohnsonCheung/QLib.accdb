Attribute VB_Name = "MxStru"
Option Compare Text
Option Explicit
Const CNs$ = "sdfsdf"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxStru."

Sub DmpStru(D As Database)
Dmp Stru(D)
End Sub

Sub DmpStruTT(D As Database, TT$)
Dmp StruzTT(D, TT)
End Sub

Function StruFld(ParamArray Ap()) As Drs
Dim Dy(), S$, I, Av(), Ele$, LikFF$, LikFld$, J
Av = Ap
For Each I In Av
    S = I
    AsgTRst S, Ele, LikFF
    For Each J In SyzSS(LikFF)
        LikFld = J
        PushI Dy, Array(Ele, LikFld)
    Next
Next
StruFld = DrszFF("Ele FldLik", Dy)
End Function

Function StruInf(D As Database) As Dt
Dim T$, TT, Dy(), Des$, NRec&, Stru$
'For Each TT In TnyDb(D)
    T = TT
'    Des = Dbt_Des(D, T)
'    Stru = RmvT1(Stru(D, T))
'    NRec = NRecDT(D, T)
    PushI Dy, Array(T, NRec, Des, Stru)
'Next
StruInf = DtzFF("Tbl", "Tbl NRec Des", Dy)
End Function
Function Stru(D As Database) As String()
Stru = AlignLyz1T(StruzTny(D, Tny(D)))
End Function

Function StruzTny(D As Database, Tny$()) As String()
Dim I
For Each I In Itr(AySrtQ(Tny))
    PushI StruzTny, StruzT(D, CStr(I))
Next
End Function

Function StruzRs$(A As DAO.Recordset)
Dim O$(), F As DAO.Field2
For Each F In A.Fields
    PushI O, FdStr(F)
Next
StruzRs = JnCrLf(O)
End Function

Function StruT$(T)
StruT = StruzT(CurrentDb, T)
End Function

Function StruzT$(D As Database, T)
Dim F$()
    F = Fny(D, T)
    F = RplStarzAy(F, T)
    F = SyzQteSqIf(F)

Dim Pk$()
    Pk = PkFny(D, T)
    Pk = RplStarzAy(Pk, T)
    Pk = SyzQteSqIf(Pk)
    
Dim P$
    P = JnSpc(Pk)
    If P <> "" Then P = " " & P & " |"

Dim R$
    Dim Rst$()
    Rst = AyMinus(F, Pk)
    R = " " & JnSpc(SyzQteSqIf(Rst))
StruzT = T & P & R
End Function

Function StruzTT(D As Database, TT$)
StruzTT = StruzTny(D, Ny(TT))
End Function

