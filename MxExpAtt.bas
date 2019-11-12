Attribute VB_Name = "MxExpAtt"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxExpAtt."


Function ExpAtt$(D As Database, Att$, ToFfn$)
'Ret Exporting the first File in [Att] to [ToFfn] if Att is newer or ToFfn not exist.
'Er if no or more than one file in att, error.
'Er if any, export and return ToFfn. @@
XThwIf_CntNe1 CSub, D, Att
Dim A As Attd: A = XAttd(D, Att)
XThwIf_ExtDif CSub, A, ToFfn
XF2(A).SaveToFile ToFfn
ExpAtt = ToFfn
Inf CSub, "Att is exported", "Att ToFfn FmDb", XAtt(A), ToFfn, D.Name
End Function
Sub ExpAttzRs(R As DAO.Recordset, AttFld, AttFn$, ToFfn$)
End Sub

Function ExpAttzFn$(D As Database, Att$, AttFn$, ToFfn$)
Const CSub$ = CMod & "ExpAttzFn"
If Ext(AttFn) <> Ext(ToFfn) Then
    Thw CSub, "AttFn & ToFfn are dif extEnsion|" & _
        "To export an AttFn to ToFfn, their file extEnsion should be same", _
        "AttFn-Ext ToFfn-Ext D Attk AttFn ToFfn", _
        Ext(AttFn), Ext(ToFfn), D.Name, Att, AttFn, ToFfn
End If
If HasFfn(ToFfn) Then
    Thw CSub, "ToFfn Has, no over write", _
        "D Attk AttFn ToFfn", _
        D.Name, Att, AttFn, ToFfn
End If
Dim Fd2 As DAO.Field2
    Set Fd2 = XF2zFn(D, Att, AttFn$)

If IsNothing(Fd2) Then
    Thw CSub, "In record of Attk there is no given AttFn, but only Act-AttFnAy", _
        "D Given-Attk Given-AttFn Act-AttFny ToFfn", _
        D.Name, Att, AttFn, AttFnAy(D, Att), ToFfn
End If
Fd2.SaveToFile ToFfn
ExpAttzFn = ToFfn
End Function

