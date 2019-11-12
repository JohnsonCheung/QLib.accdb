Attribute VB_Name = "MxBStmt"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxBStmt."
Const CNs$ = "AlignMth"
Public Const FFoBStmt$ = "L BStmtLin"
Sub RfhBStmt(M As CodeModule)

End Sub

Function DoBStmt(DyoBStmt()) As Drs
DoBStmt = Drs(FoBStmt, DyoBStmt)
End Function

Function FoBStmt() As String()
FoBStmt = SyzSS(FFoBStmt)
End Function

Function DoBStmtzM(M As CodeModule)

End Function
Function DyoBStmt(Src$()) As Variant()
Dim L, J&: For Each L In Itr(Src)
    J = J + 1
    If HasPfx(L, "'@") Then
        PushI DyoBStmt, Array(J, L)
    End If
Next
End Function

Function DoBStmtzS(Src$()) As Drs
DoBStmtzS = Drs(FoBStmt, DyoBStmt(Src))
End Function

Function InspStmtLNewO(Wi_L_MthLin As Drs, Mdn$, Mthn$) As Drs
Dim Bs As Drs:                   Bs = DoBrwStmt(Wi_L_MthLin)               ' L BsLin ! Fst2Chr = '@
Dim Src$():                     Src = StrCol(Wi_L_MthLin, "MthLin")
Dim Di As Dictionary:        Set Di = DiVarnnqDclSfx(Srcc(Src))
                      InspStmtLNewO = XBsLNewO(Bs, Di, Mdn, Mthn$)
End Function

Function DiVarnnqDclSfx(Src$()) As Dictionary
Dim A() As Variant
Stop
End Function

Function XBsLNewO(Bs As Drs, DiVarnnqDclSfx As Dictionary, Mdn$, Mthn$) As Drs
'@Bs   :Drs-L-BsLin ! Fst2Chr = '@
Dim Dy()
    Dim S$, Lin$, L&
    Dim Dr: For Each Dr In Itr(Bs.Dy)
        L = Dr(0)
        Lin = Dr(1)
        If Left(Lin, 2) <> "'@" Then Thw CSub, "BsLin is always begin with '@", "BsLin", Lin
        Dim Varnn$: Varnn = RmvFst2Chr(Lin)
        S = InspStmtzDi(Varnn, DiVarnnqDclSfx, Mdn, Mthn)
        PushI Dy, Array(L, S, Lin)
    Next
XBsLNewO = DrszFF("L NewL OldL", Dy)
End Function

Function DoBrwStmt(Wi_L_MthLin As Drs) As Drs
'Ret :Drs-L-BsLin ! Fst2Chr = '@ @@
Dim Dy()
    Dim IxL%, IxMthLin%: AsgIx Wi_L_MthLin, "L MthLin", IxL, IxMthLin
    Dim Dr: For Each Dr In Itr(Wi_L_MthLin.Dy)
        If HasPfx(Dr(IxMthLin), "'@") Then PushI Dy, Array(Dr(IxL), Dr(IxMthLin))
    Next
DoBrwStmt = DrszFF("L BsLin", Dy)
End Function
