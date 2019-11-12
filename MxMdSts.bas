Attribute VB_Name = "MxMdSts"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMdSts."
Type MdSts
    NLin As Long
    NPub As Integer
    NPrv As Integer
    NFrd As Integer
End Type

Function MdStszL(MthLinAy$(), NLin&) As MdSts
With MdStszL
    .NLin = NLin
    Dim L: For Each L In Itr(MthLinAy)
        Select Case MthMdy(L)
        Case "", "Public": .NPub = .NPub + 1
        Case "Private":    .NPrv = .NPrv + 1
        Case "Friend":     .NFrd = .NFrd + 1
        Case Else: Stop
        End Select
    Next
End With
End Function

Function MdSts(Src$()) As MdSts
MdSts = MdStszL(MthLinAy(Src), Si(Src))
End Function

Function MdStsM() As MdSts
MdStsM = MdStszM(CMd)
End Function

Function MdStszM(M As CodeModule) As MdSts
MdStszM = MdSts(SrczM(M))
End Function

