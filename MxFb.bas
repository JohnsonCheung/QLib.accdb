Attribute VB_Name = "MxFb"
Option Compare Text
Option Explicit
Const CNs$ = "sw"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxFb."

Function ArszFbq(Fb, Q) As ADODB.Recordset
Set ArszFbq = CnzFb(Fb).Execute(Q)
End Function

Sub ArunFbq(Fb, Q)
CnzFb(Fb).Execute Q
End Sub

Sub AsgFbtStr(FbtStr$, OFb$, OT$)
If FbtStr = "" Then
    OFb = ""
    OT = ""
    Exit Sub
End If
AsgBrk FbtStr, "].[", OFb, OT
If FstChr(OFb) <> "[" Then Stop
If LasChr(OT) <> "]" Then Stop
OFb = RmvFstChr(OFb)
OT = RmvLasChr(OT)
End Sub

Function CntrItmNyzFb(Fb) As String()
Dim D As Database: Set D = Db(Fb)
Dim CntrNm
For Each CntrNm In Itn(D.Containers)
    PushIAy CntrItmNyzFb, AmAddPfx(Itn(D.Containers(CntrNm).Documents), CntrNm & ".")
Next
End Function

Function CntrNyzFb(Fb) As String()
CntrNyzFb = Itn(Db(Fb).Containers)
End Function

Function CrtFb(Fb, Optional IsDltFst As Boolean) As Database
If IsDltFst Then DltFfnIf Fb
Set CrtFb = DAO.DBEngine.CreateDatabase(Fb, dbLangGeneral)
End Function

Function Db(Fb) As Database
Set Db = DAO.DBEngine.OpenDatabase(Fb)
End Function

Function DbzFb(Fb) As Database
Set DbzFb = DAO.DBEngine.OpenDatabase(Fb)
End Function

Sub DrpFbt(Fb, T)
CatzFb(Fb).Tables.Delete T
End Sub

Function DrszFbq(Fb, Q) As Drs
DrszFbq = DrszRs(Rs(Db(Fb), Q))
End Function

Function DrszQ(D As Database, Q) As Drs
DrszQ = DrszRs(Rs(D, Q))
End Function

Sub EnsFb(Fb)
If NoFfn(Fb) Then CrtFb Fb
End Sub

Function OupTnyzFb(Fb) As String()
OupTnyzFb = OupTny(Db(Fb))
End Function

Function WszFbq(Fb, Q, Optional Wsn) As Worksheet
'Set WszFbq = WszDrs(DrszFbq(Fb, Q), Wsn:=Wsn)
End Function


Sub Z_BrwFb()
BrwFb SampFbzDutyDta
End Sub

Sub Z_HasFbt()
Ass HasFbt(SampFbzDutyDta, "SkuB")
End Sub

Sub Z_OupTnyzFb()
Dim Fb$
D OupTnyzFb(Fb)
End Sub

Sub Z_TnyzFb()
DmpAy TnyzFb(SampFbzDutyDta)
End Sub

Sub Z_WszFbq()
ShwWs WszFbq(SampFbzDutyDta, "Select * from KE24")
End Sub
