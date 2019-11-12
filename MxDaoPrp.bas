Attribute VB_Name = "MxDaoPrp"
Option Compare Text
Option Explicit
Const CNs$ = "sdfsdf"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoPrp."

Function FldDes$(D As Database, T, F$)
FldDes = FldPv(D, T, F, C_Des)
End Function

Function FldDeszTd$(A As DAO.Field)
FldDeszTd = DaoPv(A.Properties, C_Des)
End Function

Function FldPv(D As Database, T, F$, P$)
If Not HasFldPrp(D, T, F, P) Then Exit Function
FldPv = D.TableDefs(T).Fields(F).Properties(P).Value
End Function

Function HasDbtPrp(D As Database, T, P) As Boolean
HasDbtPrp = HasItn(D.TableDefs(T).Properties, P)
End Function

Function HasFldPrp(D As Database, T, F$, P$) As Boolean
HasFldPrp = HasItn(D.TableDefs(T).Fields(F).Properties, P)
End Function

Function PrpDyoFd(A As DAO.Field) As Variant()
Dim PrpV, I, P$, V
For Each I In Itn(A.Properties)
    V = DaoPv(A, P)
    PushI PrpDyoFd, Array(P, V, TypeName(V))
Next
End Function

Function PrpNyzFd(A As DAO.Field) As String()
PrpNyzFd = Itn(A.Properties)
End Function

Function PrpczO(ObjWiPrpc) As DAO.Properties
On Error GoTo X
Set PrpczO = ObjWiPrpc.Properties
Exit Function
X:
    Dim E$: E = Err.Description
    Thw CSub, "Obj does not have prp-[Properties]", "Obj-TyNm Er", TypeName(ObjWiPrpc), E
End Function

Sub SetFldDes(D As Database, T, F$, Des$)
SetFldPv D, T, F, C_Des, Des
End Sub

Sub SetFldPv(D As Database, T, F$, P$, V)
Dim Fd As DAO.Field: Set Fd = D.TableDefs(T).Fields(F)
SetDaoPv Fd, P, V
End Sub

Sub SetDbtPv(D As Database, T, P$, V)
Dim Td As DAO.TableDef: Set Td = D.TableDefs(T)
SetDaoPv Td, P, V
End Sub

Sub DltDaoPrp(DaoPrpsObj, P$)
Dim Ps As DAO.Properties: Set Ps = DaoPrpsObj.Properties
If HasDaoPrp(Ps, P) Then
    Ps.Delete P
End If
End Sub
Sub SetDaoPv(DaoPrpsObj, P$, V)
':WiDaoPrpsObj: :Obj #With-Dao.Properties-Obj#
Dim Ps As DAO.Properties: Set Ps = DaoPrpsObj.Properties
If HasDaoPrp(Ps, P) Then
    Ps(P).Value = V
Else
    Ps.Append DaoPrpsObj.CreateProperty(P, DaoTy(V), V) ' will break if V=""
End If
End Sub

Function DbtPrp(D As Database, T, P)
If Not HasDbtPrp(D, T, P) Then Exit Function
DbtPrp = D.TableDefs(T).Properties(P).Value
End Function

Sub Y_DbtPrp()
Dim D As Database: Set D = TmpDb
DrpT D, "Tmp"
Rq D, "Create Table Tmp (F1 Text)"
DbtPrp(D, "Tmp", "XX") = "AFdf"
Debug.Assert DbtPrp(D, "Tmp", "XX") = "AFdf"
End Sub


Sub Z_FldPv()
Dim P$, Db As Database, T, F$, V
GoSub T0
Exit Sub
T0:
    Set Db = TmpDb
    Rq Db, "Create Table Tmp (AA Text)"
    T = "Tmp"
    F = "AA"
    P = "Ele"
    V = "Ele1234"
    GoTo Tst
Tst:
    FldPv(Db, T, F, P) = V
    Ass FldPv(Db, T, F, P) = V
    Dim Fd As DAO.Field: Set Fd = FdzTF(Db, T, F)
    Stop
    DmpDy PrpDyoFd(Fd)
    Return
End Sub

Sub Z_PrpDyoFd()
Dim Db As Database: Set Db = SampDbDutyDta
Dim Fd As DAO.Field
Dim Rs As DAO.Recordset
Set Rs = RszT(Db, "Permit")
Set Fd = Rs.Fields("Permit")
Debug.Print Fd.Value
DmpDy PrpDyoFd(Fd)
End Sub

Sub Z_PrpNy()
Dim Db As Database: Set Db = SampDbDutyDta
Dim Fd As DAO.Field
Set Fd = FdzTF(Db, "Permit", "Permit")
D PrpNyzFd(Fd)
End Sub
