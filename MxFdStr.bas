Attribute VB_Name = "MxFdStr"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxFdStr."
Public Const DaoTynn$ = "Boolean Byte Integer Int Long Single Double Char Text Memo Attachment" ' used in TzPFld

Function FdStr$(A As DAO.Field2)
Dim D$, R$, Z$, VTxt$, VRul, E$, S$
If A.Type = DAO.DataTypeEnum.dbText Then S = " TxtSz=" & A.Size
If A.DefaultValue <> "" Then D = "Dft=" & A.DefaultValue
If A.Required Then R = "Req"
If A.AllowZeroLength Then Z = "AlwZLen"
If A.Expression <> "" Then E = "Expr=" & A.Expression
If A.ValidationRule <> "" Then VRul = "VRul=" & A.ValidationRule
If A.ValidationText <> "" Then VTxt = "VTxt=" & A.ValidationText
FdStr = TLinzAp(A.Name, ShtDaoTy(A.Type), R, Z, VTxt, VRul, D, E, IIf((A.Attributes And DAO.FieldAttributeEnum.dbAutoIncrField) <> 0, "Auto", ""))
End Function

Function FdzFdStr1(FdStr) As DAO.Field2
Dim N$, S$ ' #Fldn and #EleStr
Dim O As DAO.Field2
AsgBrkSpc FdStr, N, S
Select Case True
Case S = "Boolean":  Set O = FdzBool(N)
Case S = "Byte":     Set O = FdzByt(N)
Case S = "Integer", S = "Int": Set O = FdzInt(N)
Case S = "Long":     Set O = FdzLng(N)
Case S = "Single":   Set O = FdzSng(N)
Case S = "Double":   Set O = FdzDbl(N)
Case S = "Currency": Set O = FdzCur(N)
Case S = "Char":     Set O = FdzChr(N)
Case HasPfx(S, "Text"): Set O = FdzTxt(N, BetBkt(S))
Case S = "Memo":     Set O = FdzMem(N)
Case S = "Attachment": Set O = FdzAtt(N)
Case S = "Time":     Set O = FdzTim(N)
Case S = "Date":     Set O = FdzDte(N)
Case Else: Thw CSub, "Invalid FdStr", "Nm Spec vdt-DaoTynn, N, S, DaoTynn"
End Select
Set FdzFdStr1 = O
End Function

Function FdzFdStr(FdStr$) As DAO.Field2
Dim F$, ShtTy$, Req As Boolean, AlwZLen As Boolean, Dft$, VTxt$, VRul$, TxtSz As Byte, Expr$
Dim L$: L = FdStr
Dim Vy(): Vy = ShfVy(L, EleLblss)
AsgAp Vy, _
    F, ShtTy, Req, AlwZLen, Dft, VTxt, VRul, TxtSz, Expr
Set FdzFdStr = Fd( _
    F, DaoTyzShtTy(ShtTy), Req, TxtSz, AlwZLen, Expr, Dft, VRul, VTxt)
End Function

Function FdStrAyFds(A As DAO.Fields) As String()
Dim F As DAO.Field
For Each F In A
    PushI FdStrAyFds, FdStr(F)
Next
End Function

Function FdzStr(FdStr$) As DAO.Field2
End Function

Sub Z_FdzFdStr()
Dim Act As DAO.Field2, Ept As DAO.Field2, mFdStr$
mFdStr = "AA Int Req AlwZLen Dft=ABC TxtSz=10"
Set Ept = New DAO.Field
With Ept
    .Type = DAO.DataTypeEnum.dbInteger
    .Name = "AA"
    '.AllowZeroLength = False
    .DefaultValue = "ABC"
    .Required = True
    .Size = 2
End With
GoSub Tst
Exit Sub
Tst:
    Set Act = FdzFdStr(mFdStr)
    If Not IsEqFd(Act, Ept) Then
        D LyzMsgNap("Act", "FdStr", FdStr(Act))
        D LyzMsgNap("Ept", "FdStr", FdStr(Ept))
    End If
    Return
End Sub

Sub Z_FdzFdStr1()
Dim FdStr$
FdStr = "Txt Req Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']"
GoSub Tst
Exit Sub
Tst:
    Set Act = FdzFdStr(FdStr)
    Stop
    Return
End Sub


Function FdStrAy(D As Database, T) As String()
Dim F, Td As DAO.TableDef
Set Td = D.TableDefs(T)
For Each F In Fny(D, T)
    PushI FdStrAy, FdStr(Td.Fields(F))
Next
End Function

Function FdStrzTF$(D As Database, T, F$)
FdStrzTF = FdStr(FdzTF(D, T, F$))
End Function



