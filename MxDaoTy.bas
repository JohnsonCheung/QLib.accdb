Attribute VB_Name = "MxDaoTy"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoTy."
Public Const ShtTySS$ = "A Att B Bool Byt C Chr D Dbl Dte Dec I Int L Lng M Mem S T Tim Txt"
Function CvDaoTy(A) As DAO.DataTypeEnum
CvDaoTy = A
End Function

Function DaoTyzCol(Col()) As DAO.DataTypeEnum
DaoTyzCol = DaoTyzVbTy(VbTyzCol(Col))
End Function

Function DaoTyzDtaTy(DtaTy$) As DAO.DataTypeEnum
Const CSub$ = CMod & "DaoTy"
Dim O
Select Case DtaTy
Case "Attachment": O = DAO.DataTypeEnum.dbAttachment
Case "Boolean":    O = DAO.DataTypeEnum.dbBoolean
Case "Byte":       O = DAO.DataTypeEnum.dbByte
Case "Currency":   O = DAO.DataTypeEnum.dbCurrency
Case "Date":       O = DAO.DataTypeEnum.dbDate
Case "Decimal":    O = DAO.DataTypeEnum.dbDecimal
Case "Double":     O = DAO.DataTypeEnum.dbDouble
Case "Integer":    O = DAO.DataTypeEnum.dbInteger
Case "Long":       O = DAO.DataTypeEnum.dbLong
Case "Memo":       O = DAO.DataTypeEnum.dbMemo
Case "Single":     O = DAO.DataTypeEnum.dbSingle
Case "Text":       O = DAO.DataTypeEnum.dbText
Case Else: Thw CSub, "Invalid ShtDaoTy", "ShtDaoTy Valid", DtaTy, _
    SyzSS("Attachment Boolean Byte Currency Date Decimal Double Integer Long Memo Signle Text")
End Select
DaoTyzDtaTy = O
End Function

Function DaoTyzShtTy(ShtTy) As DAO.DataTypeEnum
Dim O As DAO.DataTypeEnum
Select Case ShtTy
Case "A", "Att":  O = dbAttachment
Case "B", "Bool":  O = dbBoolean
Case "Byt": O = dbByte
Case "C", "Cur":  O = dbCurrency
Case "Chr": O = dbChar
Case "Dte": O = dbDate
Case "Dec": O = dbDecimal
Case "D", "Dbl":  O = dbDouble
Case "I", "Int":  O = dbInteger
Case "L", "Lng":  O = dbLong
Case "M", "Mem":  O = dbMemo
Case "S", "Sng":  O = dbSingle
Case "T", "Txt":  O = dbText
Case "Tim": O = dbTime
Case Else: Thw CSub, "Invalid ShtTy", "The-Invalid-ShtTy Valid-ShtTy", ShtTy, ShtTyDtaTyLy
End Select
DaoTyzShtTy = O
End Function

Function DaoTy(V) As DAO.DataTypeEnum
Dim T As VbVarType: T = VarType(V)
If T = vbString Then
    If Len(V) > 255 Then
        DaoTy = dbMemo
    Else
        DaoTy = dbText
    End If
    Exit Function
End If
DaoTy = DaoTyzVbTy(T)
End Function

Function DaoTyzVbTy(A As VbVarType) As DAO.DataTypeEnum
Dim O As DAO.DataTypeEnum
Select Case A
Case vbBoolean: O = dbBoolean
Case vbByte: O = dbByte
Case VbVarType.vbCurrency: O = dbCurrency
Case VbVarType.vbDate: O = dbDate
Case VbVarType.vbDecimal: O = dbDecimal
Case VbVarType.vbDouble: O = dbDouble
Case VbVarType.vbInteger: O = dbInteger
Case VbVarType.vbLong: O = dbLong
Case VbVarType.vbSingle: O = dbSingle
Case VbVarType.vbString: O = dbText
Case Else: Thw CSub, "VbTy cannot convert to DaoTy", "VbTy", A
End Select
DaoTyzVbTy = O
End Function

Function DiKqCntzRs(A As DAO.Recordset, Optional Fld = 0) As Dictionary
Set DiKqCntzRs = DiKqCnt(AvRsCol(A))
End Function

Property Get DShtTy() As Drs
Dim Dy(), I
For Each I In SyzSS(ShtTySS)
    PushI Dy, Sy(I, DtaTyzShtTy(I))
Next
DShtTy = DrszFF("ShtTy DtaTy", Dy)
End Property

Function DtaTy$(T As DAO.DataTypeEnum)
Dim O$
Select Case T
Case DAO.DataTypeEnum.dbAttachment: O = "Attachment"
Case DAO.DataTypeEnum.dbBoolean:    O = "Boolean"
Case DAO.DataTypeEnum.dbByte:       O = "Byte"
Case DAO.DataTypeEnum.dbCurrency:   O = "Currency"
Case DAO.DataTypeEnum.dbDate:       O = "Date"
Case DAO.DataTypeEnum.dbDecimal:    O = "Decimal"
Case DAO.DataTypeEnum.dbDouble:     O = "Double"
Case DAO.DataTypeEnum.dbInteger:    O = "Integer"
Case DAO.DataTypeEnum.dbLong:       O = "Long"
Case DAO.DataTypeEnum.dbMemo:       O = "Memo"
Case DAO.DataTypeEnum.dbSingle:     O = "Single"
Case DAO.DataTypeEnum.dbText:       O = "Text"
Case DAO.DataTypeEnum.dbChar:       O = "Char"
Case DAO.DataTypeEnum.dbTime:       O = "Time"
Case DAO.DataTypeEnum.dbLongBinary: O = "LongBinary"
Case Else: Stop
End Select
DtaTy = O
End Function

Property Get DtaTyAy() As String()
DtaTyAy = DtaTyAyzS(ShtTyAy)
End Property

Function DtaTyAyzS(ShtTyAy$()) As String()
Dim ShtTy: For Each ShtTy In Itr(ShtTyAy)
    PushI DtaTyAyzS, DtaTyzShtTy(ShtTy)
Next
End Function

Function DtaTyzShtTy$(ShtTy)
DtaTyzShtTy = DtaTy(DaoTyzShtTy(ShtTy))
End Function

Function DtaTyzTF$(D As Database, T, F$)
DtaTyzTF = DtaTy(FdzTF(D, T, F).Type)
End Function

Function EoShtTyLis(ShtTyLis$) As String()
Dim O$(), ShtTy
For Each ShtTy In CmlAy(ShtTyLis)
    If Not IsShtTy(CStr(ShtTy)) Then
        PushI EoShtTyLis, ShtTy
    End If
Next
End Function

Function IsShtTy(S) As Boolean
Select Case Len(S)
Case 1, 3
    If Not IsAscUCas(Asc(S)) Then Exit Function
    IsShtTy = HasSubStr(ShtTySS, " " & S & " ")
End Select
End Function

Function JnStrDicRsKeyJn(A As DAO.Recordset, KeyFld, JnStrFld, Optional Sep$ = " ") As Dictionary
Dim O As New Dictionary
Dim K, V$
While Not A.EOF
    K = A.Fields(KeyFld).Value
    V = Nz(A.Fields(JnStrFld).Value, "")
    If O.Exists(K) Then
        O(K) = O(K) & Sep & V
    Else
        O.Add K, CStr(Nz(V))
    End If
    A.MoveNext
Wend
Set JnStrDicRsKeyJn = O
End Function

Function JnStrDicTwoFldRs(A As DAO.Recordset, Optional Sep$ = " ") As Dictionary
Set JnStrDicTwoFldRs = JnStrDicRsKeyJn(A, 0, 1, Sep)
End Function

Function MaxSim(A As EmSimTy, B As EmSimTy) As EmSimTy
MaxSim = Max(A, B)
End Function

Function ShtAdoTyAy(A() As ADODB.DataTypeEnum) As String()
Dim I
For Each I In Itr(A)
    PushI ShtAdoTyAy, ShtAdoTy(CLng(I))
Next
End Function

Property Get ShtTyAy() As String()
ShtTyAy = SyzSS(ShtTySS)
End Property

Function ShtTyAyzShtTyLis(ShtTyLis$) As String()
ShtTyAyzShtTyLis = CmlAy(ShtTyLis)
End Function

Property Get ShtTyDtaTyLy() As String()
Dim O$(), I
For Each I In ShtTyAy
    PushI O, I & " " & DtaTyzShtTy(CStr(I))
Next
ShtTyDtaTyLy = AlignLyz2T(O)
End Property

Function ShtTyLiszDaoTyAy$(A() As DataTypeEnum)
Dim O$, I
For Each I In A
    O = O & ShtDaoTy(CvDaoTy(I))
Next
ShtTyLiszDaoTyAy = O
End Function

Function ShtAdoTy$(A As ADODB.DataTypeEnum)
Dim O$
Select Case A
Case ADODB.DataTypeEnum.adTinyInt: O = "Byt"
Case ADODB.DataTypeEnum.adInteger: O = "Lng"
Case ADODB.DataTypeEnum.adSmallInt: O = "Int"
Case ADODB.DataTypeEnum.adDate: O = "Dte"
Case ADODB.DataTypeEnum.adVarChar: O = "Txt"
Case ADODB.DataTypeEnum.adBoolean: O = "Yes"
Case ADODB.DataTypeEnum.adDouble: O = "Dbl"
Case ADODB.DataTypeEnum.adCurrency: O = "Cur"
Case ADODB.DataTypeEnum.adSingle: O = "Sng"
Case ADODB.DataTypeEnum.adDecimal: O = "Dec"
Case ADODB.DataTypeEnum.adVarWChar: O = "Mem"
Case Else: O = "?" & A & "?"
End Select
ShtAdoTy = O
End Function
Function AdoTyStr$(A As ADODB.DataTypeEnum)
Dim O$
Select Case A
Case ADODB.DataTypeEnum.adTinyInt:  O = "Byt"
Case ADODB.DataTypeEnum.adCurrency: O = "C"
Case ADODB.DataTypeEnum.adDecimal:  O = "Dec"
Case ADODB.DataTypeEnum.adDouble:   O = "D"
Case ADODB.DataTypeEnum.adSmallInt: O = "I"
Case ADODB.DataTypeEnum.adInteger:  O = "L"
Case ADODB.DataTypeEnum.adSingle:   O = "S"
Case ADODB.DataTypeEnum.adChar:     O = "Chr"
Case ADODB.DataTypeEnum.adGUID:     O = "G"
Case ADODB.DataTypeEnum.adVarChar:  O = "M"
Case ADODB.DataTypeEnum.adVarWChar: O = "M"
Case ADODB.DataTypeEnum.adLongVarChar: O = "M"
Case ADODB.DataTypeEnum.adBoolean:  O = "B"
Case ADODB.DataTypeEnum.adDate:     O = "Dte"
'Case ADODB.DataTypeEnum.adTime:     O = "Tim"
Case Else
   Thw CSub, "Not supported Case ADODB type", "ADODBTy", A
End Select
AdoTyStr = O
End Function
Function DaoTyStr$(A As DAO.DataTypeEnum)
Dim O$
Select Case A
Case DAO.DataTypeEnum.dbAttachment: O = "A"
Case DAO.DataTypeEnum.dbBoolean:    O = "B"
Case DAO.DataTypeEnum.dbByte:       O = "Byt"
Case DAO.DataTypeEnum.dbCurrency:   O = "C"
Case DAO.DataTypeEnum.dbChar:       O = "Chr"
Case DAO.DataTypeEnum.dbDate:       O = "Dte"
Case DAO.DataTypeEnum.dbDecimal:    O = "Dec"
Case DAO.DataTypeEnum.dbDouble:     O = "D"
Case DAO.DataTypeEnum.dbInteger:    O = "I"
Case DAO.DataTypeEnum.dbLong:       O = "L"
Case DAO.DataTypeEnum.dbMemo:       O = "Mem"
Case DAO.DataTypeEnum.dbSingle:     O = "Single"
Case DAO.DataTypeEnum.dbText:       O = "Text"
Case DAO.DataTypeEnum.dbTime:       O = "Time"
Case DAO.DataTypeEnum.dbTimeStamp:  O = "TimeStamp"
Case Else: Thw CSub, "Unsupported DaoTy, cannot covert to ShtTy", "DaoTy", A
End Select
End Function

Function ShtDaoTy$(A As DAO.DataTypeEnum)
Dim O$
Select Case A
Case DAO.DataTypeEnum.dbAttachment: O = "A"
Case DAO.DataTypeEnum.dbBoolean:    O = "B"
Case DAO.DataTypeEnum.dbByte:       O = "Byt"
Case DAO.DataTypeEnum.dbCurrency:   O = "C"
Case DAO.DataTypeEnum.dbChar:       O = "Chr"
Case DAO.DataTypeEnum.dbDate:       O = "Dte"
Case DAO.DataTypeEnum.dbDecimal:    O = "Dec"
Case DAO.DataTypeEnum.dbDouble:     O = "D"
Case DAO.DataTypeEnum.dbInteger:    O = "I"
Case DAO.DataTypeEnum.dbLong:       O = "L"
Case DAO.DataTypeEnum.dbMemo:       O = "Mem"
Case DAO.DataTypeEnum.dbSingle:     O = "S"
Case DAO.DataTypeEnum.dbText:       O = "T"
Case DAO.DataTypeEnum.dbTime:       O = "Tim"
Case Else: Thw CSub, "Unsupported DaoTy, cannot covert to ShtTy", "DaoTy", A
End Select
ShtDaoTy = O
End Function

Function SimTy(V) As EmSimTy
SimTy = SimTyzV(VarType(V))
End Function

Function SimTyzCol(Col()) As EmSimTy
Dim V: For Each V In Itr(Col)
    Dim O As EmSimTy: O = MaxSim(O, SimTy(V))
    If O = EiStr Then SimTyzCol = O: Exit Function
Next
End Function

Function SimTyzLo(L As ListObject) As EmSimTy()
Dim Sq(): Sq = SqzLo(L)
Dim C%: For C = 1 To UBound(Sq, 2)
    PushI SimTyzLo, SimTyzCol(ColzSq(Sq, C))
Next
End Function

Function SimTyzV(V As VbVarType) As EmSimTy
Dim O As EmSimTy
Select Case True
Case V = Empty: O = EiEmp
Case V = vbBoolean: O = EiYes
Case V = vbByte, V = vbCurrency, V = vbDecimal, V = vbDouble, V = vbInteger, V = vbLong, V = vbSingle: O = EiNum
Case V = vbDate: O = EiDte
Case V = vbString: O = EiStr
End Select
SimTyzV = O
End Function

Function SqlTyzDao$(T As DAO.DataTypeEnum, Optional Si%, Optional Precious%)
Stop '
End Function

Function MaxVbTy(A As VbVarType, B As VbVarType) As VbVarType
Dim O As VbVarType: O = vbBoolean
If IsArray(A) Then Stop
If B >= vbArray Then Stop
Select Case True
Case A = vbBoolean And B = vbBoolean: O = A
Case A = vbBoolean And B = VbVarType.vbByte: O = A
Case A = vbBoolean And B = VbVarType.vbByte: O = A
Case Else: Thw CSub, "VbTy-A-&-B cannot be compared", "VarType-A B", ShtVbTy(A), ShtVbTy(B)
End Select
End Function

Function ShtVbTy$(A As VbVarType)
End Function

Function MaxVbTyzA(Ay() As VbVarType) As VbVarType
Dim O As VbVarType: O = vbBoolean
Dim I: For Each I In Itr(Ay)
    O = MaxVbTy(O, CvVbTy(I))
Next
End Function

Function VbTyzCol(Col()) As VbVarType
Dim I: I = Itr(Col)
Select Case True
Case IsItrNull(I)
Case IsItrStr(I)
Case IsItrNum(I)
Case IsItrDte(I)
Case IsItrBool(I)
End Select
End Function
Function IsItrNull(Itr) As Boolean

End Function


Function VbTyAy(Ay) As VbVarType()
Dim V: For Each V In Ay
    PushI VbTyAy, VarType(V)
Next
End Function

Function ShtTyzCol(Col()) As EmSimTy
Stop
End Function

