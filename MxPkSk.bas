Attribute VB_Name = "MxPkSk"
Option Compare Text
Option Explicit
Const CNs$ = "a"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxPkSk."

Function ChkPk$(D As Database, T)
If HasStdPk(D, T) Then Exit Function
If HasPk(D, T) Then
    Dim Pk$(): Pk = PkFny(D, T)
    Select Case True
    Case Si(Pk) <> 1: ChkPk = FmtQQ("There is PrimaryKey-Idx, but it has [?] fields[?]", Si(Pk), TLin(Pk))
    Case Pk(0) <> T & "Id": ChkPk = FmtQQ("There is One-field-PrimaryKey-Idx of FldNm(?), but it should named as ?Id", Pk(0), T)
    Case FdzTF(D, T, 0).Name <> T & "Id": ChkPk = FmtQQ("The Pk-field(?Id) should be first fields, but now it is (?)", T, FdzTF(D, T, T & "Id").OrdinalPosition)
    End Select
End If
ChkPk = "[?] does not have PrimaryKey-Idx"
End Function

Function ChkPkSk(D As Database) As String()
Dim T$, I
For Each I In Tny(D)
    T = I
    PushIAy ChkPkSk, ChkPkSkzT(D, T)
Next
End Function

Function ChkPkSkzT(D As Database, T) As String()
PushNB ChkPkSkzT, ChkPk(D, T)
PushNB ChkPkSkzT, ChkSk(D, T)
End Function

Function ChkSk$(D As Database, T)
Dim SkIdx As DAO.Index, I As DAO.Index
If Not HasIdx(D, T, Skn) Then
    ChkSk = FmtQQ("Not SecondaryKey for Table[?] in Db[?]", T, D.Name)
    Exit Function
End If
Set SkIdx = D.TableDefs(T).Indexes(Skn)
Select Case True
Case Not SkIdx.Unique
    ChkSk = FmtQQ("SecondaryKey is not unique for Table[?] in Db[?]", T, D.Name)
Case Else
    Set I = FstUniqIdx(D, T)
    If Not IsNothing(I) Then
 '       ChkSk = FmtQQ("No SecondaryKey, but there is uniq idx, it should name as SecondaryKey for Table[?] Db[?] UniqIdxNm[?] IdxFny[?]", _
            T, D.Name, I.Name, JnTermAy(FnyzIdx(I)))
    End If
End Select
End Function

Function ChkSsk$(D As Database, T)
Dim O$, Sk$(): Sk = SkFny(D, T)
O = ChkSk(D, T): If O <> "" Then ChkSsk = O: Exit Function
If Si(Sk) <> 1 Then
'    ChkSsk = FmtQQ("Secondary is not single field. Tbl[?] Db[?] SkFfn[?]", T, D.Name, JnTermAy(Sk))
End If
End Function
