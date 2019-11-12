Attribute VB_Name = "MxIsItr"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxIsItr."

Function SomFTrue(Itr, P As IPred) As Boolean
Dim I: For Each I In Itr
    If P.Pred(I) Then SomFTrue = True: Exit Function
Next
End Function
Function AllFalse(Itr, P As IPred) As Boolean
Dim I: For Each I In Itr
    If P.Pred(I) Then Exit Function
Next
AllFalse = False
End Function

Function PredIsLines() As IPred: Static X As New PredIsLines: Set PredIsLines = X: End Function
Function PredIsNm() As IPred:    Static X As New PredIsNm:    Set PredIsNm = X:    End Function
Function PredIsPrim() As IPred:  Static X As New PredIsPrim:  Set PredIsPrim = X:  End Function
Function PredIsStr() As IPred:   Static X As New PredIsStr:   Set PredIsStr = X:   End Function
Function PredIsSy() As IPred:    Static X As New PredIsSy:    Set PredIsSy = X:    End Function

Function PredHasPfx(Pfx$) As IPred
Dim O As New PredHasPfx
O.Init Pfx
Set PredHasPfx = O
End Function

Function PredHasIdf(Idf$) As IPred
Dim O As New PredHasIdf
O.Init Idf
Set PredHasIdf = O
End Function

Function PredHasPatn(Patn$, Optional AndPatn1$, Optional AndPatn2$) As IPred
Dim O As New PredHasPatn
O.Init Patn, AndPatn1, AndPatn2
Set PredHasPatn = O
End Function

Function IsItrStr(Itr) As Boolean:   IsItrStr = AllTrue(Itr, PredIsStr):   End Function
Function IsItrPrim(Itr) As Boolean:  IsItrPrim = AllTrue(Itr, PredIsPrim): End Function
Function IsItrNm(Itr) As Boolean:    IsItrNm = AllTrue(Itr, PredIsNm):     End Function
Function IsItrSy(Itr) As Boolean:    IsItrSy = AllTrue(Itr, PredIsSy):     End Function
Function IsItrLines(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsLines(V) Then Exit Function
Next
IsItrLines = True
End Function

Function PredzKss(Kss$) As IPred
Dim O As New PredLikAy
O.Init SyzSS(Kss)
Set PredzKss = O
End Function
Function IsItrBool(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsBool(V) Then Exit Function
Next
IsItrBool = True
End Function

Function IsItrNum(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsNumeric(V) Then Exit Function
Next
IsItrNum = True
End Function

Function IsItrDte(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsDte(V) Then Exit Function
Next
IsItrDte = True
End Function

