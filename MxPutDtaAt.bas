Attribute VB_Name = "MxPutDtaAt"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxPutDtaAt."

Function PutDbtAt(Db As Database, T, At As Range) As Range
Set PutDbtAt = RgzSq(SqzT(Db, T), At)
End Function

