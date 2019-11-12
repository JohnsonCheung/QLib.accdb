Attribute VB_Name = "MxAdoDta"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxAdoDta."
Function DrszCnq(Cn As ADODB.Connection, Q) As Drs
DrszCnq = DrszArs(ArszCnq(Cn, Q))
End Function

Function DrszFbqAdo(Fb, Q) As Drs
DrszFbqAdo = DrszArs(ArszFbq(Fb, Q))
End Function

Function DrszArs(A As ADODB.Recordset) As Drs
DrszArs = Drs(FnyzArs(A), DyoArs(A))
End Function

