Attribute VB_Name = "MxMaybe"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxMaybe."
Type LyRslt: Er() As String: Ly() As String: End Type
Type LyOpt: Som As Boolean: Ly() As String: End Type
Type StrOpt: Som As Boolean: Str As String: End Type
Type BoolOpt: Som As Boolean: Bool As Boolean: End Type
Type DicOpt: Som As Boolean: Dic As Dictionary: End Type
Type LngOpt: Som As Boolean: Lng As Long: End Type
Function LyRslt(Er$(), Ly$()) As LyRslt: LyRslt.Er = Er: LyRslt.Ly = Ly: End Function
Function SomLng(Lng) As LngOpt:               SomLng.Som = True:  SomLng.Lng = Lng:     End Function
Function SomLy(Ly$()) As LyOpt:               SomLy.Som = True:   SomLy.Ly = Ly:        End Function
Function SomStr(Str) As StrOpt:               SomStr.Som = True:  SomStr.Str = Str:     End Function
Function SomBool(Bool As Boolean) As BoolOpt: SomBool.Som = True: SomBool.Bool = Bool:  End Function
Function SomDic(Dic As Dictionary) As DicOpt: SomDic.Som = True:  Set SomDic.Dic = Dic: End Function
Function SomTrue() As BoolOpt:  SomTrue = SomBool(True):   End Function
Function SomFalse() As BoolOpt: SomFalse = SomBool(False): End Function

Function IsEqStrOpt(A As StrOpt, B As StrOpt) As Boolean
Select Case True
Case A.Som And B.Som And A.Str = A.Str: IsEqStrOpt = True
End Select
End Function
