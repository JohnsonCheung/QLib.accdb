Attribute VB_Name = "MxTyDi"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxTyDi."
Function DiFldnqShtDaoTyDic(FxOrFb$, TorW) As Dictionary
Select Case True
Case IsFb(FxOrFb): Set DiFldnqShtDaoTyDic = DiFldnqShtDaoTyzFbt(FxOrFb, TorW)
Case IsFx(FxOrFb): Set DiFldnqShtDaoTyDic = DiFldnqShtAdoTyzFxw(FxOrFb, TorW)
Case Else: Thw CSub, "FxOrFb should be Fx or Fb", "FxOrFb TorW", FxOrFb, TorW
End Select
End Function

Function DiFldnqShtDaoTyzFbt(Fb, T) As Dictionary
Set DiFldnqShtDaoTyzFbt = New Dictionary
Dim F As DAO.Field: For Each F In Db(Fb).TableDefs(T).Fields
    DiFldnqShtDaoTyzFbt.Add F.Name, ShtDaoTy(F.Type)
Next
End Function

Function DiFldnqShtAdoTyzFxw(Fx, Optional W = "Sheet1") As Dictionary
Set DiFldnqShtAdoTyzFxw = New Dictionary
Dim Cat As Catalog: Set Cat = CatzFx(Fx)
Dim C As ADOX.Column: For Each C In Cat.Tables(CattnzWsn(W)).Columns
    DiFldnqShtAdoTyzFxw.Add C.Name, ShtAdoTy(C.Type)
Next
End Function

Sub Z_DiFldnqShtAdoTyzFxw()
BrwDic DiFldnqShtAdoTyzFxw(SampFx)
End Sub
