Attribute VB_Name = "MxCv"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCv."

Function MthKdzTy$(MthTy)
Select Case MthTy
Case "Function", "Sub": MthKdzTy = MthTy
Case "Property Get", "Property Let", "Property Set": MthKdzTy = "Property"
End Select
End Function

Function IsMthTy(S) As Boolean
IsMthTy = HasEle(MthTyAy, S)
End Function

Function IsMthMdy(S) As Boolean
IsMthMdy = HasEle(MthMdyAy, S)
End Function

Function MthMdyzSht$(ShtMdy)
Dim O$
Select Case ShtMdy
Case "Pub": O = "Public"
Case "Prv": O = "Private"
Case "Frd": O = "Friend"
Case ""
Case Else: Stop
End Select
MthMdyzSht = O
End Function

Function ShtMdy$(MthMdy)
Dim O$
Select Case MthMdy
Case "Public", "": O = "Pub"
Case "Private": O = "Prv"
Case "Friend": O = "Frd"
Case Else: O = "???"
End Select
ShtMdy = O
End Function

Function MthTyzSht$(ShtMthTy)
Dim O$
Select Case ShtMthTy
Case "Get": O = "Property Get"
Case "Set": O = "Property Set"
Case "Let": O = "Property Let"
Case "Fun": O = "Function"
Case "Sub": O = "Sub"
Case Else: O = "???"
End Select
MthTyzSht = O
End Function

Function ShtMthTy$(MthTy)
Dim O$
Select Case MthTy
Case "Property Get": O = "Get"
Case "Property Set": O = "Set"
Case "Property Let": O = "Let"
Case "Function":     O = "Fun"
Case "Sub":          O = "Sub"
Case Else: O = "???"
End Select
ShtMthTy = O
End Function
Sub Z_ShtMthTyzLin()
GoSub Z
Exit Sub
Z:
    Dim O$(), I, Lin
    For Each I In MthLinAyV
        Lin = I
        PushI O, ShtMthTyzLin(Lin)
    Next
    Brw O
    Return
End Sub
Function ShtMthTyzLin(Lin)
ShtMthTyzLin = ShtMthTy(TakMthTy(RmvMthMdy(Lin)))
End Function

Function ShtMthKdzShtMthTy$(ShtMthTy$)
Dim O$
Select Case ShtMthTy
Case "Get": O = "Prp"
Case "Set": O = "Prp"
Case "Let": O = "Prp"
Case "Fun": O = "Fun"
Case "Sub": O = "Sub"
Case Else: O = "???"
End Select
ShtMthKdzShtMthTy = O
End Function

Function ShtMthKd$(MthKd)
Dim O$
Select Case MthKd
Case "Property": O = "Prp"
Case "Function": O = "Fun"
Case "Sub":      O = "Sub"
Case Else: O = "???"
End Select
ShtMthKd = O
End Function
