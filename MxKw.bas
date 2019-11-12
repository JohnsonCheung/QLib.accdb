Attribute VB_Name = "MxKw"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxKw."

Const C_Enm$ = "Enum"
Const C_Prp$ = "Property"
Const C_Ty$ = "Type"
Const C_Fun$ = "Function"
Const C_Sub$ = "Sub"
Const C_Get$ = "Get"
Const C_Set$ = "Set"
Const C_Let$ = "Let"
Const C_Pub$ = "Public"
Const C_Prv$ = "Private"
Const C_Frd$ = "Friend"
Const C_PrpGet$ = C_Prp + " " + C_Get
Const C_PrpLet$ = C_Prp + " " + C_Let
Const C_PrpSet$ = C_Prp + " " + C_Set

Property Get PrpTyAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy(C_Get, C_Set, C_Let)
PrpTyAy = X
End Property

Property Get MthTyAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy(C_Fun, C_Sub, C_PrpGet, C_PrpLet, C_PrpSet)
MthTyAy = X
End Property

Property Get MthMdyAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy(C_Pub, C_Prv, C_Frd, "")
MthMdyAy = X
End Property

Property Get ShtMthMdyAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy("Pub", "Prv", "Frd", "")
ShtMthMdyAy = X
End Property
Property Get ShtMthKdAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy("Fun", "Sub", "Prp")
ShtMthKdAy = X
End Property
Property Get ShtMthTyAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy("Fun", "Sub", "Get", "Set", "Let")
ShtMthTyAy = X
End Property


Property Get MthKdAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy(C_Fun, C_Sub, C_Prp)
MthKdAy = X
End Property

Property Get DclItmAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy(C_Ty, C_Enm)
DclItmAy = X
End Property
