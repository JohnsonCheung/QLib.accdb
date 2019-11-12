Attribute VB_Name = "MxGenLibPj"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxGenLibPj."

Sub GenLibFbaP(Libv$)
GenLibFbazP CPj, Libv
End Sub

Sub GenLibFbazP(P As VBProject, Libv$)
':LibFba: :Fba #Library-Fba# ! It is a subset of @P.  The modules to be gen in @P are those const-CLib = @Libv.
End Sub

