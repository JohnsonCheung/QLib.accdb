Attribute VB_Name = "MxGenLibPj"
Option Compare Database

Sub GenLibFbaP(Libv$)
GenFbaByLibzP CPj, Libv
End Sub

Sub GenLibFbazP(P As VBProject, Libv$)
':LibFba: :Fba #Library-Fba# ! It is a subset of @P.  The modules to be gen in @P are those const-CLib = @Libv.
End Sub

