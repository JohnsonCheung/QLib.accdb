Attribute VB_Name = "MxRfDic"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxRfDic."

Function DiRfnqRffP() As Dictionary
Set DiRfnqRffP = DiRfnqRffzP(CPj)
End Function

Function DiRfnqRffzP(P As VBProject) As Dictionary
':Rfn: :Nm #Rf-Nm#
Dim Rf As VBIDE.Reference: For Each Rf In P.References
    DiRfnqRffzP.Add Rf.Name, RffzP(P, Rf)
Next
End Function

Function StdRff$(StdRfn$)

End Function
Function DiStdRfnqRff() As Dictionary
Static O As Dictionary
If Not IsNothing(O) Then GoTo X
Set O = New Dictionary
AddDicLin O, "Excel              C:\Program Files (x86)\Microsoft Office\Root\Office16\EXCEL.EXE"
AddDicLin O, "stdole             C:\Windows\SysWOW64\stdole2.tlb"
AddDicLin O, "Office             C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\MSO.DLL"
AddDicLin O, "Access             C:\Program Files (x86)\Microsoft Office\Root\Office16\MSACC.OLB"
AddDicLin O, "Scripting          C:\Windows\SysWOW64\scrrun.dll"
AddDicLin O, "DAO                C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL"
AddDicLin O, "ADODB              C:\Program Files (x86)\Common Files\System\ado\msado15.dll"
AddDicLin O, "ADOX               C:\Program Files (x86)\Common Files\System\ado\msadox.dll"
AddDicLin O, "VBScript_RegExp_55 C:\Windows\SysWOW64\vbscript.dll\3"
AddDicLin O, "VBIDE              C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
AddDicLin O, "MSForms            C:\WINDOWS\SysWOW64\FM20.DLL"
X: Set DiStdRfnqRff = O
End Function
