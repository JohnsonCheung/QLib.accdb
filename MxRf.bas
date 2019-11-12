Attribute VB_Name = "MxRf"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxRf."
Type RfLin: Lin As String: End Type
':Frf$ = "It a file Rf.Txt in Srcp with RfLin"
Public Const FFoRfLin$ = "Nm Guid Mjr Mnr Rff"
Function RfLin(Lin) As RfLin
RfLin.Lin = Lin
End Function

Function RfLinzRf(A As VBIDE.Reference) As RfLin
With A
RfLinzRf = RfLin(JnSpcAp(.Name, .GUID, .Major, .Minor, .FullPath))
End With
End Function
Function RffzRfLin(A As RfLin)
Dim P%: P = InStr(Replace(A.Lin, " ", "-", Count:=3), " ")
RffzRfLin = Mid(A.Lin, P + 1)
End Function

Function HasRf(P As VBProject, Rfn) As Boolean
HasRf = HasItn(P.References, Rfn)
End Function

Function HasRff(P As VBProject, Rff) As Boolean
HasRff = HasItrEq(P.References, "FullPath", Rff)
End Function
Property Get FrfC$()
FrfC = Frf(CPj)
End Property

Function FrfzSrcp$(Srcp$)
FrfzSrcp = EnsPthSfx(Srcp) & "Rf.txt"
End Function

Function Frf$(P As VBProject)
Frf = FrfzSrcp(SrcpzP(P))
End Function

Function FrfzDistPj$(DistPj As VBProject)
FrfzDistPj = SrcpzDistPj(DistPj) & "Rf.txt"
End Function

Function RfSrczSrcp(Srcp$) As String()
RfSrczSrcp = LyzFt(FrfzSrcp(Srcp))
End Function
Function RfSrc(P As VBProject) As String()
Dim R As VBIDE.Reference
For Each R In P.References
    PushI RfSrc, RfLinzRf(R).Lin
Next
End Function

Function CvRf(A) As VBIDE.Reference
Set CvRf = A
End Function

Function HasRfn(Pj As VBProject, Rfn)
Dim Rf As VBIDE.Reference
For Each Rf In Pj.References
    If Rf.Name = Rfn Then HasRfn = True: Exit Function
Next
End Function

Function HasRfGuid(P As VBProject, RfGuid)
HasRfGuid = HasItrEq(P.References, "GUID", RfGuid)
End Function

Function RffyzP(P As VBProject) As String()
RffyzP = SyzItrPrp(P.References, "FullPath")
End Function

Function RffPjn$(P As VBProject, Rfn$)
RffPjn = Pjp(P) & Rfn & ".xlam"
End Function

Function Rf(Rfn) As VBIDE.Reference
Set Rf = RfzP(CPj, Rfn)
End Function

Function RfzP(P As VBProject, Rfn) As VBIDE.Reference
Set RfzP = ItwNm(P.References, Rfn)
End Function
Function RfNyP() As String()
RfNyP = RfNyzP(CPj)
End Function

Function RfNyzP(P As VBProject) As String()
RfNyzP = Itn(P.References)
End Function

Function Rff$(Rfn)
':Rff: :Ffn #Rf-FileName# ! a .dll file or .mda or .fxa to be referred by a pj
Rff = RffzP(CPj, Rfn)
End Function

Function RffzP$(P As VBProject, Rfn)
On Error Resume Next
RffzP = RfzP(P, Rfn).FullPath
End Function


Sub DmpPjRfP()
DmpDrs DoPjRfP
End Sub

Sub Z_DoPjRfP()
BrwDrs DoPjRfP
End Sub

Function DoPjRfP() As Drs
DoPjRfP = DoPjRfzP(CPj)
End Function

Function DoPjRfzP(P As VBProject) As Drs
Dim Prpcc$: Prpcc = "Name GUID Major Minor FullPath Description BuiltIn Type IsBroken"
Dim A As Drs: A = DrszItrPrpcc(P.References, Prpcc)
DoPjRfzP = InsCol(A, "Pj", P.Name)
End Function

Function DoPjRfUsr() As Drs
Erase XX
X "MVb"
X "MIde  MVb MXls MAcs"
X "MXls  MVb"
X "MDao  MVb MDta"
X "MAdo  MVb"
X "MAdoX MVb"
X "MApp  MVb"
X "MDta  MVb"
X "MTp   MVb"
X "MSql  MVb"
X "AStkShpCst MVb MXls MAcs"
X "MAcs  MVb MXls"
DoPjRfUsr = DrszTRstLy(XX, "Pj Rfnn")
Erase XX
End Function

Function DoPjRfzStd() As Drs
Erase XX
X "MVb   Scripting VBScript_RegExp_55 DAO VBIDE Office"
X "MIde  Scripting VBIDE Excel"
X "MXls  Scripting Office Excel"
X "MDao  Scripting DAO"
X "MAdo  Scripting ADODB"
X "MAdoX Scripting ADOX"
X "MApp  Scripting"
X "MDta  Scripting"
X "MTp   Scripting"
X "MSql  Scripting"
X "AStkShpCst Scripting"
X "MAcs  Scripting Office Access"
DoPjRfzStd = DrszTRstLy(XX, "Pj Rfnn")
End Function

Function DoStdLib() As Drs
Erase XX
X "VBA                {000204EF-0000-0000-C000-000000000046} 4  2 C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL"
X "Access             {4AFFC9A0-5F99-101B-AF4E-00AA003F0F07} 9  0 C:\Program Files (x86)\Microsoft Office\Root\Office16\MSACC.OLB"
X "stdole             {00020430-0000-0000-C000-000000000046} 2  0 C:\Windows\SysWOW64\stdole2.tlb"
X "Excel              {00020813-0000-0000-C000-000000000046} 1  9 C:\Program Files (x86)\Microsoft Office\Root\Office16\EXCEL.EXE"
X "Scripting          {420B2830-E718-11CF-893D-00A0C9054228} 1  0 C:\Windows\SysWOW64\scrrun.dll"
X "DAO                {4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28} 12 0 C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL"
X "Office             {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52} 2  8 C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\MSO.DLL"
X "ADODB              {B691E011-1797-432E-907A-4D8C69339129} 6  1 C:\Program Files (x86)\Common Files\System\ado\msado15.dll"
X "ADOX               {00000600-0000-0010-8000-00AA006D2EA4} 6  0 C:\Program Files (x86)\Common Files\System\ado\msadox.dll"
X "VBScript_RegExp_55 {3F4DACA7-160D-11D2-A8E9-00104B365C9F} 5  5 C:\Windows\SysWOW64\vbscript.dll"
X "VBIDE              {0002E157-0000-0000-C000-000000000046} 5  3 C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
DoStdLib = Drsz4TRstLy(XX, "Libn Guid Maj Mnr Ffn")
Erase XX
End Function

Sub BrwDoPjRfzStd()
BrwDrs DoPjRfzStd
End Sub
