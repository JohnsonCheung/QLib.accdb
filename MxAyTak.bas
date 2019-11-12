Attribute VB_Name = "MxAyTak"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAyTak."

Function SyTakBefDD(Sy$()) As String()
Dim I: For Each I In Itr(Sy)
    PushI SyTakBefDD, BefDD(I)
Next
End Function

Function SyTakAftDot(Sy$()) As String()
SyTakAftDot = SyTakAft(Sy, ".")
End Function

Function SyTakAft(Sy$(), Sep$) As String()
Dim I: For Each I In Itr(Sy)
    PushI SyTakAft, Aft(I, Sep)
Next
End Function

Function SyTakAftOrAll(Sy$(), Sep$) As String()
Dim I: For Each I In Itr(Sy)
    PushI SyTakAftOrAll, AftOrAll(I, Sep)
Next
End Function

Function SyTakBef(Sy$(), Sep$) As String() 'Return a Sy which is taking Bef-Sep from Given Sy
Dim I
For Each I In Itr(Sy)
    PushI SyTakBef, Bef(CStr(I), Sep)
Next
End Function

Function SyTakBefDot(Sy$()) As String()
SyTakBefDot = SyTakBef(Sy, ".")
End Function

Function SyTakBefOrAll(Sy$(), Sep$) As String()
Dim I
For Each I In Itr(Sy)
    Push SyTakBefOrAll, BefOrAll(CStr(I), Sep)
Next
End Function

Function BetBktzAy(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI BetBktzAy, BetBkt(CStr(I))
Next
End Function
