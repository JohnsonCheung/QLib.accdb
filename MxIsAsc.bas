Attribute VB_Name = "MxIsAsc"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxIsAsc."
Const AscPlus% = &H2B  ' + sign
Const AscMinus% = &H2D  ' - sign
Const AscLDash% = 95 '_

Function IsAscDig(A%) As Boolean
IsAscDig = &H30 <= A And A <= &H39
End Function

Function IsAscSgn(A%) As Boolean
If A = AscPlus Then IsAscSgn = True: Exit Function
If A = AscMinus Then IsAscSgn = True: Exit Function
End Function

Function IsAscDigSgn(A%) As Boolean
Select Case True
Case IsAscDig(A), IsAscSgn(A): IsAscDigSgn = True
End Select
End Function

Property Get AscAyzNonPrt() As Integer()
Dim J%
For J = 0 To 255
    If IsAscNonPrt(J) Then PushI AscAyzNonPrt, J
Next
End Property

Function IsAscPrintablezStrI(S, I) As Boolean
IsAscPrintablezStrI = IsAscPrintable(Asc(Mid(S, I, 1)))
End Function

Function IsAscNonPrt(A%) As Boolean
IsAscNonPrt = Not IsAscPrintable(A)
End Function

Function IsAscPrintable(A%) As Boolean
Select Case A
Case 0, 1, 9, 10, 13, 28, 29, 30, 31, 129, 141, 143, 144, 157, 160
Case Else: IsAscPrintable = True
End Select
End Function

Function IsAscDigit(A%) As Boolean
If A < 48 Then Exit Function
If A > 57 Then Exit Function
IsAscDigit = True
End Function

Function IsAscFstNmChr(A%) As Boolean
IsAscFstNmChr = IsAscLetter(A)
End Function

Function IsAscLDash(A%) As Boolean
IsAscLDash = A = 95
End Function

Function IsAscLCas(A%) As Boolean
If A < 97 Then Exit Function
If A > 122 Then Exit Function
IsAscLCas = True
End Function

Function IsAscLetterDig(A%) As Boolean
IsAscLetterDig = True
If IsAscLetter(A) Then Exit Function
If IsAscDig(A) Then Exit Function
IsAscLetterDig = False
End Function

Function IsAscLetter(A%) As Boolean
IsAscLetter = True
If IsAscUCas(A) Then Exit Function
If IsAscLCas(A) Then Exit Function
IsAscLetter = False
End Function

Function IsAscNmChr(A%) As Boolean
IsAscNmChr = True
If IsAscLetter(A) Then Exit Function
If IsAscDig(A) Then Exit Function
IsAscNmChr = A = AscLDash
End Function

Function IsAscPun(A%) As Boolean
'  0 1 2 3 4 5 6 7 8 9 A B C D E F
'0                
'1                
'2   ! " # $ % & ' ( ) * + , - . /
'3 0 1 2 3 4 5 6 7 8 9 : ; < = > ?
'4 @ A B C D E F G H I J K L M N O
'5 P Q R S T U V W X Y Z [ \ ] ^ _
'6 ` a b c d e f g h i j k l m n o
'7 p q r s t u v w x y z { | } ~ 
Select Case True
Case IsAscPun1(A), IsAscPun2(A), IsAscPun3(A), IsAscPun4(A): IsAscPun = True
End Select
End Function

Function IsAscPun1(A%) As Boolean
IsAscPun1 = (&H21 <= A And A <= &H2F)
End Function

Function IsAscPun2(A%) As Boolean
IsAscPun2 = (&H3A <= A And A <= &H40)
End Function

Function IsAscPun3(A%) As Boolean
IsAscPun3 = (&H5B <= A And A <= &H60)
End Function

Function IsAscPun4(A%) As Boolean
IsAscPun4 = (&H7B <= A And A <= &H7F)
End Function

Function IsAscUCas(A%) As Boolean
If A < 65 Then Exit Function
If A > 90 Then Exit Function
IsAscUCas = True
End Function

Function AscAt%(S, At)
AscAt = Asc(Mid(S, At, 1))
End Function

Function IsStrAtSpcCrLf(S, At) As Boolean
IsStrAtSpcCrLf = IsAscSpcCrLf(AscAt(S, At))
End Function

Function IsAscSpcCrLf(Asc%)
Select Case True
Case Asc = 13, Asc = 10, Asc = 32: IsAscSpcCrLf = True
End Select
End Function
