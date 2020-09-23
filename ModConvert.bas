Attribute VB_Name = "ModConvert"
Option Compare Text

'Most used bases
Public Const B_BIN As Integer = 2
Public Const B_OCT As Integer = 8
Public Const B_DEC As Integer = 10
Public Const B_HEX As Integer = 16
'Some separators
Public Const DEFAULT_SEPARATOR As String = "."
Public Const COMMA_SEPARATOR As String = ","

Private Digits(0 To 35) As String 'Bases digits
Private INum() As Integer 'Input number
Private ONum() As Integer 'Output number
Private IBase As Integer 'Input base
Private OBase As Integer 'Output base

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VeryLongConvert : function that converts a huge number as string from a base to
'                  another one
'
'                  Version : 1.01
'                  Author:  Guillaume GIFFARD
'                  Date : 01/03/2002
'                  Mail : Guiland@mail.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INPUTS :  * Word As String : the huge number to convert
'
'          * FromBase As Integer : the base in witch Word is written
'
'          * ToBase As Integer : the base in witch Word is to convert
'
'          * Separator As String : this Optional variable is the decimal separator,
'          usely the point
'
'          FromBase and ToBase are integers from 2 to 36
'
'OUTPUTS : * the function returns the huge number value converted from FromBase to
'          ToBase as string. It returns "" if Word is empty or if FromBase or
'          ToBase is not between 2 and 36
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function VeryLongConvert(Word As String, FromBase As Integer, ToBase As Integer, Optional Separator As String = DEFAULT_SEPARATOR) As String
    If Word = "" Or FromBase < 2 Or FromBase > 36 Or ToBase < 2 Or ToBase > 36 Then Exit Function
    If Digits(35) <> "Z" Then InitDigits
    Call StringToArray(Word, FromBase, Separator)
    Convert (ToBase)
    VeryLongConvert = DeleteZeros(ArrayToString(Separator), Separator)
End Function

'Saves the bases digits in an array
Private Sub InitDigits()
    For i = 0 To 9
        Digits(i) = i
    Next i
    For i = 10 To 35
        Digits(i) = Chr(i + 55)
    Next i
End Sub
 
'Saves a number as string in an array of integers
'Each cell of the array in one digit of the number
Private Sub StringToArray(Word As String, Base As Integer, Optional Separator As String = DEFAULT_SEPARATOR)
    If Word = "" Or Base < 2 Or Base > 36 Then Exit Sub
    Dim Point As Integer, Min As Integer, Max As Integer, NoPoint As Integer
    IBase = Base
    Point = InStr(1, Word, Separator, vbTextCompare)
    If Point = 0 Then
        Max = Len(Word) - 1
        Min = 0
    Else
        Max = Point - 2
        Min = Point - Len(Word) + Len(Separator) - 1
    End If
    ReDim INum(Min To Max)
    For i = 0 To Len(Word) - 1
        If i <= Len(Word) - Point And i >= Len(Word) - Point - Len(Separator) + 1 Then
            NoPoint = NoPoint - 1
        Else
            INum(Min + i + NoPoint) = Number(Left(Right(Word, i + 1), 1), IBase)
        End If
    Next i
End Sub

'Returns the number corresponding to a digit as string if the digit is allowed
'by the base. e.g. : C is allowed in hexadecimal but not in decimal or in octal
Private Function Number(Digit As String, Base As Integer) As Integer
    If Digit = "" Or Base < 2 Or Base > 36 Then Exit Function
    For i = 0 To 35
        If i = Base Then Exit Function
        If UCase(Digit) = Digits(i) Then Number = i
    Next i
End Function

'THE sub that converts INum to ONum with IBase and OBase
Private Sub Convert(Base As Integer)
    If Base < 2 Or Base > 36 Then Exit Sub
    Dim Max As Integer, Min As Integer
    Dim TmpNum() As Integer, Tmp2Num() As Integer

    OBase = Base
    Max = RoundOverInt(Int((UBound(INum, 1) + 1) * Log(IBase) / Log(OBase)))
    Min = Int(LBound(INum, 1) * Log(IBase) / Log(OBase))
    ReDim ONum(Min To Max)
    i = 0 'LBound(ONum, 1)
    Call DivideVeryLong(INum, OBase, TmpNum, ONum(i), IBase)
    i = i + 1
    Do Until i > UBound(ONum, 1)
        Call DivideVeryLong(TmpNum, OBase, Tmp2Num, ONum(i), IBase)
        ReDim TmpNum(LBound(Tmp2Num, 1) To UBound(Tmp2Num, 1))
        For j = LBound(Tmp2Num, 1) To UBound(Tmp2Num, 1)
            TmpNum(j) = Tmp2Num(j)
        Next j
        i = i + 1
    Loop
End Sub

'round numbers to the closest higher integer
'e.g. : 3.9 gives 4 ; 3.4 gives 4 ; 3 gives 3
Private Function RoundOverInt(Value As Double) As Double
    If Value = Int(Value) Then RoundOverInt = Value Else RoundOverInt = Int(Value) + 1
End Function

'Divides a huge number by an integer and returns the huge quotient and the remainder
Private Sub DivideVeryLong(Numerator() As Integer, Denominator As Integer, QuotientOut() As Integer, Remainder As Integer, Base As Integer)
    Dim Tmp As Long, Decal As Long
    ReDim QuotientOut(LBound(Numerator, 1) To UBound(Numerator, 1))
    Tmp = 0
    Decal = 0
    For i = UBound(Numerator, 1) To 0 Step -1 'LBound(Numerator, 1) Step -1
        Tmp = Tmp * Base + Numerator(i)
        QuotientOut(i - Decal) = Tmp \ Denominator
        'If QuotientOut(i - Decal) = 0 And Decal = i - 1 Then
        '    Decal = Decal - 1
        '    ReDim QuotientOut(LBound(Numerator, 1) To UBound(Numerator, 1) - Decal)
        'End If
        Tmp = Tmp - QuotientOut(i - Decal) * Denominator
    Next i
    Remainder = Tmp
End Sub

'Saves an array in a string
Private Function ArrayToString(Optional Separator As String = DEFAULT_SEPARATOR) As String
    For i = UBound(ONum, 1) To LBound(ONum, 1) Step -1
        If i = -1 Then ArrayToString = ArrayToString & Separator
        ArrayToString = ArrayToString & Digits(ONum(i))
    Next i
End Function

'Deletes zeros before and after the number as string and, if possible, deletes the
'separator
Private Function DeleteZeros(Word As String, Optional Separator As String = DEFAULT_SEPARATOR) As String
    Dim Point As Integer, WordTmp As String
    WordTmp = Word
    Do
        Point = InStr(1, WordTmp, "0", vbTextCompare)
        If Point = 1 Then WordTmp = Right(WordTmp, Len(WordTmp) - 1) Else Exit Do
    Loop
    If InStr(1, WordTmp, Separator, vbTextCompare) <> 0 Then
        Do
            Point = InStr(Len(WordTmp) - 1, WordTmp, "0", vbTextCompare)
            If Point = Len(WordTmp) - 1 Then WordTmp = Left(WordTmp, Len(WordTmp) - 1) Else Exit Do
        Loop
        Do
            Point = InStr(Len(WordTmp) - 1, WordTmp, "0", vbTextCompare)
            If Point = Len(WordTmp) Then WordTmp = Left(WordTmp, Len(WordTmp) - 1) Else Exit Do
        Loop
    End If
    If WordTmp = "" Then WordTmp = "0"
    If InStr(1, WordTmp, Separator, vbTextCompare) = Len(WordTmp) - Len(Separator) + 1 Then WordTmp = Left(WordTmp, Len(WordTmp) - Len(Separator))
    DeleteZeros = WordTmp
End Function
