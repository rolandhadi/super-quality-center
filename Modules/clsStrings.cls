VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************************************************************************************************
'*      LIBRARY NAME: clsStringAndNumericOperations.cls
'*      PURPOSE: Compilation of functions that deals with Strings and Numbers
'*      AUTHOR: Roland Ross Hadi
'*      CREATION DATE: April 07, 2009
'*      NOTES:
'*
'*******************************************************************************************************************************************************************************************
'*      MODIFICATION HISTORY
'*      NAME                                            DATE OF CHANGE                                          FUNCTION NAME
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*******************************************************************************************************************************************************************************************

'*******************************************************************************************************************************************************************************************
'*      GENERAL DECLARATION
        Option Explicit
        Private Const CHARSET_LETTERS_UCASE = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Private Const CHARSET_LETTERS_LCASE = "abcdefghijklmnopqrstuvwxyz"
        Private Const CHARSET_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
        Private Const CHARSET_DIGITS = "0123456789"
        Private Const CHARSET_ALPHANUMERIC = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
        Private Const CHARSET_PRINTABLE = " !""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
        Private Const CHARSET_WHITESPACE = vbCr & vbLf & vbTab & " "
        
        Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
        Private Const EM_GETSEL = &HB0
'*******************************************************************************************************************************************************************************************


' @Function StrToArrayStr
' -----------------------------
'@Author Roland Ross Hadi
'@Description Converts string to array
'@Comments
' Parameter:
'       Array_ - Array of string
Public Function StrToArrayStr(ParamArray Array_() As Variant) As String
    Dim str1 As String
    Dim zt As Long
    Dim A
    
    str1 = nts4T(UBound(Array_) + 1)
    zt = 1
    For A = 0 To UBound(Array_)
        str1 = str1 & nts4T(zt)
        zt = zt + Len(Array_(A))
    Next
    str1 = str1 & nts4T(zt)
    For A = 0 To UBound(Array_)
        str1 = str1 & Array_(A)
    Next
    StrToArrayStr = str1
End Function
' Function StrToArrayStr
' -----------------------------

' @Function StrFromArrayStr
' -----------------------------
'@Author Roland Ross Hadi
'@Description Converts string to array
'@Comments
' Parameter:
'       Arr - Array of string
'       ArrCnt - Array number
Public Function StrFromArrayStr(Arr As String, ArrCnt As Long) As String
    Dim strs As Long
    Dim data As String
    Dim bgn, ent, ln
    strs = nts4F(Mid$(Arr, 1, 4))
    If ArrCnt > strs Then Exit Function
    data = Mid$(Arr, strs * 4 + 9)
    bgn = nts4F(Mid$(Arr, ArrCnt * 4 + 1, 4))
    ent = nts4F(Mid$(Arr, ArrCnt * 4 + 5, 4))
    ln = ent - bgn
    StrFromArrayStr = Mid$(data, bgn, ln)
End Function
' Function StrFromArrayStr
' -----------------------------

' @Function ArrayStrLen
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the array length
'@Comments
' Parameter:
'       ArrStr - Array of string
Public Function ArrayStrLen(ArrStr As String) As Long
    ArrayStrLen = nts4F(Mid$(ArrStr, 1, 4))
End Function
' Function ArrayStrLen
' -----------------------------

' @Function StrByte
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the string bytes
'@Comments
' Parameter:
'       txt - Array of string
'       a -
Public Function StrByte(txt As String, A As Long) As Variant
    StrByte = Mid$(txt, A, 1)
End Function
' Function StrByte
' -----------------------------

' @Function StrCompare
' -----------------------------
'@Author Roland Ross Hadi
'@Description Compares two strings
'@Comments
' Parameter:
'       str1 - string to be compared
'       str2 - string to be compared to
Public Function StrCompare(str1 As String, str2 As String) As Long
    Dim t As Long
    Dim tol, A
    tol = (Len(str1) + Len(str2)) / 2
    If Len(str2) > Len(str1) Then
        t = t + Len(str2) - Len(str1)
        str2 = Mid$(str2, 1, Len(str1))
    End If
    If Len(str1) > Len(str2) Then
        t = t + Len(str1) - Len(str2)
        str1 = Mid$(str1, 1, Len(str2))
    End If
    For A = 1 To Len(str1)
        If Mid$(str1, A, 1) <> Mid$(str2, A, 1) Then
            If Asc(Mid$(str1, A, 1)) > Asc(Mid$(str2, A, 1)) Then
                t = t + Asc(Mid$(str1, A, 1)) - Asc(Mid$(str2, A, 1))
            Else
                t = t + Asc(Mid$(str2, A, 1)) - Asc(Mid$(str1, A, 1))
            End If
        End If
        DoEvents
    Next
    StrCompare = t / tol
End Function
' Function StrCompare
' -----------------------------

' @Function StrToMix
' -----------------------------
'@Author Roland Ross Hadi
'@Description StrToMix
'@Comments
' Parameter:
'       strText1 - string to be compared
'       strKey - string to be compared to
Public Function StrToMix(strText1 As String, strKey)
    Dim txt2, A, dat
    txt2 = strKey
    If Len(txt2) > Len(strText1) Then txt2 = Mid$(txt2, 1, Len(strText1))
    If Len(strText1) > Len(txt2) Then strText1 = Mid$(strText1, 1, Len(txt2))
    For A = 1 To Len(strText1)
        dat = dat & Chr$(IIf(Asc(Mid$(strText1, A, 1)) + Asc(Mid$(txt2, A, 1)) >= 256, Asc(Mid$(strText1, A, 1)) + Asc(Mid$(txt2, A, 1)) - 255, Asc(Mid$(strText1, A, 1)) + Asc(Mid$(txt2, A, 1))))
        DoEvents
    Next
    StrToMix = dat
End Function
' Function StrToMix
' -----------------------------

' @Function StrFromMix
' -----------------------------
'@Author Roland Ross Hadi
'@Description StrFromMix
'@Comments
' Parameter:
'       strText1 - string to be compared
'       strKey - string to be compared to
Public Function StrFromMix(strText1, strKey)
    Dim txt2, A, dat
    txt2 = strKey
    If Len(txt2) > Len(strText1) Then txt2 = Mid$(txt2, 1, Len(strText1))
    If Len(strText1) > Len(txt2) Then txt2 = Mid$(strText1, 1, Len(txt2))
    For A = 1 To Len(strText1)
        dat = dat & Chr$(IIf(Asc(Mid$(strText1, A, 1)) - Asc(Mid$(txt2, A, 1)) < 0, Asc(Mid$(strText1, A, 1)) - Asc(Mid$(txt2, A, 1)) + 255, Asc(Mid$(strText1, A, 1)) - Asc(Mid$(txt2, A, 1))))
        DoEvents
    Next
    StrFromMix = dat
End Function
' Function StrFromMix
' -----------------------------

' @Function StrToURL
' -----------------------------
'@Author Roland Ross Hadi
'@Description Set string <StringToEncode> to URL format
'@Comments
' Parameter:
'       StringToEncode - string to be encoded
Public Function StrToURL(StringToEncode As String) As String
    Dim TempAns As String
    Dim CurChr As Integer
    CurChr = 1
    Do Until CurChr - 1 = Len(StringToEncode)
        Select Case Asc(Mid(StringToEncode, CurChr, 1))
        Case 48 To 57, 65 To 90, 97 To 122
            TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
        Case 32
            TempAns = TempAns & "%" & Hex(32)
        Case Else
            TempAns = TempAns & "%" & _
            Format(Hex(Asc(Mid(StringToEncode, _
            CurChr, 1))), "00")
        End Select
        CurChr = CurChr + 1
    Loop
    StrToURL = TempAns
End Function
' Function StrToURL
' -----------------------------

' @Function StrFromURL
' -----------------------------
'@Author Roland Ross Hadi
'@Description Decode string <StringToDecode> from URL format
'@Comments
' Parameter:
'       StringToEncode - string to be decoded
Public Function StrFromURL(StringToDecode As String) As String
    Dim TempAns As String
    Dim CurChr As Integer
    CurChr = 1
    Do Until CurChr - 1 = Len(StringToDecode)
        Select Case Mid(StringToDecode, CurChr, 1)
        Case "+"
            TempAns = TempAns & " "
        Case "%"
            TempAns = TempAns & Chr(Val("&h" & _
            Mid(StringToDecode, CurChr + 1, 2)))
            CurChr = CurChr + 2
        Case Else
            TempAns = TempAns & Mid(StringToDecode, CurChr, 1)
        End Select
        CurChr = CurChr + 1
    Loop
    StrFromURL = TempAns
End Function
' Function StrFromURL
' -----------------------------


' @Function StrSetLength
' -----------------------------
'@Author Roland Ross Hadi
'@Description Set <strText> with placeholders <strPlaceholder>
'@Comments
' Parameter:
'       strText - string to be processed
'       intLength - desired length of the string
'       strPlaceholder - string to be the placeholder
'       intSide - Side to process 1 - Left 2 = Right
Public Function StrSetLength(strText As String, intLength As Integer, Optional strPlaceholder As String = " ", Optional intSide As Integer = 2) As String
    Dim dat As String
    dat = strText
    Do
        If Len(strText) > intLength Then
            If intSide = 1 Then
                dat = StrLeft(dat, intLength)
            ElseIf intSide = 2 Then
                dat = StrRight(dat, intLength)
            End If
        ElseIf Len(strText) < intLength Then
            If intSide = 1 Then
                dat = strPlaceholder & dat
            ElseIf intSide = 2 Then
                dat = dat & strPlaceholder
            End If
        End If
        DoEvents
    Loop Until Len(dat) = intLength
    StrSetLength = dat
End Function
' Function StrSetLength
' -----------------------------

' @Function StrIn
' -----------------------------
'@Author Roland Ross Hadi
'@Description Checks if a word <strSearchText> is present in a string <strText>
'@Comments
' Parameter:
'       strText - string to processed
'       strSearchText - search string
Public Function StrIn(strText, strSearchText) As Boolean
    If InStr(strText, strSearchText) <> 0 Then StrIn = True
End Function
' Function StrIn
' -----------------------------

' @Function StrReplace
' -----------------------------
'@Author Roland Ross Hadi
'@Description Replace <strText> with <arrString()>
'@Comments
' Parameter:
'       strText - string to  processed
'       arrString() - search string
Public Function StrReplace(strText, ParamArray arrString() As Variant)
    Dim dat As String
    Dim A
    dat = strText
    For A = 0 To UBound(arrString) Step 2
        dat = Replace$(dat, CStr(arrString(A)), CStr(arrString(A + 1)))
    Next
    StrReplace = dat
End Function
' Function StrReplace
' -----------------------------

' @Function StrLeft
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the left part of the string
'@Comments
' Parameter:
'       strText - string to  processed
'       a - left count
'       b - left count with len()
Public Function StrLeft(strText, Optional A = -1, Optional B = -1)
    If A <> -1 Then
        StrLeft = Left$(strText, A)
    ElseIf B <> -1 Then
        StrLeft = Left$(strText, Len(strText) - Len(B))
    End If
End Function
' Function StrLeft
' -----------------------------

' @Function StrRight
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the right part of the string
'@Comments
' Parameter:
'       strText - string to  processed
'       a - left count
'       b - left count with len()
Public Function StrRight(strText, Optional A = -1, Optional B = -1)
    If A <> -1 Then
        StrRight = Right$(strText, A)
    ElseIf B <> -1 Then
        StrRight = Right$(strText, Len(strText) - Len(B))
    End If
End Function
' Function StrRight
' -----------------------------

' @Function StrTo
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the part of the string <strText> up to <strUpTo>
'@Comments
' Parameter:
'       strText - string to  processed
'       strUpTo - search up to this word
Public Function StrTo_(strText As String, strUpTo As String) As String
    Dim A, c
    For A = 1 To Len(strText)
        If Mid$(strText, A, Len(strUpTo)) = strUpTo Then Exit For
        c = c & Mid$(strText, A, 1)
    Next
    StrTo_ = c
End Function
' Function StrTo
' -----------------------------

' @Function strFrom_
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the part of the string <strText> from <strFrom>
'@Comments
' Parameter:
'       strText - string to  processed
'       strFrom - search up to this word
Public Function strFrom_(strText As String, strFrom As String) As String
    strFrom_ = Mid$(strText, InStr(strText, strFrom) + Len(strFrom))
End Function
' Function StrFrom
' -----------------------------

' @Function StrBetween
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the part of the string <strText> from <strFrom> to <strTo>
'@Comments
' Parameter:
'       strText - string to  processed
'       strFrom - search from to this word
'       strTo - search up to this word
Public Function StrBetween(strText As String, strFrom As String, StrTo As String) As String
    Dim dat As String
    dat = strFrom_(strText, strFrom)
    StrBetween = StrTo_(dat, StrTo)
End Function
' Function StrBetween
' -----------------------------

' @Function StrBetweenWithLoc
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the part of the string <strText> from <strFrom> to <strTo>
'@Comments
' Parameter:
'       strText - string to  processed
'       strFrom - search from to this word
'       strTo - search up to this word
Public Function StrBetweenWithLoc(strText As String, strFrom As String, StrTo As String, StartAt) As String
    Dim dat As String
    dat = strFrom_(Right(strText, Len(strText) - StartAt), strFrom)
    StrBetweenWithLoc = StrTo_(dat, StrTo)
End Function
' Function StrBetween
' -----------------------------

' @Function StrToHex
' -----------------------------
'@Author Roland Ross Hadi
'@Description Converts <strText> to Hex
'@Comments
' Parameter:
'       strText - string to  processed
'       st - search from to this word
Public Function StrToHex(strText, Optional st = " ")
    Dim A, B
    For A = 1 To Len(strText)
        B = B & StrSetLength(Hex(Asc(Mid$(strText, A, 1))), 2, "0", 1) & st
        DoEvents
    Next
    StrToHex = Mid$(B, 1, Len(B) - Len(st))
End Function
' Function StrToHex
' -----------------------------

' @Function StrFromHex
' -----------------------------
'@Author Roland Ross Hadi
'@Description Converts <strText> from Hex
'@Comments
' Parameter:
'       strText - string to  processed
'       st - search from to this word
Public Function StrFromHex(strText, Optional st = " ")
    Dim G, A, B
    G = Split(strText, st)
    For A = 0 To UBound(G)
        B = B & Chr$(XHexToDecimall(G(A)))
    Next
    StrFromHex = B
End Function
' Function StrFromHex
' -----------------------------

' @Function XHexToDecimall
' -----------------------------
'@Author Roland Ross Hadi
'@Description Converts <num> to Decimal
'@Comments
' Parameter:
'       num - string to  processed
Private Function XHexToDecimall(num As Variant)
    Dim A, zh, G, gh, numm, gg
    For A = 1 To Len(num)
        If Mid$(num, A, 1) <> "0" Then
            Exit For
        Else
            zh = True
        End If
    Next
    If zh = True Then num = Mid$(num, A)
    num = UCase$(num)
    Dim nums(13) As Currency
    nums(1) = 1
    nums(2) = 16
    For A = 3 To 13
        nums(A) = nums(A - 1) * 16
    Next
    For A = Len(num) To 1 Step -1
        G = G + Mid$(num, A, 1)
    Next
    num = G
    For A = 1 To Len(num)
        gh = Mid$(num, A, 1)
        If gh = "0" Then numm = 0
        If gh = "1" Then numm = 1
        If gh = "2" Then numm = 2
        If gh = "3" Then numm = 3
        If gh = "4" Then numm = 4
        If gh = "5" Then numm = 5
        If gh = "6" Then numm = 6
        If gh = "7" Then numm = 7
        If gh = "8" Then numm = 8
        If gh = "9" Then numm = 9
        If gh = "A" Then numm = 10
        If gh = "B" Then numm = 11
        If gh = "C" Then numm = 12
        If gh = "D" Then numm = 13
        If gh = "E" Then numm = 14
        If gh = "F" Then numm = 15
        numm = numm * nums(A)
        gg = gg + numm
    Next
    XHexToDecimall = gg
End Function
' Function XHexToDecimall
' -----------------------------

' @Function nts4T
' -----------------------------
'@Author Roland Ross Hadi
'@Description Converts <num> to nts4T
'@Comments
' Parameter:
'       num - string to  processed
Private Function nts4T(num As Long) As String
    Dim hlen As String
    hlen = Format(Hex(num), "00000000")
    If Len(hlen) < 8 Then hlen = String(8 - Len(hlen), "0") & hlen
    nts4T = Chr$(XHexToDecimall(Mid$(hlen, 1, 2))) & Chr$(XHexToDecimall(Mid$(hlen, 3, 2))) & Chr$(XHexToDecimall(Mid$(hlen, 5, 2))) & Chr$(XHexToDecimall(Mid$(hlen, 7, 2)))
End Function
' Function nts4T
' -----------------------------

' @Function nts4F
' -----------------------------
'@Author Roland Ross Hadi
'@Description Converts <num> to nts4F
'@Comments
' Parameter:
'       num - string to  processed
Private Function nts4F(num As String) As Long
    If Len(num) <> 4 Then Exit Function
    num = StrSetLength(Hex(Asc(Mid$(num, 1, 1))), 2, "0", 1) & StrSetLength(Hex(Asc(Mid$(num, 2, 1))), 2, "0", 1) & StrSetLength(Hex(Asc(Mid$(num, 3, 1))), 2, "0", 1) & StrSetLength(Hex(Asc(Mid$(num, 4, 1))), 2, "0", 1)
    nts4F = XHexToDecimall(CStr(num))
End Function
' Function nts4F
' -----------------------------

' @Function LeftPad
' -----------------------------
'@Author Roland Ross Hadi
'@Description Pad the left side of the string with enough PadCharacter characters until the string is intSize characters long.
'@Comments
' Parameter:
'       strText - string to  processed
'       intSize - intSize of the whole string
'       PadCharacter - string to  processed
Public Function LeftPad(strText, intSize As Long, Optional PadCharacter As String = " ") As String
    LeftPad = "" & strText
    While Len(LeftPad) < intSize
        LeftPad = PadCharacter & LeftPad
    Wend
End Function
' Function LeftPad
' -----------------------------

' @Function RightPad
' -----------------------------
'@Author Roland Ross Hadi
'@Description Pad the right side of the string with enough PadCharacter characters until the string is intSize characters long.
'@Comments
' Parameter:
'       strText - string to  processed
'       intSize - intSize of the whole string
'       PadCharacter - string to  processed
Public Function RightPad(strText, intSize As Long, Optional PadCharacter As String = " ") As String
    RightPad = "" & strText
    While Len(RightPad) < intSize
        RightPad = RightPad & PadCharacter
    Wend
End Function
' Function RightPad
' -----------------------------

' @Function AreLetters
' -----------------------------
'@Author Roland Ross Hadi
'@Description Are all of the characters in strText letters (i.e., A - Z and a - z)?
'@Comments
' Parameter:
'       strText - string to  processed
Public Function AreLetters(strText As String) As Boolean
    Dim i As Long, lLen As Long, nChar As Integer
    lLen = Len(strText)
    For i = 1 To lLen
        nChar = Asc(UCase(Mid(strText, i, 1)))
        If nChar < 65 Or nChar > 90 Then Exit Function
    Next
    AreLetters = True
End Function
' Function AreLetters
' -----------------------------

' @Function AreDigits
' -----------------------------
'@Author Roland Ross Hadi
'@Description Are all of the characters in strText digits (i.e., 0 - 9)?
'@Comments
' Parameter:
'       strText - string to  processed
Public Function AreDigits(strText As String) As Boolean
    Dim i As Long, lLen As Long, nChar As Integer
    lLen = Len(strText)
    For i = 1 To lLen
        nChar = Asc(UCase(Mid(strText, i, 1)))
        If nChar < 48 Or nChar > 57 Then Exit Function
    Next
    AreDigits = True
End Function
' Function AreDigits
' -----------------------------

' @Function AreLettersOrDigits
' -----------------------------
'@Author Roland Ross Hadi
'@Description Are all of the characters in strText letters or digits?
'@Comments
' Parameter:
'       strText - string to  processed
Public Function AreLettersOrDigits(strText As String) As Boolean
    Dim i As Long, lLen As Long, nChar As Integer
    lLen = Len(strText)
    For i = 1 To lLen
        nChar = Asc(UCase(Mid(strText, i, 1)))
        If (nChar < 65 Or nChar > 90) And (nChar < 48 Or nChar > 57) Then Exit Function
    Next
    AreLettersOrDigits = True
End Function
' Function AreLettersOrDigits
' -----------------------------

' @Function AreInSet
' -----------------------------
'@Author Roland Ross Hadi
'@Description Are all of the characters in strText within the specified set?
'               An example of a set is "0123456789 .,Ee-", which should suffice to
'               recognize most simple number formats and even scientific notation.
'@Comments
' Parameter:
'       strText - string to  processed
'       CharacterSet - string to  processed
'       CaseSensitive - string to  processed
Public Function AreInSet(strText As String, CharacterSet As String, Optional CaseSensitive As Boolean = True) As Boolean
    Dim i As Long, lLen As Long
    lLen = Len(strText)
    For i = 1 To Len(strText)
        If InStr(1, CharacterSet, Mid(strText, i, 1), CaseSensitive + 1) = 0 Then Exit Function
    Next
    AreInSet = True
End Function
' Function AreInSet
' -----------------------------

' @Function AltCaps
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will make the caps in text go on and off for each letter, like this:  cOoL
'@Comments
' Parameter:
'       strText - string to  processed
Function AltCaps(strText As String)
        On Error GoTo error
        Dim i As Integer
        Dim s As String
        Dim keyval
        s = ""
        For i = 1 To Len(strText$)
          keyval = Asc(Mid$(strText$, i, 1))
          If (keyval >= 96 And keyval < 96 + 26) Or (keyval >= 64 And keyval < 64 + 26) Then
            If (i And 1) = 1 Then
              If keyval < 96 Then
                s = s + Chr$(96 + keyval - 64)
              Else
                s = s + Chr$(keyval)
              End If
            Else
              If keyval >= 96 Then
                s = s + Chr$(64 + keyval - 96)
              Else
                s = s + Chr$(keyval)
              End If
            End If
          Else
            s = s + Chr$(keyval)
          End If
        Next i
        strText$ = s
        AltCaps = strText$
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
        End Function

' @Function BackwardsText
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will make text go backwards, like this:  looC (Cool)
'@Comments
' Parameter:
'       strText - string to  processed
Function BackwardsText(strText As String)
        Dim i, stringy$, final$
        On Error GoTo error
        For i = 1 To Len(strText$)
        stringy = Mid$(strText$, i, 1)
        final = stringy$ + final$
        Next i
        BackwardsText = final$
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function BackwardsText
' -----------------------------

' @Function EliteType
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will change characters to make them "elite", example:  ©00|_
'@Comments
' Parameter:
'       strText - string to  processed
Function EliteType(strText As String)
        Dim s(52)
        On Error GoTo error
        s(0) = "æ"
        s(1) = "å"
        s(2) = "b"
        s(3) = "<"
        s(4) = "c|"
        s(5) = "ë"
        s(6) = "f"
        s(7) = "9"
        s(8) = "h"
        s(9) = "ï"
        s(10) = "j"
        s(11) = "|<"
        s(12) = "|_"
        s(13) = "/x\"
        s(14) = "|\|"
        s(15) = "0"
        s(16) = "p"
        s(17) = "q"
        s(18) = "r"
        s(19) = "_/¯"
        s(20) = "-|-"
        s(21) = "µ"
        s(22) = "\/"
        s(23) = "\/\/"
        s(24) = "×"
        s(25) = "ÿ"
        s(26) = "¯/_"
        s(27) = "Ä"
        s(28) = "ß"
        s(29) = "©"
        s(30) = "|}"
        s(31) = "È"
        s(32) = "F"
        s(33) = "G"
        s(34) = "|-|"
        s(35) = "I"
        s(36) = "J"
        s(37) = "]<"
        s(38) = "]_"
        s(39) = "/\/\"
        s(40) = "|\|"
        s(41) = "{}"
        s(42) = "P"
        s(43) = "¶"
        s(44) = "|2"
        s(45) = "§"
        s(46) = "¯|¯"
        s(47) = "|_|"
        s(48) = "\/"
        s(49) = "\x/"
        s(50) = "><"
        s(51) = "¥"
        s(52) = "¯/_"
        strText$ = Replace(strText$, "a", s(1))
        strText$ = Replace(strText$, "b", s(2))
        strText$ = Replace(strText$, "c", s(3))
        strText$ = Replace(strText$, "d", s(4))
        strText$ = Replace(strText$, "e", s(5))
        strText$ = Replace(strText$, "f", s(6))
        strText$ = Replace(strText$, "g", s(7))
        strText$ = Replace(strText$, "h", s(8))
        strText$ = Replace(strText$, "i", s(9))
        strText$ = Replace(strText$, "j", s(10))
        strText$ = Replace(strText$, "k", s(11))
        strText$ = Replace(strText$, "l", s(12))
        strText$ = Replace(strText$, "m", s(13))
        strText$ = Replace(strText$, "n", s(14))
        strText$ = Replace(strText$, "o", s(15))
        strText$ = Replace(strText$, "p", s(16))
        strText$ = Replace(strText$, "q", s(17))
        strText$ = Replace(strText$, "r", s(18))
        strText$ = Replace(strText$, "s", s(19))
        strText$ = Replace(strText$, "t", s(20))
        strText$ = Replace(strText$, "u", s(21))
        strText$ = Replace(strText$, "v", s(22))
        strText$ = Replace(strText$, "w", s(23))
        strText$ = Replace(strText$, "x", s(24))
        strText$ = Replace(strText$, "y", s(25))
        strText$ = Replace(strText$, "z", s(26))
        strText$ = Replace(strText$, "A", s(27))
        strText$ = Replace(strText$, "B", s(28))
        strText$ = Replace(strText$, "C", s(29))
        strText$ = Replace(strText$, "D", s(30))
        strText$ = Replace(strText$, "E", s(31))
        strText$ = Replace(strText$, "F", s(32))
        strText$ = Replace(strText$, "G", s(33))
        strText$ = Replace(strText$, "H", s(34))
        strText$ = Replace(strText$, "I", s(35))
        strText$ = Replace(strText$, "J", s(36))
        strText$ = Replace(strText$, "K", s(37))
        strText$ = Replace(strText$, "L", s(38))
        strText$ = Replace(strText$, "M", s(39))
        strText$ = Replace(strText$, "N", s(40))
        strText$ = Replace(strText$, "O", s(41))
        strText$ = Replace(strText$, "P", s(42))
        strText$ = Replace(strText$, "Q", s(43))
        strText$ = Replace(strText$, "R", s(44))
        strText$ = Replace(strText$, "S", s(45))
        strText$ = Replace(strText$, "T", s(46))
        strText$ = Replace(strText$, "U", s(47))
        strText$ = Replace(strText$, "V", s(48))
        strText$ = Replace(strText$, "W", s(49))
        strText$ = Replace(strText$, "X", s(50))
        strText$ = Replace(strText$, "Y", s(51))
        strText$ = Replace(strText$, "Z", s(52))
        strText$ = Replace(strText$, "ae", s(0))
        EliteType = strText$
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function EliteType
' -----------------------------

' @Function SpaceCharacters
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will put a space between every character in the text, like this:  C o o l
'@Comments
' Parameter:
'       strText - string to  processed
'       AmountOfSpaces - Number of Spaces
Function SpaceCharacters(strText As String, AmountOfSpaces As Integer)
        On Error GoTo error
        Dim i As Long
        Dim SpaceStr As String
        If AmountOfSpaces > 100 Then
        AmountOfSpaces = 100
        ElseIf AmountOfSpaces < 1 Then
        AmountOfSpaces = 1
        End If
        For i = 1 To AmountOfSpaces
        SpaceStr$ = SpaceStr$ + " "
        Next i
        Dim endstr As String
        For i = 1 To Len(strText$)
        endstr$ = endstr$ & Mid$(strText$, i, 1) & SpaceStr$
        Next i
        endstr$ = Mid$(endstr$, 1, Len(endstr$) - 1)
        SpaceCharacters = endstr$
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function SpaceCharacters
' -----------------------------

' @Function DoubleCharacters
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will double every character in the text, like this:  CCooooll
'@Comments
' Parameter:
'       strText - string to  processed
'       AmountOfExtras - Number of Extra
Function DoubleCharacters(strText As String, AmountOfExtras As Integer)
        On Error GoTo error
        Dim i As Long
        Dim i2 As Long
        Dim endstr As String
        If AmountOfExtras > 100 Then
        AmountOfExtras = 100
        ElseIf AmountOfExtras < 1 Then
        AmountOfExtras = 1
        End If
        For i = 1 To Len(strText$)
          For i2 = 1 To AmountOfExtras
          endstr$ = endstr$ & Mid$(strText$, i, 1)
          Next i2
        Next i
        DoubleCharacters = endstr$
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function DoubleCharacters
' -----------------------------

' @Function EchoText
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will "echo" the text, like this:  Cool ool ol l
'@Comments
' Parameter:
'       strText - string to  processed
'       Reverse - Reverse order
Function EchoText(strText As String, Reverse As Boolean)
        On Error GoTo error
        Dim i As Long
        Dim endstr As String
        For i = 1 To Len(strText$)
          If Reverse = True Then
          endstr$ = Mid$(strText$, i, Len(strText$) - (i - 1)) & " " & endstr$
          Else
          endstr$ = endstr$ & Mid$(strText$, i, Len(strText$) - (i - 1)) & " "
          End If
        Next i
        endstr$ = Mid$(endstr$, 1, Len(endstr$) - 1)
        EchoText = endstr$
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function EchoText
' -----------------------------

' @Function Scramble
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will scramble text up, example:  oCol
'@Comments
' Parameter:
'       strText - string to  processed
'       intKey - key number
Function Scramble(strText As String, intKey As Integer)
        On Error GoTo error
        Dim RndNum As Long
        Dim i As Long
        Dim endstr As String
        Dim ListN(10000) As Long
        Dim CurPos As Long
        Randomize intKey
        CurPos = 0
        strText$ = Mid$(strText$, 1, 10000)
Start:
        RndNum = Int((Len(strText$) - 1 + 1) * Rnd + 1)
        For i = 0 To CurPos
          If RndNum = ListN(i) Then
          GoTo Start
          End If
        Next i
        ListN(CurPos) = RndNum
        CurPos = CurPos + 1
        If Not CurPos = Len(strText$) Then
        GoTo Start
        End If
        For i = 0 To CurPos - 1
        endstr$ = endstr$ & Mid$(strText$, ListN(i), 1)
        Next i
        Scramble = endstr$
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function Scramble
' -----------------------------

' @Function NUMERIC FUNCTIONS
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will scramble text up, example:  oCol
'@Comments
Function Add(num1 As Long, num2 As Long) As Long
'Add two numbers
On Error GoTo error
Add = Val(num1) + Val(num2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Subtract(num1 As Long, num2 As Long) As Long
'Subtract two numbers
On Error GoTo error
Subtract = Val(num1) - Val(num2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Divide(num1 As Long, num2 As Long) As Long
'Divide two numbers
On Error GoTo error
Divide = Val(num1) / Val(num2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Percentage(num1 As Long, Total As Long) As Single
'Divide two numbers
On Error GoTo error
Percentage = Format((Val(num1) / Val(Total)) * 100, "0.0")
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Multiply(num1 As Long, num2 As Long) As Long
'Multiply two numbers
On Error GoTo error
Multiply = Val(num1) * Val(num2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function ToPower(num1 As Long, num2 As Long) As Long
'Bring num1 to the power (exponent) of num2
On Error GoTo error
ToPower = Val(num1) ^ Val(num2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function ToRoot(num1 As Long, num2 As Long) As Long
'Bring num1 to the root of num2
On Error GoTo error
ToRoot = Val(num1) ^ (1 / Val(num2))
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function FractionToDecimal(numerator As Integer, denominator As Integer) As Long
'Turns a fraction into a decimal
On Error GoTo error
FractionToDecimal = numerator / denominator
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function DecimalToPercentage(DecimalNum As Long) As String
'Turns a decimal into a percentage
On Error GoTo error
DecimalToPercentage = (DecimalNum * 100) & "%"
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function PercentageToDecimal(PercentNum As String) As Long
'Turns a percentage into a decimal
On Error GoTo error
If Mid$(PercentNum$, Len(PercentNum$), 1) = "%" Then
PercentNum$ = Mid$(PercentNum$, 2, Len(PercentNum$) - 1)
End If
PercentageToDecimal = Val(PercentNum$) / 100
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function AreaOfCircle(radius As Long)
Dim pi
'Gets the area of a circle
On Error GoTo error
pi = 3.141592654
AreaOfCircle = pi * (radius ^ 2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Circumference(radius As Long)
Dim pi
'Gets the circumference of a circle
On Error GoTo error
pi = 3.141592654
Circumference = pi * 2 * radius
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function AreaOfSquare(side As Long)
'Gets the area of a square
On Error GoTo error
AreaOfSquare = side ^ 2
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function PerimeterOfSquare(side As Long)
'Gets the perimeter of a square
On Error GoTo error
PerimeterOfSquare = 4 * side
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function PerimeterOfRectangle(Length As Long, width As Long)
'Gets the perimeter of a rectangle
On Error GoTo error
PerimeterOfRectangle = (2 * Length) + (2 * width)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function AreaOfRectangle(Length As Long, width As Long)
'Gets the area of a rectangle
On Error GoTo error
AreaOfRectangle = Length * width
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function AreaOfTriangle(base As Long, height As Long)
'Gets the area of a triangle
On Error GoTo error
AreaOfTriangle = (1 / 2) * base * height
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function PerimeterOfTriangle(side1 As Long, side2 As Long, side3 As Long)
'Gets the perimeter of a triangle
On Error GoTo error
PerimeterOfTriangle = side1 + side2 + side3
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function PerimeterOf4SidedPolygon(side1 As Long, side2 As Long, side3 As Long, side4 As Long)
'Gets the perimeter of any 4 sided polygon
On Error GoTo error
PerimeterOf4SidedPolygon = side1 + side2 + side3 + side4
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function VolumeOfCube(edge As Long)
'Gets the volume of a cube
On Error GoTo error
VolumeOfCube = edge ^ 3
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function VolumeOfPrism(base As Long, height As Long)
'Gets the volume of a prism
On Error GoTo error
VolumeOfPrism = base * height
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function VolumeOfSphere(radius As Long)
Dim pi
'Gets the volume of a sphere
On Error GoTo error
pi = 3.141592654
VolumeOfSphere = (4 / 3) * pi * (radius ^ 3)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function VolumeOfPyramid(base As Long, height As Long)
'Gets the volume of a pyramid
On Error GoTo error
VolumeOfPyramid = (1 / 3) * base * height
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function VolumeOfCone(radius As Long, height As Long)
Dim pi
'Gets the volume of a cone
On Error GoTo error
pi = 3.141592654
VolumeOfCone = (1 / 3) * pi * (radius ^ 2) * height
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function VolumeOfCylinder(radius As Long, height As Long)
Dim pi
'Gets the volume of a cylinder
On Error GoTo error
pi = 3.141592654
VolumeOfCylinder = pi * height * (radius ^ 2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function RandomNumber(Max As Double, Min As Double)
'Create a random number
On Error GoTo error
Randomize Timer
RandomNumber = Int((Max - Min + 1) * Rnd + Min)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function NUMERIC FUNCTIONS
' -----------------------------

' @Function RemoveApos
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will remove all apostraphes
'@Comments
' strText - String to be processed
Public Function RemoveApos(strText As String) As String
  strText = Replace(strText, "'", "`")
  strText = Replace(strText, ";", "")
  strText = Trim(strText)
  RemoveApos = Trim(strText)
End Function
' Function RemoveApos
' -----------------------------

' @Function ProperCase
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns proper cased text
'@Comments
' strText - String to be processed
Function ProperCase(strText As String) As String
    Dim i As Integer

    ' prepare the result
    ProperCase = StrConv(strText, vbProperCase)

    ' restore all those characters that were capitalized
    For i = 1 To Len(strText)
        Select Case Asc(Mid$(strText, i, 1))
            Case 65 To 90   ' A-Z
                Mid$(ProperCase, i, 1) = Mid$(strText, i, 1)
        End Select
    Next
End Function
' Function ProperCase
' -----------------------------

' @Function CleanText
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns proper cased text
'@Comments
' strText - String to be processed
Public Function CleanText(strText As String) As Double
    strText = Replace(strText, ";", "")
    strText = Replace(strText, "%", "")
    strText = Replace(strText, ",", "")
    strText = Replace(strText, "'", "`")
    CleanText = Val(strText)
End Function
' Function CleanText
' -----------------------------

' @Function FilterText
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns proper cased text
'@Comments
' strText - String to be processed
Public Function FilterText(strText As String) As String
    strText = Replace(strText, ";", ":")
    strText = Replace(strText, "<", "")
    strText = Replace(strText, ">", "")
    strText = Replace(strText, "'", "`")
    FilterText = Trim(strText)
End Function
' Function FilterText
' -----------------------------


' @Function GetValueFromKey
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns value based from key <KEY=Value>
'@Comments
' strText - String to be processed
Public Function GetValueFromKey(strText As String, strKey As String) As String
GetValueFromKey = StrBetween(strText, "«" & strKey & "=", "»")
End Function
' Function GetValueFromKey
' -----------------------------


' @Function OnlyFolderFriendly
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns only AlphaNumeric numbers
'@Comments
' strText - String to be processed
Public Function OnlyFolderFriendly(theString) As String
Dim strAlphaNumeric, i, CleanedString, strChar
      strAlphaNumeric = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-()@#$^&_+=,[].& " 'Used to check for numeric characters.
      For i = 1 To Len(theString)
          strChar = Mid(theString, i, 1)
          If InStr(strAlphaNumeric, strChar) Then
              CleanedString = CleanedString & strChar
          Else
              CleanedString = CleanedString & "_"
          End If
      Next
      OnlyFolderFriendly = CleanedString
End Function
' Function OnlyAlphaNumeric
' -----------------------------

' @Function OnlyAlphaNumeric
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns only AlphaNumeric numbers
'@Comments
' strText - String to be processed
Public Function OnlyAlphaNumeric(theString) As String
Dim strAlphaNumeric, i, CleanedString, strChar
      strAlphaNumeric = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ" 'Used to check for numeric characters.
      For i = 1 To Len(theString)
          strChar = Mid(theString, i, 1)
          If InStr(strAlphaNumeric, strChar) Then
              CleanedString = CleanedString & strChar
          Else
              CleanedString = CleanedString & "_"
          End If
      Next
      OnlyAlphaNumeric = CleanedString
End Function
' Function OnlyAlphaNumeric
' -----------------------------

' @Function RemoveAllEnter
' -----------------------------
'@Author Roland Ross Hadi
'@Description Removes all linefeed
'@Comments
' Expr - String to be processed
Public Function RemoveAllEnter(Expr)
Dim i, tmp
tmp = Expr
For i = 1 To 1
    tmp = Replace(tmp, vbCrLf, " ")
Next
For i = 1 To 1
    tmp = Replace(tmp, Chr(10) & Chr(13), " ")
Next
For i = 1 To 1
    tmp = Replace(tmp, Chr(10), " ")
Next
For i = 1 To 1
    tmp = Replace(tmp, Chr(13), " ")
Next
For i = 1 To 1
    tmp = Replace(tmp, vbTab, " ")
Next
RemoveAllEnter = tmp
End Function
' Function RemoveAllEnter
' -----------------------------

' @Function Array_RemoveItem
' -----------------------------
'@Author Roland Ross Hadi
'@Description Removes array element
'@Comments
' ItemArray - String to be processed
' ItemElement - Index
Public Sub Array_RemoveItem(ItemArray As Variant, ByVal ItemElement As Long)
Dim lCtr As Long
Dim lTop As Long
Dim lBottom As Long

If Not IsArray(ItemArray) Then
    Err.Raise 13, , "Type Mismatch"
    Exit Sub
End If

lTop = UBound(ItemArray)
lBottom = LBound(ItemArray)

If ItemElement < lBottom Or ItemElement > lTop Then
    Err.Raise 9, , "Subscript out of Range"
    Exit Sub
End If

For lCtr = ItemElement To lTop - 1
    ItemArray(lCtr) = ItemArray(lCtr + 1)
Next
On Error GoTo ErrorHandler:

ReDim Preserve ItemArray(lBottom To lTop - 1)

Exit Sub
ErrorHandler:
  'An error will occur if array is fixed
    Err.Raise Err.Number, , _
       "You must pass a resizable array to this function"
End Sub
' Function RemoveAllEnter
' -----------------------------

' @Function TextFill
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns only AlphaNumeric numbers
'@Comments
' strText - String to be processed
Public Function TextFill(theString, Length, FillChar, Optional FromStart As Boolean = False) As String
Dim i, tmp
    If FromStart = True Then
        tmp = theString
        For i = 1 To Length - Len(theString)
            tmp = FillChar & tmp
        Next
        TextFill = tmp
    Else
        tmp = theString
        For i = 1 To Length - Len(theString)
            tmp = tmp & FillChar
        Next
        TextFill = tmp
    End If
End Function
' Function OnlyAlphaNumeric
' -----------------------------

' @Function CountInstance
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns only AlphaNumeric numbers
'@Comments
' strText - String to be processed
Public Function CountInstance(strText As String, find As String)
Dim i, cnt
For i = 1 To Len(strText)
    If Mid(strText, i, 1) = find Then cnt = cnt + 1
Next
CountInstance = cnt
End Function
' Function CountInstance
' -----------------------------

' @Function CountWordInstance
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns only AlphaNumeric numbers
'@Comments
' strText - String to be processed
Public Function CountWordInstance(strText As String, find As String)
Dim i, cnt
cnt = 0
If Trim(find) = "" Then CountWordInstance = 0: Exit Function
For i = 1 To Len(strText)
    If Mid(strText, i, Len(find)) = find Then cnt = cnt + 1
Next
CountWordInstance = cnt
End Function
' Function CountInstance
' -----------------------------

' @Function GetCaretPosition
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the caret position in a text file
'@Comments
' strText - String to be processed
Public Function GetCaretPosition(RTF) As Long
Dim StartSelect, Endselect As Long
SendMessage RTF.hwnd, EM_GETSEL, StartSelect, Endselect
GetCaretPosition = Endselect
End Function
' Function GetCaretPosition
' -----------------------------

' @Function VBTextColored
' -----------------------------
'@Author Roland Ross Hadi
'@Description Change the font color and size vb style
'@Comments
' rtf - RichTextBox object
Public Sub VBTextColored(RTF)
    Dim sBuffer    As String
    Dim i          As Long
    Dim nJ         As Long
    Dim sTmpWord   As String
    Dim nStartPos  As Long
    Dim nSelLen    As Long
    Dim nWordPos   As Long
    Dim gsBlackKeywords    As String
    Dim gsBlueKeyWords     As String
    Dim oleComment As OLE_COLOR
    Dim oleKeyword As OLE_COLOR
    Dim oleText As OLE_COLOR
    
    gsBlackKeywords = "*Abs*Add*AddItem*AppActivate*Array*Asc*Atn*Beep*Begin*BeginProperty*ChDir*ChDrive*Choose*Chr*Clear*Collection*Command*Cos*CreateObject*CurDir*DateAdd*DateDiff*DatePart*DateSerial*DateValue*Day*DDB*DeleteSetting*Dir*DoEvents*EndProperty*Environ*EOF*Err*Exp*FileAttr*FileCopy*FileDateTime*FileLen*Fix*Format*FV*GetAllSettings*GetAttr*GetObject*GetSetting*Hex*Hide*Hour*InputBox*InStr*Int*Int*IPmt*IRR*IsArray*IsDate*IsEmpty*IsError*IsMissing*IsNull*IsNumeric*IsObject*Item*Kill*LCase*Left*Len*Load*Loc*LOF*Log*LTrim*Me*Mid*Minute*MIRR*MkDir*Month*Now*NPer*NPV*Oct*Pmt*PPmt*PV*QBColor*Raise*Randomize*Rate*Remove*RemoveItem*Reset*RGB*Right*RmDir*Rnd*RTrim*SaveSetting*Second*SendKeys*SetAttr*Sgn*Shell*Sin*Sin*SLN*Space*Sqr*Str*StrComp*StrConv*Switch*SYD*Tan*Text*Time*Time*Timer*TimeSerial*TimeValue*Trim*TypeName*UCase*Unload*Val*VarType*WeekDay*Width*Year*"
    gsBlueKeyWords = "*#Const*#Else*#ElseIf*#End If*#If*Alias*Alias*And*As*Base*Binary*Boolean*Byte*ByVal*Call*Case*CBool*CByte*CCur*CDate*CDbl*CDec*CInt*CLng*Close*Compare*Const*CSng*CStr*Currency*CVar*CVErr*Decimal*Declare*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase*Error*Exit*Explicit*False*For*Function*Get*Global*GoSub*GoTo*If*Imp*In*Input*Input*Integer*Is*LBound*Let*Lib*Like*Line*Lock*Long*Loop*LSet*Name*New*Next*Not*Object*On*Open*Option*Or*Output*Print*Private*Property*Public*Put*Random*Read*ReDim*Resume*Return*RSet*Seek*Select*Set*Single*Spc*Static*String*Stop*Sub*Tab*Then*Then*True*Type*UBound*Unlock*Variant*Wend*While*With*Xor*Nothing*To*"
    oleComment = &H8000&
    oleKeyword = vbBlue
    oleText = vbBlack
    
    '--
    'Apply the font and colour styles
    '--
    RTF.SelStart = 0
    RTF.SelLength = Len(RTF.Text)
    RTF.SelColor = oleText
    RTF.SelFontName = "Courier New"
    RTF.SelFontSize = 10

    '--
    'Loop through every character in the text and check
    'for different letters and characters
    '--
    sBuffer = RTF.Text
    sTmpWord = ""

    For i = 1 To Len(sBuffer)

        Select Case Mid(sBuffer, i, 1)

            Case "A" To "Z", "a" To "z", "_"

                If sTmpWord = "" Then nStartPos = i
                sTmpWord = sTmpWord & Mid(sBuffer, i, 1)

            Case Chr(34)

                nSelLen = 1

                For nJ = 1 To 9999999

                    If Mid(sBuffer, i + 1, 1) = Chr(34) Then
                        i = i + 2
                        Exit For
                    Else
                        nSelLen = nSelLen + 1
                        i = i + 1
                    End If

                Next

            Case Chr(39)
                RTF.SelStart = i - 1
                nSelLen = 0

                For nJ = 1 To 9999999

                    If Mid(sBuffer, i, 2) = vbCrLf Then
                        Exit For
                    Else
                        nSelLen = nSelLen + 1
                        i = i + 1
                        If nSelLen > Len(sBuffer) Then Exit For
                    End If

                Next

                RTF.SelLength = nSelLen
                RTF.SelColor = oleComment

            Case Else
                If Trim(sTmpWord) <> "" Then
                    RTF.SelStart = nStartPos - 1
                    RTF.SelLength = Len(sTmpWord)
                    nWordPos = InStr(1, gsBlackKeywords, "*" & sTmpWord & "*", 1)
                    If nWordPos <> 0 Then
                        RTF.SelColor = oleText
                        RTF.SelText = Mid(gsBlackKeywords, nWordPos + 1, Len(sTmpWord))
                    End If

                    nWordPos = InStr(1, gsBlueKeyWords, "*" & sTmpWord & "*", 1)

                    If nWordPos <> 0 Then
                        RTF.SelColor = oleKeyword
                        RTF.SelText = Mid(gsBlueKeyWords, nWordPos + 1, Len(sTmpWord))
                    End If

                    If UCase(sTmpWord) = "REM" Then
                        RTF.SelStart = i - 4
                        RTF.SelLength = 3

                        For nJ = 1 To 9999999
                            If Mid(sBuffer, i, 2) = vbCrLf Then
                                Exit For
                            Else
                                RTF.SelLength = RTF.SelLength + 1
                                i = i + 1
                            End If
                        Next

                        RTF.SelColor = oleKeyword
                        RTF.SelText = LCase(RTF.SelText)

                    End If
                End If

                sTmpWord = ""
        End Select
    Next
    RTF.SelStart = 0
End Sub
' Function CountInstance
' -----------------------------

' @Function HighlightWords
' -----------------------------
'@Author Roland Ross Hadi
'@Description Change the font color and size vb style
'@Comments
' rtf - RichTextBox object
 Public Function HighlightWords(rtb, _
                                  sFindString As String, _
                                  lColor As Long) _
                                  As Integer

        Dim lFoundPos As Long           'Position of first character
                                        'of match
        Dim lFindLength As Long         'Length of string to find
        Dim lOriginalSelStart As Long
        Dim lOriginalSelLength As Long
        Dim iMatchCount As Integer      'Number of matches

        'Save the insertion points current location and length
        lOriginalSelStart = rtb.SelStart
        lOriginalSelLength = rtb.SelLength

        'Cache the length of the string to find
        lFindLength = Len(sFindString)

        'Attempt to find the first match
        lFoundPos = rtb.find(sFindString, 0, , rtfNoHighlight)
        While lFoundPos > 0
          iMatchCount = iMatchCount + 1

          rtb.SelStart = lFoundPos
          'The SelLength property is set to 0 as
          'soon as you change SelStart
          rtb.SelLength = lFindLength
          rtb.SelColor = lColor

          'Attempt to find the next match
          lFoundPos = rtb.find(sFindString, _
            lFoundPos + lFindLength, , rtfNoHighlight)
        Wend

        'Restore the insertion point to its original
        'location and length
        rtb.SelStart = lOriginalSelStart
        rtb.SelLength = lOriginalSelLength

        'Return the number of matches
        HighlightWords = iMatchCount

      End Function
' Function HighlightWords
' -----------------------------

' @Function Switch_Month_Day_DateFormat
' -----------------------------
'@Author Roland Ross Hadi
'@Description Swap date formats
'@Comments
'
Function Switch_Month_Day_DateFormat(strDate, strDelimiter, strFormatting)
    Dim tmp
    tmp = Split(strDate, strDelimiter)
    Switch_Month_Day_DateFormat = Format(tmp(1) & strDelimiter & tmp(0) & strDelimiter & tmp(2), strFormatting)
End Function
' Function Switch_Month_Day_DateFormat
' -----------------------------
