VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyboardAndMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************************************************************************************************
'*      LIBRARY NAME: clsKeyboardAndMouse.cls
'*      PURPOSE: Compilation of functions that deals with Keyboard and Mouse
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
        
        Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
        Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
        Private Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
        Private Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
        Private Declare Function GetDoubleClickTime Lib "user32" () As Long
        Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
        
        ' this sub contains all of the mouse and keyboard subs.
        Private Const KEYEVENTF_KEYUP = &H2
        Private Const SPI_SCREENSAVERRUNNING = 97&
        Private Declare Function CharToOem& Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String)
        Private Declare Function GetAsyncKeyState% Lib "user32" (ByVal vKey As Long)
        Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
        'Declare Function GetKeyboardState& Lib "user32" (pbKeyState As String)
        Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
        Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
        Declare Function OemKeyScan& Lib "user32" (ByVal wOemChar As Integer)
        Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
        Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
        Declare Function VkKeyScan% Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte)

        Private Type RECT
            left As Long
            top As Long
            right As Long
            bottom As Long
        End Type

        Private Type POINTAPI
            x As Long
            Y As Long
        End Type

        Private Property Get MouseX() As Long
            Dim t As POINTAPI
            GetCursorPos t
            MouseX = t.x
        End Property

        Private Property Get MouseY() As Long
            Dim t As POINTAPI
            GetCursorPos t
            MouseY = t.Y
        End Property

        Private Property Let MouseX(newval As Long)
            SetCursorPos newval, MouseY
        End Property

        Private Property Let MouseY(newval As Long)
            SetCursorPos MouseX, newval
        End Property

        Private Property Get MouseDoubleClickTime() As Long
            MouseDoubleClickTime = GetDoubleClickTime
        End Property

        Private Property Let MouseDoubleClickTime(newval As Long)
            SetDoubleClickTime newval
        End Property
        
'*******************************************************************************************************************************************************************************************

' @Function MouseClip
' -----------------------------
'@Author Roland Ross Hadi
'@Description Clips mouse cursor
'@Comments
' Parameter:
'       mX - x axis
'       mY - y axis
'       mWidth - clip width
'       mHeight - clip height
Public Sub MouseClip(mX, mY, mWidth, mHeight)
    Dim rct As RECT
    rct.left = mX
    rct.top = mY
    rct.right = mWidth + mX
    rct.bottom = mHeight + mY
    ClipCursor rct
End Sub
' Function MouseClip
' -----------------------------

' @Function MouseSwapButtons
' -----------------------------
'@Author Roland Ross Hadi
'@Description Clips mouse cursor
'@Comments
' Parameter:
'       bolSwap - True/False
Public Sub MouseSwapButtons(bolSwap As Boolean)
    MouseSwapButtons CLng(bolSwap)
End Sub
' Function MouseSwapButtons
' -----------------------------

' @Function ShowMouseCursor
' -----------------------------
'@Author Roland Ross Hadi
'@Description Shows mouse cursor
'@Comments
' Parameter:
'
Sub ShowMouseCursor()
    Dim rtn As Long
    rtn = ShowCursor(True)
End Sub
' Function ShowMouseCursor
' -----------------------------

' @Function HideMouseCursor
' -----------------------------
'@Author Roland Ross Hadi
'@Description Hides mouse cursor
'@Comments
' Parameter:
'
Sub HideMouseCursor()
    Dim rtn As Long
    rtn = ShowCursor(False)
End Sub
' Function HideMouseCursor
' -----------------------------

' @Function DisableTaskKeys
' -----------------------------
'@Author Roland Ross Hadi
'@Description Disables ctrl-alt-del, alt-tab, ctrl-f4, etc., keeping you in charge...
'@Comments
' Parameter:
'
Sub DisableTaskKeys()
    Dim rtn As Long
    rtn = SystemParametersInfo(SPI_SCREENSAVERRUNNING, 1&, 0&, 0)
End Sub
' Function DisableTaskKeys
' -----------------------------

' @Function EnableTaskKeys
' -----------------------------
'@Author Roland Ross Hadi
'@Description Enables ctrl-alt-del, alt-tab, ctrl-f4, etc., keeping you in charge...
'@Comments
' Parameter:
'
Sub EnableTaskKeys()
    Dim rtn As Long
    rtn = SystemParametersInfo(SPI_SCREENSAVERRUNNING, 0&, 0&, 0)
End Sub
' Function EnableTaskKeys
' -----------------------------

' @Function ScanCodeToAscii
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns scan code to ascii
'@Comments
' Parameter:
'     lngScanCode - Scan code
Function ScanCodeToAscii(lngScanCode As Long) As Long
    ScanCodeToAscii = MapVirtualKey(lngScanCode, 2)
End Function
' Function ScanCodeToAscii
' -----------------------------

' @Function WasKeyPressed
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns if key was pressed
'@Comments
' Parameter:
'     lngVBKey - VB key code
Function WasKeyPressed(lngVBKey As Long)
    Dim ScanCode As Integer
    ScanCode = GetKeyState(lngVBKey)
    If ScanCode And &HFFF0 > 0 Then
        WasKeyPressed = True
    Else
        WasKeyPressed = False
    End If
End Function
' Function WasKeyPressed
' -----------------------------

' @Function IsKeyPressed
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns if key was pressed
'@Comments
' Parameter:
'     lngVBKey - VB key code
Function IsKeyPressed(lngVBKey As Long) As Boolean
    Dim KeyState As Integer
    KeyState = GetAsyncKeyState(lngVBKey)

    If KeyState And &H8000 = &H8000 Then
        IsKeyPressed = True     ' : Debug.Print VBKey
    Else
        IsKeyPressed = False
    End If
End Function
' Function IsKeyPressed
' -----------------------------

' @Function IsAsciiKeyPressed
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns if key was pressed
'@Comments
' Parameter:
'     strAscii - ASCII code
'Function IsAsciiKeyPressed(strAscii As Integer) As Boolean
'    '
'End Function
' Function IsAsciiKeyPressed
' -----------------------------

' @Function TypeAsciiKey
' -----------------------------
'@Author Roland Ross Hadi
'@Description Types ASCII keys
'@Comments
' Parameter:
'     strChar - Character code
'     intAllStates - Key state code
Public Sub TypeAsciiKey(ByVal strChar As Integer, intAllStates As Integer)
    Dim c As String
    Dim vk%
    Dim Scan%
    Dim OemChar$
'    Dim dl&
    Dim ShiftState As Integer
'    Dim ss As Long
    Dim CtrlState As Integer
    Dim AltState As Integer
    Dim ShiftScanCode As Integer
    Dim CtrlScanCode As Integer
    Dim AltScanCode As Integer
    ShiftScanCode = MapVirtualKey(vbKeyShift, 0)
    CtrlScanCode = MapVirtualKey(vbKeyControl, 0)
    AltScanCode = MapVirtualKey(vbKeyMenu, 0)
    c = ChrW$(strChar)
    'ss = (VkKeyScan(AscW(c$)) And &H100) / &HFF
    ' MsgBox ss
    ' Get the virtual key code for this character
    vk% = VkKeyScan(strChar) And &HFF
    If intAllStates = 0 Then intAllStates = (VkKeyScan(AscW(c$)) And &H100) / &HFF
    ShiftState = intAllStates And 1
    CtrlState = (intAllStates And 2) / 2
    AltState = (intAllStates And 4) / 4
    OemChar$ = "  "                                        ' 2 character buffer
    ' Get the OEM character - preinitialize the buffer
    CharToOem left$(c$, 1), OemChar$
    ' Get the scan code for this key
    Scan% = OemKeyScan(AscW(OemChar$)) And &HFF
    If ShiftState = 1 Then
        keybd_event vbKeyShift, ShiftScanCode, 0, 0
        DoEvents
    End If
    If CtrlState = 1 Then
        keybd_event vbKeyControl, CtrlScanCode, 0, 0
        DoEvents
    End If
    If AltState = 1 Then
        keybd_event vbKeyMenu, AltScanCode, 0, 0
        DoEvents
    End If
    ' Send the key down
    keybd_event vk%, Scan%, 0, 0
    DoEvents
    ' Send the key up
    keybd_event vk%, Scan%, KEYEVENTF_KEYUP, 0
    DoEvents
    If ShiftState = 1 Then
        keybd_event vbKeyShift, ShiftScanCode, KEYEVENTF_KEYUP, 0
        DoEvents
    End If
    If CtrlState = 1 Then
        keybd_event vbKeyControl, CtrlScanCode, KEYEVENTF_KEYUP, 0
        DoEvents
    End If
    If AltState = 1 Then
        keybd_event vbKeyMenu, AltScanCode, KEYEVENTF_KEYUP, 0
        DoEvents
    End If
End Sub
' Function TypeAsciiKey
' -----------------------------

' @Function TypeKey
' -----------------------------
'@Author Roland Ross Hadi
'@Description Types ASCII keys
'@Comments
' Parameter:
'     strKeyToType - Key code
'     intToggleKeys - Toggle keys
Public Sub TypeKey(strKeyToType As String, intToggleKeys As Integer)
    'togglekeys: bit1 = shift, 2 = ctrl, 3 = alt.
    Call TypeAsciiKey(AscW(strKeyToType), intToggleKeys)
End Sub
' Function TypeKey
' -----------------------------

' @Function MySendKeys
' -----------------------------
'@Author Roland Ross Hadi
'@Description Types ASCII keys
'@Comments 'doesn't work exactly like the vb version.
            'to send a shifted character, simply type it.
            'it will be converted automatically.  There are
            'three exceptions: ~^%{.  Each of these keys has to
            'be typed 2 times in a row.  They are used to toggle
            'on shift, ctrl, and/or alt states of the keyboard
            'for the next character typed.
            'special codes: + = shift, ^ = ctrl, % = alt
            'a * after the ' means it is implemented, otherwise
            ' it is slated for future implementation
            '*BACKSPACE   {BACKSPACE}, {BS}, or {BKSP}
            '*BREAK   {BREAK}
            '*CAPS LOCK   {CAPSLOCK}
            '*DEL or DELETE   {DELETE} or {DEL}
            '*DOWN ARROW  {DOWN}
            '*END {END}
            '*ENTER   {ENTER}or ~
            '*ESC {ESC}
            ' HELP    {HELP}
            ' HOME    {HOME}
            '*INS or INSERT   {INSERT} or {INS}
            '*LEFT ARROW  {LEFT}
            ' NUM LOCK    {NUMLOCK}
            ' PAGE DOWN   {PGDN}
            ' PAGE UP {PGUP}
            '*PRINT SCREEN    {PRTSC}
            '*RIGHT ARROW {RIGHT}
            ' SCROLL LOCK {SCROLLLOCK}
            '*TAB {TAB}
            '*UP ARROW    {UP}
            ' F1  {F1}
            ' F2  {F2}
            ' F3  {F3}
            ' F4  {F4}
            ' F5  {F5}
            ' F6  {F6}
            ' F7  {F7}
            ' F8  {F8}
            ' F9  {F9}
            ' F10 {F10}
            ' F11 {F11}
            ' F12 {F12}
            ' F13 {F13}
            ' F14 {F14}
            ' F15 {F15}
            ' F16 {F16}
' Parameter:
'     strStringToType - String To Type
Sub MySendKeys(strStringToType As String)
    Dim ToggleKeys As Integer
    Dim Char As String
    ToggleKeys = 0
    Dim cl As Long
    cl = 1
    Do While cl <= Len(strStringToType)
        ToggleKeys = 0
'        LcaseStringToType = LCase$(strStringToType)
        Do While Char = "%" Or Char = "^" Or Char = "+" Or Char = "{" Or Char = "~"
            Char = Mid$(strStringToType, cl, 1)
            If Char = "{" Then
                cl = cl + 1
                If Mid$(strStringToType, cl - 1, 2) = "{{" Then
                    cl = cl + 1
                    Call TypeKey(Char, 0)
                Else
                    If StrComp(Mid$(strStringToType, cl, 10), "backspace}", vbTextCompare) = 0 Then
                        cl = cl + 10
                        VirtualKeyPress vbKeyBack
                    ElseIf StrComp(Mid$(strStringToType, cl, 3), "bs}", vbTextCompare) = 0 Then
                        cl = cl + 3
                        VirtualKeyPress vbKeyBack
                    ElseIf StrComp(Mid$(strStringToType, cl, 5), "bksp}", vbTextCompare) = 0 Then
                        cl = cl + 5
                        VirtualKeyPress vbKeyBack
                    ElseIf StrComp(Mid$(strStringToType, cl, 7), "delete}", vbTextCompare) = 0 Then
                        cl = cl + 7
                        VirtualKeyPress vbKeyDelete
                    ElseIf StrComp(Mid$(strStringToType, cl, 4), "del}", vbTextCompare) = 0 Then
                        cl = cl + 4
                        VirtualKeyPress vbKeyDelete
                    ElseIf StrComp(Mid$(strStringToType, cl, 5), "home}", vbTextCompare) = 0 Then
                        cl = cl + 5
                        VirtualKeyPress vbKeyHome
                    ElseIf StrComp(Mid$(strStringToType, cl, 6), "enter}", vbTextCompare) = 0 Then
                        cl = cl + 6
                        VirtualKeyPress vbKeyReturn
                    ElseIf StrComp(Mid$(strStringToType, cl, 4), "tab}", vbTextCompare) = 0 Then
                        cl = cl + 4
                        VirtualKeyPress vbKeyTab
                    ElseIf StrComp(Mid$(strStringToType, cl, 6), "prtsc}", vbTextCompare) = 0 Then
                        cl = cl + 6
                        VirtualKeyPress vbKeySnapshot
                    ElseIf StrComp(Mid$(strStringToType, cl, 4), "end}", vbTextCompare) = 0 Then
                        cl = cl + 4
                        VirtualKeyPress vbKeyEnd
                    ElseIf StrComp(Mid$(strStringToType, cl, 4), "ins}", vbTextCompare) = 0 Then
                        cl = cl + 4
                        VirtualKeyPress vbKeyInsert
                    ElseIf StrComp(Mid$(strStringToType, cl, 7), "insert}", vbTextCompare) = 0 Then
                        cl = cl + 7
                        VirtualKeyPress vbKeyInsert
                    ElseIf StrComp(Mid$(strStringToType, cl, 5), "left}", vbTextCompare) = 0 Then
                        cl = cl + 5
                        VirtualKeyPress vbKeyLeft
                    ElseIf StrComp(Mid$(strStringToType, cl, 6), "right}", vbTextCompare) = 0 Then
                        cl = cl + 6
                        VirtualKeyPress vbKeyRight
                    ElseIf StrComp(Mid$(strStringToType, cl, 6), "break}", vbTextCompare) = 0 Then
                        cl = cl + 6
                        VirtualKeyPress vbKeyPause
                    ElseIf StrComp(Mid$(strStringToType, cl, 9), "capslock}", vbTextCompare) = 0 Then
                        cl = cl + 9
                        VirtualKeyPress vbKeyCapital
                    ElseIf StrComp(Mid$(strStringToType, cl, 10), "downarrow}", vbTextCompare) = 0 Then
                        cl = cl + 10
                        VirtualKeyPress vbKeyDown
                    ElseIf StrComp(Mid$(strStringToType, cl, 8), "uparrow}", vbTextCompare) = 0 Then
                        cl = cl + 8
                        VirtualKeyPress vbKeyUp
                    ElseIf StrComp(Mid$(strStringToType, cl, 4), "esc}", vbTextCompare) = 0 Then
                        cl = cl + 4
                        VirtualKeyPress vbKeyEscape
                    End If
                End If
            ElseIf Char = "~" Then
                cl = cl + 1
                If Mid$(strStringToType, cl, 1) = "~" Then
                    cl = cl + 1
                    Call TypeKey(Char, 0)
                Else
                    VirtualKeyPress vbKeyReturn
                End If
            ElseIf Char = "+" Then
                cl = cl + 1
                If Mid$(strStringToType, cl, 1) = "+" Then
                    cl = cl + 1
                    Call TypeKey(Char, 0)
                Else
                    ToggleKeys = ToggleKeys + 1
                End If
            ElseIf Char = "^" Then
                cl = cl + 1
                If Mid$(strStringToType, cl - 1, 2) = "^^" Then
                    cl = cl + 1
                    Call TypeKey(Char, 0)
                Else
                    ToggleKeys = ToggleKeys + 2
                End If
            ElseIf Char = "%" Then
                cl = cl + 1
                If Mid$(strStringToType, cl - 1, 2) = "%%" Then
                    cl = cl + 1
                    Call TypeKey(Char, 0)
                Else
                    ToggleKeys = ToggleKeys + 4
                End If
            End If
        Loop
'        Debug.Print Char, ToggleKeys
        If Len(Char) <> 0 Then Call TypeKey(Char, ToggleKeys)
    Loop
End Sub
' Function MySendKeys
' -----------------------------

' @Function VirtualKeyPress
' -----------------------------
'@Author Roland Ross Hadi
'@Description Types virtual keys
'@Comments
' Parameter:
'     intVirtualKeyCode - Character code
Sub VirtualKeyPress(intVirtualKeyCode As Integer)
    Dim ScanCode As Integer
    ScanCode = MapVirtualKey(intVirtualKeyCode, 0)
    keybd_event intVirtualKeyCode, ScanCode, 0, 0
    DoEvents
    keybd_event intVirtualKeyCode, ScanCode, KEYEVENTF_KEYUP, 0
    DoEvents
End Sub
' Function VirtualKeyPress
' -----------------------------

' @Function VirtualKeyPressEx
' -----------------------------
'@Author Roland Ross Hadi
'@Description Types virtual keys
'@Comments
' Parameter:
'     intVirtualKeyCode - Character code
'     intShift - Shift key
'     intCtrl - Ctrl key
'     intAlt - Alt key
Public Sub VirtualKeyPressEx(intVirtualKeyCode As Integer, intShift As Integer, intCtrl As Integer, intAlt As Integer)
    Dim ShiftScanCode As Integer
    Dim CtrlScanCode As Integer
    Dim AltScanCode As Integer

    ShiftScanCode = MapVirtualKey(vbKeyShift, 0)
    CtrlScanCode = MapVirtualKey(vbKeyControl, 0)
    AltScanCode = MapVirtualKey(vbKeyMenu, 0)

    ' Get the virtual key code for this character
    Dim ScanCode As Integer
    ScanCode = MapVirtualKey(intVirtualKeyCode, 0)

'    intShift = AllStates And 1
'    intCtrl = (AllStates And 2) / 2
'    intAlt = (AllStates And 4) / 4

    If intCtrl = 1 Then
        keybd_event vbKeyControl, CtrlScanCode, 0, 0
        DoEvents
    End If
    If intShift = 1 Then
        keybd_event vbKeyShift, ShiftScanCode, 0, 0
        DoEvents
    End If
    If intAlt = 1 Then
        keybd_event vbKeyMenu, AltScanCode, 0, 0
        DoEvents
    End If

    ' Send the key down
    keybd_event intVirtualKeyCode%, ScanCode, 0, 0
    DoEvents
    ' Send the key up
    keybd_event intVirtualKeyCode, ScanCode, KEYEVENTF_KEYUP, 0
    DoEvents

    If intShift = 1 Then
        keybd_event vbKeyShift, ShiftScanCode, KEYEVENTF_KEYUP, 0
        DoEvents
    End If
    If intCtrl = 1 Then
        keybd_event vbKeyControl, CtrlScanCode, KEYEVENTF_KEYUP, 0
        DoEvents
    End If
    If intAlt = 1 Then
        keybd_event vbKeyMenu, AltScanCode, KEYEVENTF_KEYUP, 0
        DoEvents
    End If
End Sub
' Function VirtualKeyPressEx
' -----------------------------

