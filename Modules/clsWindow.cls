VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************************************************************************************************
'*      LIBRARY NAME: clsWindow.cls
'*      PURPOSE: Compilation of functions that deals with webcam
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

        Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
        Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

        Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
        Private Const HWND_NOTOPMOST = -2
        Private Const HWND_TOPMOST = -1
        Private Const SWP_NOMOVE = &H2
        Private Const SWP_NOSIZE = &H1
        Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE

        Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
        Private Const SPI_SCREENSAVERRUNNING = 97
        
        Private Type SECURITY_ATTRIBUTES
          nLength As Long
          lpSecurityDescriptor As Long
          bInheritHandle As Boolean
        End Type
        
        Private Const LWA_COLORKEY = 1
        Private Const LWA_ALPHA = 2
        Private Const LWA_BOTH = 3
        Private Const WS_EX_LAYERED = &H80000
        Private Const GWL_EXSTYLE = -20
        Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal color As Long, ByVal x As Byte, ByVal alpha As Long) As Boolean
        Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
        Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'*******************************************************************************************************************************************************************************************

' @Function WndHide
' -----------------------------
'@Author Roland Ross Hadi
'@Description Hide window
'@Comments
' Parameter:
'   lngHwnd - Window hwnd
Public Function WndHide(lngHwnd)
    ShowWindow CLng(lngHwnd), 0
End Function
' Function WndHide
' -----------------------------

' @Function WndShow
' -----------------------------
'@Author Roland Ross Hadi
'@Description Show window
'@Comments
' Parameter:
'   lngHwnd - Window hwnd
Public Function WndShow(lngHwnd)
    ShowWindow CLng(lngHwnd), 5
End Function
' Function WndShow
' -----------------------------

' @Function WndFlash
' -----------------------------
'@Author Roland Ross Hadi
'@Description Flash window
'@Comments
' Parameter:
'   lngHwnd - Window hwnd
Public Function WndFlash(lngHwnd)
    FlashWindow CLng(lngHwnd), True
End Function
' Function WndFlash
' -----------------------------

' @Function DisableCtrlAltDel
' -----------------------------
'@Author Roland Ross Hadi
'@Description Disable Ctrl + Alt + Del
'@Comments
' Parameter:
'
Public Sub DisableCtrlAltDel()
        On Error GoTo error
        Dim ret As Integer
        Dim pOld As Boolean
        ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
        Exit Sub
error:          MsgBox Err.Description, vbExclamation, "Error"
End Sub
' @Function DisableCtrlAltDel
' -----------------------------

' @Function EnableCtrlAltDel
' -----------------------------
'@Author Roland Ross Hadi
'@Description Enable Ctrl + Alt + Del
'@Comments
' Parameter:
'
Public Sub EnableCtrlAltDel()
        On Error GoTo error
        Dim ret As Integer
        Dim pOld As Boolean
        ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
        Exit Sub
error:          MsgBox Err.Description, vbExclamation, "Error"
End Sub
' Function EnableCtrlAltDel
' -----------------------------

' @Function HideCtrlAltDel
' -----------------------------
'@Author Roland Ross Hadi
'@Description Hide this app from Ctrl + Alt + Del
'@Comments
' Parameter:
'
Public Sub HideCtrlAltDel()
        On Error GoTo error
        App.TaskVisible = False
        Exit Sub
error:          MsgBox Err.Description, vbExclamation, "Error"
End Sub
' Function HideCtrlAltDel
' -----------------------------

' @Function ShowCtrlAltDel
' -----------------------------
'@Author Roland Ross Hadi
'@Description Show this app in Ctrl + Alt + Del
'@Comments
' Parameter:
'
Public Sub ShowCtrlAltDel()
        On Error GoTo error
        App.TaskVisible = True
        Exit Sub
error:          MsgBox Err.Description, vbExclamation, "Error"
End Sub
' Function ShowCtrlAltDel
' -----------------------------

' @Function Ontop
' -----------------------------
'@Author Roland Ross Hadi
'@Description Make a form always ontop of other windows
'@Comments
' Parameter:
'
Public Sub Ontop(FormName As Form)
        On Error GoTo error
        Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
        Exit Sub
error:          MsgBox Err.Description, vbExclamation, "Error"
End Sub
' Function Ontop
' -----------------------------

' @Function NotOntop
' -----------------------------
'@Author Roland Ross Hadi
'@Description Make a form not always ontop of other windows
'@Comments
' Parameter:
'
Public Sub NotOntop(FormName As Form)
        On Error GoTo error
        Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
        Exit Sub
error:          MsgBox Err.Description, vbExclamation, "Error"
End Sub
' Function NotOntop
' -----------------------------

' @Function SetTranslucent
' -----------------------------
'@Author Roland Ross Hadi
'@Description Make a form translucent
'@Comments
' Parameter:
'
Sub SetTranslucent(ThehWnd As Long, nTrans As Integer)
    On Error GoTo ErrorRtn

    'SetWindowLong and SetLayeredWindowAttributes are API functions, see MSDN for details
    Dim attrib As Long
    attrib = GetWindowLong(ThehWnd, GWL_EXSTYLE)
    SetWindowLong ThehWnd, GWL_EXSTYLE, attrib Or WS_EX_LAYERED
    SetLayeredWindowAttributes ThehWnd, RGB(255, 255, 0), nTrans, LWA_ALPHA
    Exit Sub
ErrorRtn:
    MsgBox Err.Description & " Source : " & Err.Source
End Sub
' Function SetTranslucent
' -----------------------------

