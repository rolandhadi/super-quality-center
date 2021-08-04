VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************************************************************************************************
'*      LIBRARY NAME: clsSystray.cls
'*      PURPOSE: Compilation of functions that deals with system tray
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
  Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

  Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
  End Type

    Private Const NIM_ADD = 0
    Private Const NIM_MODIFY = 1
    Private Const NIM_DELETE = 2
    Private Const NIF_MESSAGE = 1
    Private Const NIF_ICON = 2
    Private Const NIF_TIP = 4
    Private Const STI_CALLBACKEVENT = &H201
    Private Const REG_SZ = 1
    Private Const LOCALMACHINE = &H80000002
    Private Const WM_MOUSEMOVE = &H200
    Private Const WM_LBUTTONDOWN = &H201     'Button down
    Private Const WM_LBUTTONUP = &H202       'Button up
    Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
    Private Const WM_RBUTTONDOWN = &H204     'Button down
    Private Const WM_RBUTTONUP = &H205       'Button up
    Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
'*******************************************************************************************************************************************************************************************

' @Function TrayAdd
' -----------------------------
'@Author Roland Ross Hadi
'@Description Add application to the system tray
'@Comments
' Parameter:
'       parentForm - Form
'       Tip - Tip text
Public Sub TrayAdd(parentForm As Form, Tip As String)
    On Error Resume Next
    Dim notIcon As NOTIFYICONDATA
    
    With notIcon
        .cbSize = Len(notIcon)
        .hwnd = parentForm.hwnd
        .uID = vbNull
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uCallbackMessage = STI_CALLBACKEVENT
        .hIcon = parentForm.Icon
        .szTip = Tip & vbNullChar
    End With
    
    Shell_NotifyIconA NIM_ADD, notIcon
End Sub
' Function TrayAdd
' -----------------------------

' @Function TrayModify
' -----------------------------
'@Author Roland Ross Hadi
'@Description Modifies application to the system tray
'@Comments
' Parameter:
'       parentForm - Form
'       Tip - Tip text
Public Sub TrayModify(parentForm As Form, Tip As String)
    On Error Resume Next
    Dim notIcon As NOTIFYICONDATA
    
    With notIcon
        .cbSize = Len(notIcon)
        .hwnd = parentForm.hwnd
        .uID = vbNull
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uCallbackMessage = STI_CALLBACKEVENT
        .hIcon = parentForm.Icon
        .szTip = Tip & vbNullChar
    End With
    
    Shell_NotifyIconA NIM_MODIFY, notIcon
End Sub
' Function TrayModify
' -----------------------------

' @Function TrayDelete
' -----------------------------
'@Author Roland Ross Hadi
'@Description Deletes application to the system tray
'@Comments
' Parameter:
'       parentForm - Form
Public Sub TrayDelete(parentForm As Form)
    On Error Resume Next
    Dim notIcon As NOTIFYICONDATA
    
    With notIcon
      .cbSize = Len(notIcon)
      .hwnd = parentForm.hwnd
      .uID = vbNull
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
      .uCallbackMessage = vbNull
      .hIcon = vbNull
      .szTip = "" & vbNullChar
    End With
    
    Shell_NotifyIconA NIM_DELETE, notIcon
End Sub
' Function TrayDelete
' -----------------------------
