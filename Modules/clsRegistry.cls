VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************************************************************************************************
'*      LIBRARY NAME: clsRegistry.cls
'*      PURPOSE: Compilation of functions that deals with Registry
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

        Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
        Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
        Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
        Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
        Type SECURITY_ATTRIBUTES
          nLength As Long
          lpSecurityDescriptor As Long
          bInheritHandle As Boolean
        End Type
        Private Const REG_OPTION_NON_VOLATILE = 0
        Private Const KEY_ALL_ACCESS = 0
        Private Const ERROR_SUCCESS = 0&
'*******************************************************************************************************************************************************************************************

' @Function MakeStartupReg
' -----------------------------
'@Author Roland Ross Hadi
'@Description Add your application to windows startup registry
'@Comments
' Parameter:
'       strAppTitle - Application Title
Public Sub MakeStartupReg(strAppTitle As String)
        Dim A
        On Error GoTo error
        A = MakeRegFile(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", strAppTitle$, App.path & "\" & App.EXEName & ".exe")
        Exit Sub
error:          MsgBox Err.Description, vbExclamation, "Error"
End Sub
' @Function MakeStartupReg
' -----------------------------

' @Function AddToStartupDir
' -----------------------------
'@Author Roland Ross Hadi
'@Description Add your application to the windows startup folder
'@Comments
' Parameter:
'
Public Sub AddToStartupDir()
        On Error GoTo error
        FileCopy App.path & "\" & App.EXEName & ".EXE", Mid$(App.path, 1, 3) & "WINDOWS\START MENU\PROGRAMS\STARTUP\" & App.EXEName & ".EXE"
        Exit Sub
error:          MsgBox Err.Description, vbExclamation, "Error"
End Sub
' @Function AddToStartupDir
' -----------------------------

' @Function MakeRegFile
' -----------------------------
'@Author Roland Ross Hadi
'@Description For make startup and make registry setting:  Makes the registry setting
'@Comments
' Parameter:
'       hKey - hkey
'       lpszSubKey - lpszSubKey
'       sSetValue - sSetValue
'       sValue - sValue
Public Function MakeRegFile(ByVal hKey As Long, ByVal lpszSubKey As String, ByVal sSetValue As String, ByVal sValue As String) As Boolean
        On Error GoTo error
        Dim phkResult As Long
        Dim lResult As Long
        Dim SA As SECURITY_ATTRIBUTES
        Dim lCreate As Long
        RegCreateKeyEx hKey, lpszSubKey, 0, "", REG_OPTION_NON_VOLATILE, _
        KEY_ALL_ACCESS, SA, phkResult, lCreate
        lResult = RegSetValueEx(phkResult, sSetValue, 0, 1, sValue, _
        CLng(Len(sValue) + 1))
        RegCloseKey phkResult
        MakeRegFile = (lResult = ERROR_SUCCESS)
        Exit Function
error:
        MakeRegFile = False
End Function
' @Function MakeRegFile
' -----------------------------


