VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAPIfunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************************************************************************************************
'*      LIBRARY NAME: clsAPIfunctions.cls
'*      PURPOSE: Compilation of functions that deals with API functions
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
        
        Private Declare Function FindExecutableA Lib "shell32.dll" _
           (ByVal lpFile As String, ByVal lpdirectory As _
           String, ByVal lpResult As String) As Long

        Private Declare Function CloseHandle Lib "KERNEL32" (hObject As Long) As Boolean

        Private Declare Function CreateProcessA Lib "KERNEL32" _
            (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
            ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
            ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
            ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
            lpStartupInfo As STARTUPINFO, _
            lpProcessInformation As PROCESS_INFORMATION) As Long

        Private Declare Function GetUserNameA Lib "advapi32.dll" _
           (ByVal lpBuffer As String, nSize As Long) As Long

        Private Declare Function GetComputerNameA Lib "KERNEL32" _
           (ByVal lpBuffer As String, nSize As Long) As Long

        Private Declare Function ExitWindowsEx Lib "user32" (ByVal _
           uFlags As Long, ByVal dwReserved As Long) As Long

          Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

        Private Const EXIT_LOGOFF = 0
        Private Const EXIT_SHUTDOWN = 1
        Private Const EXIT_REBOOT = 2

        Private Const MAX_FILENAME_LEN = 256

        Private Declare Function WaitForSingleObject Lib "KERNEL32" _
            (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

        Private Const NORMAL_PRIORITY_CLASS = &H20&
        Private Const INFINITE = -1&
        Private Const SYNCHRONIZE = &H100000

        Private Type SHFILEOPSTRUCT
                hwnd As Long
                wFunc As Long
                pFrom As String
                pTo As String
                fFlags As Integer
                fAborted As Boolean
                hNameMaps As Long
                sProgress As String
        End Type

        Private Const FO_DELETE = &H3
        Private Const FOF_ALLOWUNDO = &H40
        Private Const FOF_SILENT = &H4
        Private Const FOF_NOCONFIRMATION = &H10

        Private Declare Function SHFileOperation Lib "shell32.dll" Alias _
           "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

        Private Type STARTUPINFO
            cb As Long
            lpReserved As String
            lpDesktop As String
            lpTitle As String
            dwX As Long
            dwY As Long
            dwXSize As Long
            dwYSize As Long
            dwXCountChars As Long
            dwYCountChars As Long
            dwFillAttribute As Long
            dwFlags As Long
            wShowWindow As Integer
            cbReserved2 As Integer
            lpReserved2 As Long
            hStdInput As Long
            hStdOutput As Long
            hStdError As Long
        End Type

        Private Type PROCESS_INFORMATION
            hProcess As Long
            hThread As Long
            dwProcessId As Long
            dwThreadID As Long
        End Type
        
        Private Const EW_REBOOTSYSTEM = &H43
        Private Const EW_RESTARTWINDOWS = &H42
        Private Const EW_EXITWINDOWS = 0
        
        '// Windows Registry Messages
        Private Const REG_SZ As Long = 1
        Private Const REG_DWORD As Long = 4
        Private Const HKEY_CLASSES_ROOT = &H80000000
        Private Const HKEY_CURRENT_USER = &H80000001
        Private Const HKEY_LOCAL_MACHINE = &H80000002
        Private Const HKEY_USERS = &H80000003
        
        '// Windows Error Messages
        Private Const ERROR_NONE = 0
        Private Const ERROR_BADDB = 1
        Private Const ERROR_BADKEY = 2
        Private Const ERROR_CANTOPEN = 3
        Private Const ERROR_CANTREAD = 4
        Private Const ERROR_CANTWRITE = 5
        Private Const ERROR_OUTOFMEMORY = 6
        Private Const ERROR_INVALID_PARAMETER = 7
        Private Const ERROR_ACCESS_DENIED = 8
        Private Const ERROR_INVALID_PARAMETERS = 87
        Private Const ERROR_NO_MORE_ITEMS = 259
        
        '// Windows Security Messages
        Private Const KEY_ALL_ACCESS = &H3F
        Private Const REG_OPTION_NON_VOLATILE = 0
        
        '// Windows Registry API calls
        Private Declare Function RegCloseKey Lib "advapi32.dll" _
         (ByVal hKey As Long) As Long
        
        Private Declare Function RegCreateKeyEx _
          Lib "advapi32.dll" Alias "RegCreateKeyExA" _
         (ByVal hKey As Long, _
          ByVal lpSubKey As String, _
          ByVal Reserved As Long, _
          ByVal lpClass As String, _
          ByVal dwOptions As Long, _
          ByVal samDesired As Long, _
          ByVal lpSecurityAttributes As Long, _
          phkResult As Long, _
          lpdwDisposition As Long) As Long
        
        Private Declare Function RegOpenKeyEx _
          Lib "advapi32.dll" Alias "RegOpenKeyExA" _
         (ByVal hKey As Long, _
          ByVal lpSubKey As String, _
          ByVal ulOptions As Long, _
          ByVal samDesired As Long, _
          phkResult As Long) As Long
        
        Private Declare Function RegSetValueExString _
          Lib "advapi32.dll" Alias "RegSetValueExA" _
         (ByVal hKey As Long, _
          ByVal lpValueName As String, _
          ByVal Reserved As Long, _
          ByVal dwType As Long, _
          ByVal lpValue As String, _
          ByVal cbData As Long) As Long
        
        Private Declare Function RegSetValueExLong _
          Lib "advapi32.dll" Alias "RegSetValueExA" _
         (ByVal hKey As Long, _
          ByVal lpValueName As String, _
          ByVal Reserved As Long, _
          ByVal dwType As Long, _
          lpValue As Long, _
          ByVal cbData As Long) As Long
          
          Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

          Private Const SPI_SETDESKWALLPAPER = 20
          
          ' These are common sounds available from the registry
            Private Const elDefault = ".Default"
            Private Const elGPF = "AppGPFault"
            Private Const elClose = "Close"
            Private Const elEmptyRecycleBin = "EmptyRecycleBin"
            Private Const elMailBeep = "MailBeep"
            Private Const elMaximize = "Maximize"
            Private Const elMenuCommand = "MenuCommand"
            Private Const elMenuPopUp = "MenuPopup"
            Private Const elMinimize = "Minimize"
            Private Const elOpen = "Open"
            Private Const elRestoreDown = "RestoreDown"
            Private Const elRestoreUp = "RestoreUp"
            Private Const elSystemAsterisk = "SystemAsterisk"
            Private Const elSystemExclaimation = "SystemExclaimation"
            Private Const elSystemExit = "SystemExit"
            Private Const elSystemHand = "SystemHand"
            Private Const elSystemQuestion = "SystemQuestion"
            Private Const elSystemStart = "SystemStart"
            Private Declare Function PlaySound& Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long)
' hModule is only used if SND_RESOURCE& is set and represents
' an HINSTANCE handle.  This example doesn't support playing
' from a resource file.

'*-------------------------------------*
'* Playsound flags: store in dwFlags   *
'*-------------------------------------*
' lpszName points to a registry entry
' Do not use SND_RESOURSE or SND_FILENAME
Private Const SND_ALIAS& = &H10000
' Playsound returns immediately
' Do not use SND_SYNC
Private Const SND_ASYNC& = &H1
' The name of a wave file.
' Do not use with SND_RESOURCE or SND_ALIAS
Private Const SND_FILENAME& = &H20000
' Unless used, the default beep will
' play if the specified resource is missing
Private Const SND_NODEFAULT& = &H2
' Fail the call & do not wait for
' a sound device if it is otherwise unavailable
Private Const SND_NOWAIT& = &H2000
' Use a resource file as the source.
' Do not use with SND_ALIAS or SND_FILENAME
Private Const SND_RESOURCE& = &H40004
' Playsound will not return until the
' specified sound has played.  Do not
' use with SND_ASYNC
Private Const SND_SYNC& = &H0

Public Enum enSound_Source
    ssFile = SND_FILENAME&
    ssRegistry = SND_ALIAS&
End Enum

'*******************************************************************************************************************************************************************************************

' @Function ExitWindows
' -----------------------------
'@Author Roland Ross Hadi
'@Description Exit Windows
'@Comments
' Parameter:
'       uFlags - uFlags
Public Sub ExitWindows(ByVal uFlags As Long)
   Call ExitWindowsEx(uFlags, 0)
End Sub
' Function ExitWindows
' -----------------------------

' @Function GetUserName
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the current-username
'@Comments
' Parameter:
'
Public Function GetUserName() As String
   Dim UserName As String * 255

   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
' Function GetUserName
' -----------------------------

' @Function GetComputerName
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the computer's name
'@Comments
' Parameter:
'
Public Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
' @Function GetComputerName
' -----------------------------

' @Function Pause
' -----------------------------
'@Author Roland Ross Hadi
'@Description Sleeps for a given number of seconds.
'@Comments
' Parameter:
'       sngSeconds - Seconds (Time)
Public Sub Pause(ByVal sngSeconds As Single)
   Call Sleep(Int(sngSeconds * 1000#))
End Sub
' Function Pause
' -----------------------------

' @Function ShellDelete
' -----------------------------
'@Author Roland Ross Hadi
'@Description Deletes a single file, or an array of files to the trashcan.
'@Comments
' Parameter:
'       vntFileName() - Filename
Public Function ShellDelete(ParamArray vntFileName() As Variant) As Boolean
   Dim i As Integer
   Dim sFileNames As String
   Dim SHFileOp As SHFILEOPSTRUCT

   For i = LBound(vntFileName) To UBound(vntFileName)
      sFileNames = sFileNames & vntFileName(i) & vbNullChar
   Next
        
   sFileNames = sFileNames & vbNullChar

   With SHFileOp
      .wFunc = FO_DELETE
      .pFrom = sFileNames
      .fFlags = FOF_ALLOWUNDO + FOF_SILENT + FOF_NOCONFIRMATION
   End With

   i = SHFileOperation(SHFileOp)
   
   If i = 0 Then
      ShellDelete = True
   Else
      ShellDelete = False
   End If
End Function
' Function ShellDelete
' -----------------------------

' @Function ShellDelete
' -----------------------------
'@Author Roland Ross Hadi
'@Description Runs a command as the Shell command does but waits for the command
'@Comments
' Parameter:
'       cCommandLine() - Command line text
Public Function ShellWait(cCommandLine As String) As Boolean
    Dim NameOfProc As PROCESS_INFORMATION
    Dim NameStart As STARTUPINFO
    Dim i As Long

    NameStart.cb = Len(NameStart)
    i = CreateProcessA(0&, cCommandLine, 0&, 0&, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, NameStart, NameOfProc)
   
    If i <> 0 Then
       Call WaitForSingleObject(NameOfProc.hProcess, INFINITE)
       Call CloseHandle(NameOfProc.hProcess)
       ShellWait = True
    Else
       ShellWait = False
    End If
    
End Function
' Function ShellDelete
' -----------------------------

' @Function ExecuteWait
' -----------------------------
'@Author Roland Ross Hadi
'@Description As the Execute function but waits for the process to finish before returning
'@Comments
' Parameter:
'       strExe - Executable file
'       strParam - Exe parameter
Public Function ExecuteWait(strExe As String, Optional strParam As Variant) As Boolean
   Dim s2 As String
   
   s2 = FindExecutable(strExe)
   
   If s2 <> "" Then
      ExecuteWait = ShellWait(s2 & _
         IIf(IsMissing(strParam), " ", " " & CStr(strParam) & " ") & strExe)
   Else
      ExecuteWait = False
   End If
End Function
' Function ExecuteWait
' -----------------------------

' @Function FindExecutable
' -----------------------------
'@Author Roland Ross Hadi
'@Description Finds the executable associated with a file
'@Comments
' Parameter:
'       strExe - Executable file
Public Function FindExecutable(strExe As String) As String
   Dim i As Integer
   Dim s2 As String
   
   s2 = String(MAX_FILENAME_LEN, 32) & Chr$(0)
   
   i = FindExecutableA(strExe & Chr$(0), vbNullString, s2)
   
   If i > 32 Then
      FindExecutable = Left$(s2, InStr(s2, Chr$(0)) - 1)
   Else
      FindExecutable = ""
   End If
End Function
' Function FindExecutable
' -----------------------------

' @Function RestartWindows
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will restart windows
'@Comments
' Parameter:
'
Function RestartWindows()
        On Error GoTo error
        Dim RetVal As Integer
        RetVal = ExitWindowsEx(EW_RESTARTWINDOWS, 0)
        RestartWindows = RetVal
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function RestartWindows
' -----------------------------

' @Function DoExitWindows
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will exit windows
'@Comments
' Parameter:
'
Function DoExitWindows()
        On Error GoTo error
        Dim RetVal As Integer
        RetVal = ExitWindowsEx(EW_EXITWINDOWS, 0)
        DoExitWindows = RetVal
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function DoExitWindows
' -----------------------------

' @Function RebootComputer
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will reboot the computer
'@Comments
' Parameter:
'
Function RebootComputer()
        On Error GoTo error
        Dim RetVal As Integer
        RetVal = ExitWindowsEx(EW_REBOOTSYSTEM, 0)
        RebootComputer = RetVal
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function RebootComputer
' -----------------------------

' @Function GetAppVersion
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will retrieve the current version of your application
'@Comments
' Parameter:
'
Function GetAppVersion()
        On Error GoTo error
        GetAppVersion = App.Major & "." & App.Minor & "." & App.Revision
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function GetAppVersion
' -----------------------------

' @Function GetAppName
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will get the application's .exe name
'@Comments
' Parameter:
'       ShowEXE - True/False
Function GetAppName(ShowEXE As Boolean)
        On Error GoTo error
        GetAppName = App.EXEName
        If ShowEXE = True Then
        GetAppName = GetAppName & ".exe"
        End If
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function GetAppName
' -----------------------------

' @Function GetAppPath
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will get the application's current path
'@Comments
' Parameter:
'
Function GetAppPath()
        On Error GoTo error
        GetAppPath = App.path
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function GetAppPath
' -----------------------------

' @Function GetAppDescription
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will get the application's file description
'@Comments
' Parameter:
'
Function GetAppDescription()
        On Error GoTo error
        GetAppDescription = App.FileDescription
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function GetAppDescription
' -----------------------------

' @Function GetAppCopyRight
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will get the application's copyright
'@Comments
' Parameter:
'
Function GetAppCopyRight()
        On Error GoTo error
        GetAppCopyRight = App.LegalCopyright
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function GetAppCopyRight
' -----------------------------

' @Function GetAppComment
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will get the application's comment
'@Comments
' Parameter:
'
Function GetAppComment()
        On Error GoTo error
        GetAppComment = App.Comments
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function GetAppComment
' -----------------------------

' @Function GetAppTitle
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will get the application's title
'@Comments
' Parameter:
'
Function GetAppTitle()
        On Error GoTo error
        GetAppTitle = App.Title
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function GetAppTitle
' -----------------------------

' @Function GetAppCompanyName
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will get the application's company name
'@Comments
' Parameter:
'
Function GetAppCompanyName()
        On Error GoTo error
        GetAppCompanyName = App.CompanyName
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function GetAppCompanyName
' -----------------------------

' @Function GetAppProductName
' -----------------------------
'@Author Roland Ross Hadi
'@Description This will get the application's product name
'@Comments
' Parameter:
'
Function GetAppProductName()
        On Error GoTo error
        GetAppProductName = App.ProductName
        Exit Function
error:          MsgBox Err.Description, vbExclamation, "Error"
End Function
' Function GetAppProductName
' -----------------------------

' @Function CreateAssociation
' -----------------------------
'@Author Roland Ross Hadi
'@Description Call CreateAssociation() to create file association
'@Comments
' Parameter:
'   strFileExtension - File extension
'   strDescription - File extension description
'   strAppName - Application name
Public Sub CreateAssociation(strFileExtension As String, strDescription As String, strAppName As String)
  Dim sPath As String
  CreateNewKey strFileExtension, HKEY_CLASSES_ROOT
  SetKeyValue strFileExtension, "", strDescription, REG_SZ
  CreateNewKey strDescription & "\shell\Open in " & strAppName & "\command", HKEY_CLASSES_ROOT
  CreateNewKey strDescription & "\DefaultIcon", HKEY_CLASSES_ROOT
  SetKeyValue strDescription & "\DefaultIcon", "", App.path & "\" & App.EXEName & ".exe,0", REG_SZ
  SetKeyValue strDescription, "", strAppName & " file", REG_SZ
  sPath = App.path & "\" & App.EXEName & ".exe %1"

  SetKeyValue strDescription & "\shell\Open in " & strAppName & "\command", "", sPath, REG_SZ
End Sub
' Function CreateAssociation
' -----------------------------

' @Function SetValueEx
' -----------------------------
'@Author Roland Ross Hadi
'@Description Set new key value in the registry
'@Comments
' Parameter:
'
Public Function SetValueEx(ByVal hKey As Long, _
 sValueName As String, lType As Long, _
 vValue As Variant) As Long

  Dim nValue As Long
  Dim sValue As String

  Select Case lType
    Case REG_SZ
      sValue = vValue & Chr$(0)
      SetValueEx = RegSetValueExString(hKey, _
        sValueName, 0&, lType, sValue, Len(sValue))

    Case REG_DWORD
      nValue = vValue
      SetValueEx = RegSetValueExLong(hKey, sValueName, _
        0&, lType, nValue, 4)
  End Select
End Function
' Function SetValueEx
' -----------------------------

' @Function CreateNewKey
' -----------------------------
'@Author Roland Ross Hadi
'@Description Creates new key in the registry
'@Comments
' Parameter:
'     strNewKeyName - Key name
Public Sub CreateNewKey(strNewKeyName As String, _
  lPredefinedKey As Long)

  '// handle to the new key
  Dim hKey As Long

  '// result of the RegCreateKeyEx function
  Dim r As Long

  r = RegCreateKeyEx(lPredefinedKey, strNewKeyName, 0&, _
    vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, r)

  Call RegCloseKey(hKey)
End Sub
' Function CreateNewKey
' -----------------------------

' @Function SetKeyValue
' -----------------------------
'@Author Roland Ross Hadi
'@Description Set new key value in the registry
'@Comments
' Parameter:
'     strNewKeyName - Key name
'     strValueName - Key value name
'     varValueSetting - Key value setting
'     strNewKeyName - New key name
Public Sub SetKeyValue(strKeyName As String, strValueName As String, _
varValueSetting As Variant, lngValueType As Long)

  '// result of the SetValueEx function
  Dim r As Long

  '// handle of opened key
  Dim hKey As Long

  '// open the specified key
  r = RegOpenKeyEx(HKEY_CLASSES_ROOT, strKeyName, 0, _
    KEY_ALL_ACCESS, hKey)

  r = SetValueEx(hKey, strValueName, lngValueType, varValueSetting)

  Call RegCloseKey(hKey)
End Sub
' Function SetKeyValue
' -----------------------------

' @Function ChangeWallPaper
' -----------------------------
'@Author Roland Ross Hadi
'@Description Creates new key in the registry
'@Comments
' Parameter:
'     strNewKeyName - Key name
Public Sub ChangeWallPaper(strNewPictureFileName As String)
    Dim ChangeWP
    ChangeWP = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, strNewPictureFileName, 0)
End Sub
' Function ChangeWallPaper
' -----------------------------


' Plays sounds from the registry or a disk file
' Doesn't care if the file is missing
Public Function EZPlay(ssname As String) As Boolean
On Error Resume Next
    If PlaySound(ssname, 0&, SND_ASYNC + SND_NODEFAULT) Then
        EZPlay = True
    Else
        EZPlay = False
    End If
End Function
