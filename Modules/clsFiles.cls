VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************************************************************************************************
'*      LIBRARY NAME: clsFilesAndDirectories.cls
'*      PURPOSE: Compilation of functions that deals with Files and Directories
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
        Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
        Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
        Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
        
        Private Const MAX_FILENAME_LEN = 256
        
        ' File and Disk functions.
        Private Const DRIVE_CDROM = 5
        Private Const DRIVE_FIXED = 3
        Private Const DRIVE_RAMDISK = 6
        Private Const DRIVE_REMOTE = 4
        Private Const DRIVE_REMOVABLE = 2
        Private Const DRIVE_UNKNOWN = 0    'Unknown, or unable to be determined.
        
        Private Declare Function GetDriveTypeA Lib "kernel32" (ByVal nDrive As String) As Long
        
        Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" _
           (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, _
            ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
            lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
            ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
        
        Private Declare Function GetWindowsDirectoryA Lib "kernel32" _
           (ByVal lpBuffer As String, ByVal nSize As Long) As Long
           
        Private Declare Function GetTempPathA Lib "kernel32" _
           (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
        
        Private Const UNIQUE_NAME = &H0
        
        Private Declare Function GetTempFileNameA Lib "kernel32" (ByVal _
           lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique _
           As Long, ByVal lpTempFileName As String) As Long
           
        Private Declare Function GetSystemDirectoryA Lib "kernel32" _
           (ByVal lpBuffer As String, ByVal nSize As Long) As Long
        
        Private Declare Function GetShortPathNameA Lib "kernel32" _
           (ByVal lpszLongPath As String, ByVal lpszShortPath _
           As String, ByVal cchBuffer As Long) As Long
           
        Private Type SHFILEOPSTRUCT
                hWnd As Long
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
        
        Private Const NORMAL_PRIORITY_CLASS = &H20&
        Private Const INFINITE = -1&
        Private Const SYNCHRONIZE = &H100000
        
        Private Declare Function MakeSureDirectoryPathExists Lib _
        "IMAGEHLP.DLL" (ByVal DirPath As String) As Long
        
        Private Declare Function SetVolumeLabelA Lib "kernel32" _
        (ByVal lpRootPathName As String, _
        ByVal lpVolumeName As String) As Long
        
        Private Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias _
            "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
        Private Declare Function GetPrivateProfileStringKeys& Lib "kernel32" Alias _
            "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
        Private Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias _
            "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
        Private Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias _
            "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)
        Private Declare Function WritePrivateProfileStringToDeleteKey& Lib "kernel32" Alias _
            "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String)
        
'        Private strMySystemFile As String
'        Private strFileName As String
'        Private strSection As String
        Private Const BUFF_SIZ As Long = 9160
        Private Const READ_BUFF As Long = 1000000
        Public LoadedXMLData As String
'*******************************************************************************************************************************************************************************************


' @Function FileAppend
' -----------------------------
'@Author Roland Ross Hadi
'@Description Append <strText> to <strFileName>
'@Comments
' Parameter:
'       strFileName - Filename to be appended with the strText
'       strText - text to be appended to the file
Public Function FileAppend(strFileName As String, strText As String)
    Dim ff As Integer
    On Error Resume Next
    ff = FreeFile
    Open strFileName For Append As ff
        Print #ff, CStr(strText)
    Close ff
End Function
' Function FileAppend
' -----------------------------

' @Function FileWrite
' -----------------------------
'@Author Roland Ross Hadi
'@Description Write <strText> to <strFileName>
'@Comments
' Parameter:
'       strFileName - Filename to be created with the strText
'       strText - text to be wriiten to the file
Public Function FileWrite(strFileName As String, strText As String)
    Dim ff As Integer
    On Error Resume Next
    ff = FreeFile
    Kill strFileName
    Open strFileName For Binary As ff
        Put #ff, , CStr(strText)
    Close ff
End Function
' Function FileWrite
' -----------------------------

' @Function FileDelete
' -----------------------------
'@Author Roland Ross Hadi
'@Description Deletes <strFileName>
'@Comments
' Parameter:
'       strFileName - Filename to be deleted
Public Function FileDelete(strFileName)
    On Error Resume Next
    Kill strFileName
End Function
' Function FileDelete
' -----------------------------

' @Function FileData
' -----------------------------
'@Author Roland Ross Hadi
'@Description Extract text from <strFileName>
'@Comments
' Parameter:
'       strFileName - Filename to be extacted
Public Function FileData(strFileName As String) As String
    Dim ff As Integer
    Dim dat As String
    ff = FreeFile
    Open strFileName For Binary As ff
        dat = Space$(LOF(ff))
        Get #ff, , dat
    Close ff
    FileData = dat
End Function
' Function FileData
' -----------------------------

' @Function FileCopy
' -----------------------------
'@Author Roland Ross Hadi
'@Description Copy <strFile1> to <strFile2>
'@Comments
' Parameter:
'       strFile1 - Filename to be copied
'       strFile2 - Destination strFileName
Public Function FileCopy(strFile1 As String, strFile2 As String)
    FileCopy = CopyFile(CStr(strFile1), CStr(strFile2), False)
End Function
' Function FileCopy
' -----------------------------

' @Function FileMove
' -----------------------------
'@Author Roland Ross Hadi
'@Description Move <strFile1> to <strFile2>
'@Comments
' Parameter:
'       strFile1 - Filename to be moved
'       strFile2 - Destination strFileName
Public Function FileMove(strFile1 As String, strFile2 As String)
    FileMove = MoveFile(CStr(strFile1), CStr(strFile2))
End Function
' Function FileMove
' -----------------------------

' @Function FileLen
' -----------------------------
'@Author Roland Ross Hadi
'@Description Gets the Length of <strFileName>
'@Comments
' Parameter:
'       strFileName - Filename to be processed
Public Function FileLen(strFileName As String)
    Dim ff As Integer
    ff = FreeFile
    Open strFileName For Binary As ff
        FileLen = LOF(ff)
    Close ff
End Function
' Function FileLen
' -----------------------------

' @Function FileAttributes
' -----------------------------
'@Author Roland Ross Hadi
'@Description Gets the strFileName Attributes of <strFileName>
'@Comments
' Parameter:
'       strFileName - Filename to be processed
Public Function FileAttributes(strFileName As String)
    FileAttributes = GetFileAttributes(CStr(strFileName))
End Function
' Function FileAttributes
' -----------------------------

' @Function GetFileName
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the file portion of a file + pathname
'@Comments
' Parameter:
'       strFileName - Filename to be processed
Public Function GetFileName(strFileName As String) As String
   Dim i As Integer
   Dim j As Integer
   
   i = 0
   j = 0
   
   i = InStr(strFileName, "\")
   Do While i <> 0
      j = i
      i = InStr(j + 1, strFileName, "\")
   Loop
   
   If j = 0 Then
      GetFileName = ""
   Else
      GetFileName = Right$(strFileName, Len(strFileName) - j)
   End If
End Function
' Function GetFileName
' -----------------------------

' @Function GetFilePath
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the path portion of a file + pathname
'@Comments
' Parameter:
'       strFileName - Filename to be processed
Public Function GetFilePath(ByVal strFileName As String) As String
   Dim i As Integer
   Dim j As Integer
   
   i = 0
   j = 0
   
   i = InStr(strFileName, "\")
   Do While i <> 0
      j = i
      i = InStr(j + 1, strFileName, "\")
   Loop
   
   If j = 0 Then
      GetFilePath = ""
   Else
      GetFilePath = Left$(strFileName, j)
   End If
End Function
' Function GetFilePath
' -----------------------------

' @Function GetSerialNumber
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns a volume's serial number
'@Comments
' Parameter:
'       strDrive - Drive to be processed
Public Function GetSerialNumber(strDrive As String) As Long
   Dim ser As Long
   Dim s As String * MAX_FILENAME_LEN
   Dim s2 As String * MAX_FILENAME_LEN
   Dim i As Long
   Dim j As Long
   
   Call GetVolumeInformation(strDrive + ":\" & Chr$(0), s, MAX_FILENAME_LEN, ser, i, j, s2, MAX_FILENAME_LEN)
   GetSerialNumber = ser
End Function
' Function GetSerialNumber
' -----------------------------

' @Function GetShortPathName
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the short file name as in "untitl~1.txt' for "untitledfile.txt"
'@Comments
' Parameter:
'       strFileName - Filename to be processed
Public Function GetShortPathName(strFileName As String) As String
   Dim s As String
   Dim i As Long
   
   i = Len(strFileName) + 1
   s = String(i, 0)
   GetShortPathNameA strFileName, s, i
   
   GetShortPathName = Left$(s, InStr(s, Chr$(0)) - 1)
End Function
' Function GetShortPathName
' -----------------------------

' @Function GetVolumeName
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns a volume's name
'@Comments
' Parameter:
'       strDrive - Drive to be processed
Public Function GetVolumeName(strDrive As String) As String
   Dim ser As Long
   Dim s As String * MAX_FILENAME_LEN
   Dim s2 As String * MAX_FILENAME_LEN
   Dim i As Long
   Dim j As Long
   
   Call GetVolumeInformation(strDrive + ":\" & Chr$(0), s, MAX_FILENAME_LEN, ser, i, j, s2, MAX_FILENAME_LEN)
   GetVolumeName = Left$(s, InStr(s, Chr$(0)) - 1)
End Function
' Function GetVolumeName
' -----------------------------

' @Function SetVolumeName
' -----------------------------
'@Author Roland Ross Hadi
'@Description Sets the volume name.  Returns true on success, false on failure.
'@Comments
' Parameter:
'       strDrive - Drive to be processed
'       strNewVolumeName - New volume name
Public Function SetVolumeName(strDrive As String, strNewVolumeName As String) As Boolean
   Dim i As Long
   
   i = SetVolumeLabelA(strDrive + ":\" & Chr$(0), strNewVolumeName & Chr$(0))
   
   SetVolumeName = IIf(i = 0, False, True)
End Function
' Function SetVolumeName
' -----------------------------

' @Function GetSystemDirectory
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the system directory.
'@Comments
' Parameter:
'
Public Function GetSystemDirectory() As String
   Dim s As String
   Dim i As Integer
   i = GetSystemDirectoryA("", 0)
   s = Space(i)
   Call GetSystemDirectoryA(s, i)
   GetSystemDirectory = AddBackslash(Left$(s, i - 1))
End Function
' Function GetSystemDirectory
' -----------------------------

' @Function GetTempFileName
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns a unique tempfile name.
'@Comments
' Parameter:
'
Public Function GetTempFileName() As String
   Dim s As String
   Dim s2 As String
   
   s2 = GetTempPath
   s = Space(Len(s2) + MAX_FILENAME_LEN)
   Call GetTempFileNameA(s2, App.EXEName, UNIQUE_NAME, s)
   GetTempFileName = Left$(s, InStr(s, Chr$(0)) - 1)
End Function
' Function GetTempFileName
' -----------------------------
 
' @Function GetTempPath
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the path to the temp directory.
'@Comments
' Parameter:
'
Public Function GetTempPath() As String
   Dim s As String
   Dim i As Integer
   i = GetTempPathA(0, "")
   s = Space(i)
   Call GetTempPathA(i, s)
   GetTempPath = AddBackslash(Left$(s, i - 1))
End Function
' Function GetTempPath
' -----------------------------

' @Function GetWindowsDirectory
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the windows directory.
'@Comments
' Parameter:
'
Public Function GetWindowsDirectory() As String
   Dim s As String
   Dim i As Integer
   i = GetWindowsDirectoryA("", 0)
   s = Space(i)
   Call GetWindowsDirectoryA(s, i)
   GetWindowsDirectory = AddBackslash(Left$(s, i - 1))
End Function
' Function GetWindowsDirectory
' -----------------------------

' @Function sDriveType
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the drive type if possible.
'@Comments
' Parameter:
'       strDrive - Drive to be processed
Public Function sDriveType(strDrive As String) As String
Dim lRet As Long

    lRet = GetDriveTypeA(strDrive & ":\")
    Select Case lRet
        Case 0
            'sDriveType = "Cannot be determined!"
            sDriveType = "Unknown"
            
        Case 1
            'sDriveType = "The root directory does not exist!"
            sDriveType = "Unknown"
        Case DRIVE_CDROM:
            sDriveType = "CD-ROM Drive"
            
        Case DRIVE_REMOVABLE:
            sDriveType = "Removable Drive"
            
        Case DRIVE_FIXED:
            sDriveType = "Fixed Drive"
            
        Case DRIVE_REMOTE:
            sDriveType = "Remote Drive"
        End Select
End Function
' Function sDriveType
' -----------------------------

' @Function GetDriveType
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns the drive type if possible.
'@Comments
' Parameter:
'       strDrive - Drive to be processed
Public Function GetDriveType(strDrive As String) As Long
  Dim lRet As Long
  lRet = GetDriveTypeA(strDrive & ":\")
  
  If lRet = 1 Then
     lRet = 0
  End If

  GetDriveType = lRet
End Function
' @Function GetDriveType
' -----------------------------

' @Function FileExists
' -----------------------------
'@Author Roland Ross Hadi
'@Description Determines whether the specified file exists
'@Comments
' Parameter:
'       strPathName - Drive to be processed
Public Function FileExists(ByVal strPathName As String) As Boolean
        If Dir$(strPathName) <> "" Then
            FileExists = True
        Else
            FileExists = False
        End If
End Function
' Function FileExists
' -----------------------------

' @Function ConvertFileSize
' -----------------------------
'@Author Roland Ross Hadi
'@Description Convert byte sizes
'@Comments
' Parameter:
'       lngFileSizeInBytes - file size in bytes
Public Function ConvertFileSize(lngFileSizeInBytes As Long) As String
        If lngFileSizeInBytes < 1024 Then
        ConvertFileSize = lngFileSizeInBytes & " Bytes"
        Exit Function
        End If

        If lngFileSizeInBytes < 1048576 Then
        ConvertFileSize = Format(lngFileSizeInBytes / 1024, "#####.0#") & " Kb"
        Exit Function
        Else
        ConvertFileSize = Format((lngFileSizeInBytes / 1024) / 1024, "#######.0#") & " Mb"
        End If
End Function
' Function ConvertFileSize
' -----------------------------

' @Function MakeFolder
' -----------------------------
'@Author Roland Ross Hadi
'@Description Create a new folder
'@Comments
' Parameter:
'       strPathToCreate - folder to be created
Public Function MakeFolder(ByVal strPathToCreate As String) _
  As Boolean
    Dim sSomePath As String
    Dim bAns As Boolean
    
   sSomePath = strPathToCreate
    
    If CreatePath(sSomePath) = True Then
        bAns = True
    Else
       bAns = False
    End If
MakeFolder = bAns
End Function
' @Function MakeFolder
' -----------------------------

' @Function CreatePath
' -----------------------------
'@Author Roland Ross Hadi
'@Description Creates a new file path
'@Comments
' Parameter:
'       strNewPath - path to be created
Private Function CreatePath(strNewPath) As Boolean
    Dim sPath As String
    'Add a trailing slash if none
    sPath = strNewPath & IIf(Right$(strNewPath, 1) = "\", "", "\")

    'Call API
    If MakeSureDirectoryPathExists(sPath) <> 0 Then
        'No errors, return True
        CreatePath = True
    End If
End Function
' Function CreatePath
' -----------------------------

' @Function WriteNewFile
' -----------------------------
'@Author Roland Ross Hadi
'@Description Writes a new file
'@Comments
' Parameter:
'       strPath - path to be created
'       strValue - value to be entered in the file
Public Function WriteNewFile(strPath As String, strValue As String)
    Open strPath For Output As #1
    Print #1, strValue
    Close #1
End Function
' Function WriteNewFile
' -----------------------------

' @Function WriteToEndOfFile
' -----------------------------
'@Author Roland Ross Hadi
'@Description Writes to the end of a file
'@Comments
' Parameter:
'       strPath - path to be created
'       strValue - value to be entered in the file
Public Function WriteToEndOfFile(strPath As String, strValue As String)
    If FileExists(strPath) = True Then
      Open strPath For Append As #1
      Print #1, strValue
      Close #1
    Else
      WriteNewFile strPath, strValue
    End If
End Function
' Function WriteToEndOfFile
' -----------------------------

' @Function ReadFromFile
' -----------------------------
'@Author Roland Ross Hadi
'@Description Reads data form a file
'@Comments
' Parameter:
'       strPath - path to be created
Public Function ReadFromFile(strPath As String) As String

    Dim ReturnValue As String
    Dim Value As String

    On Error GoTo Err1
    
    Open strPath For Input As #1
    While Not EOF(1)
    Line Input #1, Value
    ReturnValue = ReturnValue & Value & vbNewLine
    Wend
    Close #1

    ReadFromFile = ReturnValue
    Exit Function
Err1:
    ReadFromFile = "No Data Found"

End Function
' Function ReadFromFile
' -----------------------------

' @Function ReadFromFileToArray
' -----------------------------
'@Author Roland Ross Hadi
'@Description Reads data form a file
'@Comments
' Parameter:
'       strPath - path to be created
Public Function ReadFromFileToArray(strPath As String)

    Dim ReturnValue() As String
    Dim Value As String

    On Error GoTo Err1
    ReDim ReturnValue(0)
    Open strPath For Input As #1
    While Not EOF(1)
    Line Input #1, Value
        ReturnValue(UBound(ReturnValue)) = Value
        ReDim Preserve ReturnValue(UBound(ReturnValue) + 1)
    Wend
    Close #1

    ReadFromFileToArray = ReturnValue
    Exit Function
Err1:
    ReadFromFileToArray = "No Data Found"

End Function
' Function ReadFromFile
' -----------------------------

' @Function ReadFromFile
' -----------------------------
'@Author Roland Ross Hadi
'@Description Reads data form a file
'@Comments
' Parameter:
'       strPath - path to be created
Public Function ReadFromFile_FAILED(strPath As String) As String

    Dim ReturnValue As String
    Dim Value As String

    On Error GoTo Err1
    
    Open strPath For Input As #1
    While Not EOF(1)
        Line Input #1, Value
        If InStr(1, Value, "(FAILED)") <> 0 Then ReturnValue = ReturnValue & Value & vbNewLine
    Wend
    Close #1

    ReadFromFile_FAILED = ReturnValue
    Exit Function
Err1:
    ReadFromFile_FAILED = "No Data Found"

End Function
' Function ReadFromFile
' -----------------------------

' @Function WriteKeyToFile
' -----------------------------
'@Author Roland Ross Hadi
'@Description Write key data to a file
'@Comments
' Parameter:
'       strFileSection - string used as a file subheader for every section
'       strKey - string used as key to write the value to file ( ex: UserName)
'       strValue - the string value to write to file ( ex: Password)
' Returns:  -1 string to write is more than 255 chars
'            0 system write failure
'            1 write to file succesful
Public Function WriteKeyToFile(strFileSection As String, strKey As String, strValue As String) As Long
    If Len(strKey) > READ_BUFF Or Len(strValue) > READ_BUFF Then
        MsgBox "Can't write more than " & READ_BUFF & " characters for key or value."
        WriteKeyToFile = -1
        Exit Function
    End If
    WriteKeyToFile = WritePrivateProfileStringByKeyName(App.EXEName, strKey, strValue, strFileSection)
End Function
' Function WriteKeyToFile
' -----------------------------

' @Function ReadKeyFromFile
' -----------------------------
'@Author Roland Ross Hadi
'@Description Reads key data form a file
'@Comments
' Parameter:
'       strFileSection - string used as a file subheader for every section
'       strKey - string used as key to read the value from file ( ex: UserName)
' Returns:   a string with the value tied to the key when written to file
'            a null string "" if Key is not on file
Public Function ReadKeyFromFile(strFileSection As String, strKey As String) As String
    Dim strValue As String
    Dim lngRetLen As Long
    
    strValue = String(READ_BUFF + 1, Space(1))
    lngRetLen = GetPrivateProfileStringByKeyName(App.EXEName, strKey, "", strValue, READ_BUFF, strFileSection)
    If lngRetLen > 1 Then
        ReadKeyFromFile = Left(strValue, lngRetLen)
    Else
        ReadKeyFromFile = ""
    End If
End Function
' Function ReadKeyFromFile
' -----------------------------

' @Function DeleteFromFile
' -----------------------------
'@Author Roland Ross Hadi
'@Description Deletes key data form a file
'@Comments
' Parameter:
'       strFileSection - string used as file subheader for every section
'       strKey - the string used in writing is also used for deleting
' Returns: -1 key or section is a null string
'           0 for system delete failure
'           1 for successfule delete
Public Function DeleteFromFile(strFileSection As String, strKey As String) As Long
    If Len(strFileSection) = 0 Or Len(strKey) = 0 Then
        MsgBox "Null string parameter not allowed for DeleteFromFile."
        DeleteFromFile = -1
        Exit Function
    End If
    DeleteFromFile = WritePrivateProfileStringToDeleteKey(App.EXEName, strKey, 0, strFileSection)
End Function
' Function DeleteFromFile
' -----------------------------

' @Function ReadAndWriteFile
' -----------------------------
'@Author Roland Ross Hadi
'@Description Reads and writes key data form a file
'@Comments
' Parameter:
'       strFileSection - string used as a file subheader for every section
'       strKey - string used as key to read the value from file ( ex: UserName)
'       strValue - Value to be entered
' Returns:   a string with the value tied to the key when written to file
'            a null string "" if Key is not on file
Public Function ReadAndWriteFile(strFileSection As String, strKey As String, strValue As String) As String
Dim tmp As String
tmp = ReadKeyFromFile(strFileSection, strKey)
If tmp = "" Then
  WriteKeyToFile strFileSection, strKey, strValue
  ReadAndWriteFile = strValue
Else
  ReadAndWriteFile = tmp
End If
End Function
' Function ReadAndWriteFile
' -----------------------------

' @Function AddBackslash
' -----------------------------
'@Author Roland Ross Hadi
'@Description Adds a backslash if the string doesn't have one already.
'@Comments
' Parameter:
'       strText - string to be processed
Public Function AddBackslash(strText As String) As String
   If Len(strText) > 0 Then
      If Right$(strText, 1) <> "\" Then
         AddBackslash = strText + "\"
      Else
         AddBackslash = strText
      End If
   Else
      AddBackslash = "\"
   End If
End Function
' Function AddBackslash
' -----------------------------

' @Function RemoveBackslash
' -----------------------------
'@Author Roland Ross Hadi
'@Description Removes a backslash if the string have one already.
'@Comments
' Parameter:
'       strText - string to be processed
Public Function RemoveBackslash(strText As String) As String
   If Len(strText) > 0 Then
      If Right$(strText, 1) = "\" Then
         RemoveBackslash = Left(strText, Len(strText) - 1)
      Else
         RemoveBackslash = strText
      End If
   Else
      RemoveBackslash = ""
   End If
End Function
' Function RemoveBackslash
' -----------------------------

' @Function FolderExists
' -----------------------------
'@Author Roland Ross Hadi
'@Description Check if folders exist
'@Comments
' Parameter:
'       strFullPath - string to be processed
Public Function FolderExists(strFullPath As String) As Boolean
    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FolderExists = True
EarlyExit:
    On Error GoTo 0
End Function
' Function FolderExists
' -----------------------------

' @Function GetFolders
' -----------------------------
'@Author Roland Ross Hadi
'@Description Check if folders exist
'@Comments
' Parameter:
'       strFullPath - string to be processed
Public Function GetFolders(strFullPath As String)
  Dim fso As New FileSystemObject
  Dim fld As folder
  Dim tmp()
  ReDim tmp(0)
  Set fld = fso.GetFolder(strFullPath)
  If fld.SubFolders.Count > 0 Then
    For Each fld In fld.SubFolders
      DoEvents
      tmp(UBound(tmp)) = fld.path
      ReDim Preserve tmp(UBound(tmp) + 1)
    Next
    ReDim Preserve tmp(UBound(tmp) - 1)
    GetFolders = tmp
  End If
End Function
' Function GetFolders
' -----------------------------

' @Function LoadXMLDocument
' -----------------------------
'@Author Roland Ross Hadi
'@Description Check if folders exist
'@Comments
' Parameter:
'       strFullPath - string to be processed
Public Sub LoadXMLDocument(xmlPath)
Dim xDoc As MSXML2.DOMDocument
Set xDoc = New MSXML2.DOMDocument
xDoc.validateOnParse = False
If xDoc.Load(xmlPath) Then
   LoadedXMLData = ""
   DisplayNode xDoc.childNodes, 0
Else
   MsgBox "XML data not loaded sucessfully"
End If
End Sub
Private Sub DisplayNode(ByRef Nodes As MSXML2.IXMLDOMNodeList, _
   ByVal Indent As Integer)
   Dim xNode As MSXML2.IXMLDOMNode
   Indent = Indent + 2
   For Each xNode In Nodes
      If xNode.NodeType = NODE_TEXT And (xNode.parentNode.nodeName = "ELAPSED_TIME" Or xNode.parentNode.nodeName = "COMPONENT_NAME" Or xNode.parentNode.nodeName = "STEP_DESCRIPTION" Or xNode.parentNode.nodeName = "IMAGE_PATH" Or xNode.parentNode.nodeName = "STEP_SUMMARY" Or xNode.parentNode.nodeName = "EXECUTION_TIME" Or xNode.parentNode.nodeName = "STEP_RESULT" Or xNode.parentNode.nodeName = "PC" Or xNode.parentNode.nodeName = "STATUS") Then
         LoadedXMLData = LoadedXMLData & xNode.parentNode.nodeName & _
            ":" & xNode.nodeValue & vbCrLf
      End If
      If xNode.hasChildNodes Then
         DisplayNode xNode.childNodes, Indent
      End If
   Next xNode
End Sub
' Function LoadXMLDocument
' -----------------------------

' @Function LoadXMLDocument_v2
' -----------------------------
'@Author Roland Ross Hadi
'@Description Check if folders exist
'@Comments
' Parameter:
'       strFullPath - string to be processed
Public Sub LoadXMLDocument_v2(xmlPath)
Dim xDoc As MSXML2.DOMDocument
Set xDoc = New MSXML2.DOMDocument
xDoc.validateOnParse = False
If xDoc.Load(xmlPath) Then
   LoadedXMLData = ""
   DisplayNode_v2 xDoc.childNodes, 0
Else
   MsgBox "XML data not loaded sucessfully"
End If
End Sub
Private Sub DisplayNode_v2(ByRef Nodes As MSXML2.IXMLDOMNodeList, _
   ByVal Indent As Integer)
   Dim xNode As MSXML2.IXMLDOMNode
   Indent = Indent + 2
   For Each xNode In Nodes
      With xNode.parentNode
        If xNode.NodeType = NODE_TEXT And (.nodeName = "ELAPSED_TIME" Or .nodeName = "COMPONENT_NAME" Or .nodeName = "STEP_DESCRIPTION" Or .nodeName = "IMAGE_PATH" Or .nodeName = "STEP_SUMMARY" Or .nodeName = "EXECUTION_TIME" Or .nodeName = "STEP_RESULT" Or .nodeName = "TAOVERSION" Or .nodeName = "TYPE" Or .nodeName = "QCUSER" Or .nodeName = "QCDOMAIN" Or .nodeName = "QCPROJECT" Or .nodeName = "TESTSETNAME" Or .nodeName = "TESTSETID" Or .nodeName = "TESTNAME" Or .nodeName = "TESTID" Or .nodeName = "TITLE" Or .nodeName = "TABLE_HEADER_TEXT" Or .nodeName = "TABLE_HEADER_TIME" Or .nodeName = "DEBUG_LOG" Or .nodeName = "LOG_FOLDER" Or .nodeName = "EXECUTED_BY" Or .nodeName = "PC" Or .nodeName = "STATUS") Then
           LoadedXMLData = LoadedXMLData & .nodeName & _
              ":" & xNode.nodeValue & vbCrLf
        End If
        If xNode.hasChildNodes Then
           DisplayNode_v2 xNode.childNodes, Indent
        End If
      End With
   Next xNode
End Sub
' Function LoadXMLDocument
' -----------------------------

