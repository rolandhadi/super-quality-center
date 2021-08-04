Attribute VB_Name = "mdlQC"
Option Explicit
Public QCConnection As New TDConnection
Public Const PROD = 5
Public Const TRAINING = 9
Public ENVIRON As Integer

Public frmControl_Option As Integer

Public Type RunTimeParams
    TestInstanceID As String
    ParamName() As String
    ParamValue() As String
End Type

Public Type BrowseInfo
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Public Type HPQC_FIELDS
    Table_Name As String
    Field_Name As String
    Technical_Name As String
    Filter_Value As String
    IsSpecial As Boolean
End Type

Public Type MyParam
    ParameterName As String
    ParameterValue As String
End Type

Public Type All_BC_Folders
    BC_ID As String
    BC_Name As String
    BC_Path As String
    BC_Parameters() As MyParam
End Type

Public Type All_BC_QC
    BC_ID As String
    BC_Name As String
End Type

Public Type QTP_Script
    Type As String
    ID As String
    LogicalName As String
    Class As String
    Name As String
    Action As String
    Value As String
    URI As String
End Type

Public Type URI
    Type As String
    ID As String
    LogicalName As String
    Class As String
    Name As String
    URI As String
End Type

Public Type TestPlan_Pars
    ID As String
    BC_ID As String
    Name As String
    Value As String
    Used As Boolean
End Type

Public GetAllBusinessComponents_MyComps() As All_BC_Folders


Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Const ADMIN_ID = "hpqc_qcvbn"
Public Const ADMIN_PASS = "hpqc_qcvbn_pass16"

Public ACCESS_ As String
'Public Const ACCESS_ = "Ultimate"
'Public Const ACCESS_ = "Team"
'Public Const ACCESS_ = "User"
'Public Const ACCESS_ = "Reporter"
 
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public curInstance As String, curDomain As String, curProject As String, curUser As String, curQCInstance As String

Public CheckedItems() As String
Public TPTARGET As String
Public BCTARGET As String
Public SPEF As String
Public REALTIME As Boolean
Public Control_Auto As Boolean
Public LastOthers As String

Public FXGirl As New clsAPIfunctions
Public FXSQCExtractCompleted As String
Public FXScriptConsolidationCompleted As String
Public FXExportToExcel As String
Public FXDataUploadCompleted As String
Public FX25 As String
Public FX50 As String
Public FX75 As String

Public GlobalStrings As New clsStrings

Public RepositoryFrom
Public ParentObject
Public All_URI() As URI
Public All_BC_QC() As All_BC_QC
Public CurVersion As String
 

Public Function CleanHTML(strText As String) As String
        Dim tmp As String
        tmp = Replace(strText, "<html><body>", "", 1, , vbTextCompare)
        tmp = Replace(tmp, "</body></html>", "", 1, , vbTextCompare)
        tmp = Replace(tmp, "&", "&amp;", 1, , vbTextCompare)
        tmp = Replace(tmp, "'", "''", 1, , vbTextCompare)
        tmp = Replace(tmp, "<", "&lt;", 1, , vbTextCompare)
        tmp = Replace(tmp, ">", "&gt;", 1, , vbTextCompare)
        tmp = Replace(tmp, """", "&quot;", 1, , vbTextCompare)
        tmp = Replace(tmp, vbCrLf, "<br>", 1, , vbTextCompare)
        tmp = ReplaceAllEnter(CStr(tmp))
        CleanHTML = tmp
End Function

Public Function CleanHTML_v1(strText As String) As String
        Dim tmp As String
        tmp = Replace(strText, "<html><body>", "", 1, , vbTextCompare)
        tmp = Replace(tmp, "</body></html>", "", 1, , vbTextCompare)
        tmp = Replace(tmp, "&", "&amp;", 1, , vbTextCompare)
        tmp = Replace(tmp, "'", "''", 1, , vbTextCompare)
        tmp = Replace(tmp, "<", "&lt;", 1, , vbTextCompare)
        tmp = Replace(tmp, ">", "&gt;", 1, , vbTextCompare)
        tmp = Replace(tmp, """", "&quot;", 1, , vbTextCompare)
        CleanHTML_v1 = tmp
End Function


Public Function ReverseCleanHTML(strText As String) As String
        Dim tmp As String
        tmp = Replace(strText, "<html><body>", "", 1, , vbTextCompare)
        tmp = Replace(tmp, "</body></html>", "", 1, , vbTextCompare)
        tmp = Replace(tmp, "&amp;", "&", 1, , vbTextCompare)
        tmp = Replace(tmp, "&apos;", "'", 1, , vbTextCompare)
        tmp = Replace(tmp, "&lt;", "<", 1, , vbTextCompare)
        tmp = Replace(tmp, "&gt;", ">", 1, , vbTextCompare)
        tmp = Replace(tmp, "&quot;", """", 1, , vbTextCompare)
        tmp = Replace(tmp, "<br>", vbCrLf, 1, , vbTextCompare)
        tmp = ReplaceAllEnter(tmp)
        ReverseCleanHTML = tmp
End Function

Public Function ReverseCleanHTMLnoBR(strText As String) As String
        Dim tmp As String
        tmp = Replace(strText, "<html><body>", "", 1, , vbTextCompare)
        tmp = Replace(tmp, "</body></html>", "", 1, , vbTextCompare)
        tmp = Replace(tmp, "&amp;", "&", 1, , vbTextCompare)
        tmp = Replace(tmp, "&apos;", "'", 1, , vbTextCompare)
        tmp = Replace(tmp, "&lt;", "<", 1, , vbTextCompare)
        tmp = Replace(tmp, "&gt;", ">", 1, , vbTextCompare)
        tmp = Replace(tmp, "&quot;", """", 1, , vbTextCompare)
        ReverseCleanHTMLnoBR = tmp
End Function

Public Function ReplaceAllEnter(Expr)
Dim i, tmp
tmp = Expr
For i = 1 To 1
    tmp = Replace(tmp, vbCrLf, "")
Next
For i = 1 To 1
    tmp = Replace(tmp, Chr(10) & Chr(13), "")
Next
For i = 1 To 1
    tmp = Replace(tmp, Chr(10), "")
Next
For i = 1 To 1
    tmp = Replace(tmp, Chr(13), "")
Next
For i = 1 To 1
    tmp = Replace(tmp, vbTab, "")
Next
For i = 1 To 1
    tmp = Replace(tmp, Chr(34), "")
Next
ReplaceAllEnter = tmp
End Function

Public Function GetFromTable(strID, strIDField, strPathField, strTableName)
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT " & strPathField & " FROM " & strTableName & " WHERE " & strIDField & " = " & strID
    Set rs = objCommand.Execute
    If rs.RecordCount <> 0 Then
        GetFromTable = rs.FieldValue(strPathField)
        Exit Function
    End If
End Function


' @Function FileWrite
' -----------------------------
'@Author Roland Ross Hadi
'@Description Write <strText> to <strFileName>
'@Comments
' Parameter:
'       strFileName - Filename to be created with the strText
'       strText - text to be wriiten to the file
Public Function FileWrite(strFileName As String, strText)
    Dim ff As Integer
    On Error Resume Next
    ff = FreeFile
    Kill strFileName
    Open strFileName For Binary As ff
        Put #ff, , CStr(strText)
    Close ff
End Function

' @Function FileAppend
' -----------------------------
'@Author Roland Ross Hadi
'@Description Append <strText> to <strFileName>
'@Comments
' Parameter:
'       strFileName - Filename to be appended with the strText
'       strText - text to be appended to the file
Public Function FileAppend(strFileName As String, strText)
    Dim ff As Integer
    On Error Resume Next
    ff = FreeFile
    Open strFileName For Append As ff
        Print #ff, CStr(strText)
    Close ff
End Function
' Function FileAppend
' -----------------------------

Public Function ColumnLetter(ColumnNumber As Integer) As String
  If ColumnNumber > 26 Then

    ' 1st character:  Subtract 1 to map the characters to 0-25,
    '                 but you don't have to remap back to 1-26
    '                 after the 'Int' operation since columns
    '                 1-26 have no prefix letter

    ' 2nd character:  Subtract 1 to map the characters to 0-25,
    '                 but then must remap back to 1-26 after
    '                 the 'Mod' operation by adding 1 back in
    '                 (included in the '65')

    ColumnLetter = Chr(Int((ColumnNumber - 1) / 26) + 64) & _
                   Chr(((ColumnNumber - 1) Mod 26) + 65)
  Else
    ' Columns A-Z
    ColumnLetter = Chr(ColumnNumber + 64)
  End If
End Function

'---------------------------------------------------------------------------------------
' Module    : modShellSort
' DateTime  : 27-12-2004
' Author    : Flyguy
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : ShellSortMultiColumn
' DateTime  : 27-12-2004
' Author    : Flyguy
' Purpose   : Sort a 2D string array on multiple columns
'---------------------------------------------------------------------------------------
Public Sub ShellSortMultiColumn(sArray() As String, cColumns As Collection, _
  cOrder As Collection)
  
  Dim lLoop1 As Long, lHValue As Long
  Dim lUBound As Long, lLBound As Long
  Dim lUBound2 As Long, lLBound2 As Long
  Dim lNofColumns As Long
  Dim aColumns() As Long, aOrder() As Long
  Dim bSorted As Boolean
  
  If cColumns Is Nothing Then Exit Sub
  If cOrder Is Nothing Then Exit Sub
  If cColumns.Count <> cOrder.Count Then Exit Sub
  
  lNofColumns = cColumns.Count
  ReDim aColumns(lNofColumns)
  ReDim aOrder(lNofColumns)
  For lLoop1 = 1 To lNofColumns
    aColumns(lLoop1) = cColumns(lLoop1)
    aOrder(lLoop1) = cOrder(lLoop1)
  Next lLoop1
  
  lUBound = UBound(sArray)
  lLBound = LBound(sArray)
  
  lUBound2 = UBound(sArray, 2)
  lLBound2 = LBound(sArray, 2)
  
  lHValue = (lUBound - lLBound) \ 2

  Do While lHValue > lLBound
    Do
      bSorted = True
      For lLoop1 = lLBound To lUBound - lHValue
        If CompareValues(sArray, lLoop1, lLoop1 + lHValue, lNofColumns, _
          aColumns, aOrder) Then
          SwapLines sArray, lLoop1, lLoop1 + lHValue, lLBound2, lUBound2
          bSorted = False
        End If
      Next lLoop1
      
    Loop Until bSorted
    lHValue = lHValue \ 2
  Loop

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SwapLines
' DateTime  : 27-12-2004
' Author    : Flyguy
' Purpose   : Swap a row of data in a 2D array
'---------------------------------------------------------------------------------------
Public Sub SwapLines(ByRef sArray() As String, lIndex1 As Long, _
  lIndex2 As Long, lLBound As Long, lUBound As Long)
  
  Dim i As Long, sTemp As String
  
  For i = lLBound To lUBound
    sTemp = sArray(lIndex1, i)
    sArray(lIndex1, i) = sArray(lIndex2, i)
    sArray(lIndex2, i) = sTemp
  Next i
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CompareValues
' DateTime  : 27-12-2004
' Author    : Flyguy
' Purpose   : Compare column values for multicolumn sorting
' Revision  : 15-02-2005, take in account numeric and date values
'---------------------------------------------------------------------------------------
Public Function CompareValues(ByRef sArray() As String, lIndex1 As Long, _
  lIndex2 As Long, lNofColumns As Long, aColumns() As Long, aOrder() As Long)
  
  Dim i As Long
  Dim LCol As Long
  Dim sValue1 As String, sValue2 As String
  Dim dValue1 As Double, dValue2 As Double
  Dim bNumeric As Boolean
  
  On Error Resume Next
  
  For i = 1 To lNofColumns
    LCol = aColumns(i)
    If aOrder(i) = 1 Then
      sValue1 = sArray(lIndex1, LCol)
      sValue2 = sArray(lIndex2, LCol)
    Else
      sValue1 = sArray(lIndex2, LCol)
      sValue2 = sArray(lIndex1, LCol)
    End If
        
    If IsDate(sValue1) And IsDate(sValue2) Then
      dValue1 = CDate(sValue1)
      dValue2 = CDate(sValue2)
      bNumeric = True
    ElseIf IsNumeric(sValue1) And IsNumeric(sValue2) Then
      dValue1 = CDbl(sValue1)
      dValue2 = CDbl(sValue2)
      bNumeric = True
    Else
      bNumeric = False
    End If
    
    If bNumeric Then
      If dValue1 < dValue2 Then
        Exit For
      ElseIf dValue1 > dValue2 Then
        CompareValues = True
        Exit For
      End If
    Else
      If sValue1 < sValue2 Then
        Exit For
      ElseIf sValue1 > sValue2 Then
        CompareValues = True
        Exit For
      End If
    End If
  Next i
  
End Function

Public Function CleanTheString(theString) As String
Dim strAlphaNumeric, i, CleanedString, strChar, j
      strAlphaNumeric = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-()@#$^&_+=,[].& " 'Used to check for numeric characters.
      For i = 1 To Len(theString)
          strChar = Mid(theString, i, 1)
          If InStr(strAlphaNumeric, strChar) Then
              CleanedString = CleanedString & strChar
          Else
              CleanedString = CleanedString & "_"
          End If
      Next
      For j = 1 To 10
        CleanedString = Replace(CleanedString, "  ", " ")
      Next
      CleanTheString = CleanedString
End Function

Public Function CleanTheString_PARAMS(theString) As String
Dim strAlphaNumeric, i, CleanedString, strChar
      strAlphaNumeric = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_" 'Used to check for numeric characters.
      For i = 1 To Len(theString)
          strChar = Mid(theString, i, 1)
          If InStr(strAlphaNumeric, strChar) Then
              CleanedString = CleanedString & strChar
          Else
              CleanedString = CleanedString & "_"
          End If
      Next
      CleanTheString_PARAMS = CleanedString
End Function

Public Function InvalidParameterCheck(theString) As Boolean
Dim strAlphaNumeric, i, CleanedString, strChar
      strAlphaNumeric = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_" 'Used to check for numeric characters.
      For i = 1 To Len(theString)
          strChar = Mid(theString, i, 1)
          If InStr(strAlphaNumeric, strChar) Then
              
          Else
              InvalidParameterCheck = True
              Exit Function
          End If
      Next
End Function

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
 
Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
 
    'declare variables to be used
     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim sPath As String
     Dim udtBI As BrowseInfo
 
    'initialise variables
     With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With
 
    'Call the browse for folder API
     lpIDList = SHBrowseForFolder(udtBI)
 
    'get the resulting string path
     If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
     End If
 
    'If cancel was pressed, sPath = ""
     BrowseForFolder = sPath
 
End Function

'########################### Extract Parameters from Step ###########################
Public Function ExtractParameters(StepDesc)
    Dim tmp, i, tmp2()
    ReDim tmp2(0)
    tmp = Split(StepDesc, "<<<", , vbTextCompare)
    For i = LBound(tmp) To UBound(tmp)
        If InStr(1, tmp(i), ">>>", vbTextCompare) <> 0 Then
            tmp2(UBound(tmp2)) = Left(tmp(i), InStr(1, tmp(i), ">>>", vbTextCompare) - 1)
            ReDim Preserve tmp2(UBound(tmp2) + 1)
        End If
    Next
    On Error GoTo Err1
    ReDim Preserve tmp2(UBound(tmp2) - 1)
    ExtractParameters = tmp2
Exit Function
Err1:
    ExtractParameters = -1
End Function
'########################### End Of Extract Parameters from Step ###########################

'########################### Extract Parameters from Step ###########################
Public Function ExtractParametersWithFix(StepDesc)
    Dim tmp, i, tmp2(), j
    ReDim tmp2(0)
    tmp = Split(StepDesc, "<<<", , vbTextCompare)
    For i = LBound(tmp) To UBound(tmp)
        If InStr(1, tmp(i), ">>>", vbTextCompare) <> 0 Then
            tmp2(UBound(tmp2)) = Trim(Left(tmp(i), InStr(1, tmp(i), ">>>", vbTextCompare) - 1))
            tmp2(UBound(tmp2)) = LCase(CleanTheString_PARAMS(tmp2(UBound(tmp2))))
            For j = 1 To 10
                tmp2(UBound(tmp2)) = Replace(tmp2(UBound(tmp2)), "__", "_")
            Next
            If UCase(Left(tmp2(UBound(tmp2)), 2)) = "P_" Then
                If UCase(Left(tmp2(UBound(tmp2)), 2)) <> "P_" Then
                    tmp2(UBound(tmp2)) = "p_" & tmp2(UBound(tmp2))
                End If
            ElseIf UCase(Left(tmp2(UBound(tmp2)), 2)) = "O_" Then
                If UCase(Left(tmp2(UBound(tmp2)), 2)) <> "O_" Then
                    tmp2(UBound(tmp2)) = "o_" & tmp2(UBound(tmp2))
                End If
            End If
            ReDim Preserve tmp2(UBound(tmp2) + 1)
        End If
    Next
    On Error GoTo Err1
    ReDim Preserve tmp2(UBound(tmp2) - 1)
    ExtractParametersWithFix = tmp2
Exit Function
Err1:
    ExtractParametersWithFix = -1
End Function
'########################### End Of Extract Parameters from Step ###########################


'########################### Checks Parameters from Step ###########################
Public Function HasParameters(StepDesc) As Boolean
    Dim tmp, i, tmp2()
    ReDim tmp2(0)
    tmp = Split(StepDesc, "<<<", , vbTextCompare)
    For i = LBound(tmp) To UBound(tmp)
        If InStr(1, tmp(i), ">>>", vbTextCompare) <> 0 Then
            tmp2(UBound(tmp2)) = Left(tmp(i), InStr(1, tmp(i), ">>>", vbTextCompare) - 1)
            ReDim Preserve tmp2(UBound(tmp2) + 1)
        End If
    Next
    On Error GoTo Err1
    ReDim Preserve tmp2(UBound(tmp2) - 1)
    HasParameters = True
Exit Function
Err1:
    HasParameters = False
End Function
'########################### End Of Checks Parameters from Step ###########################

Public Function IsParameterDeclared(AllParam, ParName) As Boolean
Dim k
On Error Resume Next
For k = LBound(AllParam) To UBound(AllParam)
    If Err.Number <> 0 Then Exit Function
    If UCase(AllParam(k)) = UCase(ParName) Then
        IsParameterDeclared = True
        Exit Function
    End If
Next
End Function

Public Function GetRequirementFolderPath(strID As String) As String
Dim Fact As ReqFactory
Dim Obj As Req
Dim FileFunct As New clsFiles
If Trim(strID) = "" Then Exit Function
Set Fact = QCConnection.ReqFactory
Set Obj = Fact.Item(strID)
GetRequirementFolderPath = FileFunct.GetFilePath(CStr(Obj.path))
End Function

Public Function GetBusinessComponentFolderPath(strID As String) As String
Dim Fact As ComponentFactory
Dim Obj As Component
If Trim(strID) = "" Then Exit Function
Set Fact = QCConnection.ComponentFactory
Set Obj = Fact.Item(strID)
GetBusinessComponentFolderPath = Obj.folder.path
End Function

Public Function GetBusinessComponentName(strID As String) As String
Dim Fact As ComponentFactory
Dim Obj As Component
If Trim(strID) = "" Then Exit Function
Set Fact = QCConnection.ComponentFactory
Set Obj = Fact.Item(strID)
GetBusinessComponentName = Obj.Name
End Function

Public Function GetTestFolderPath(strID As String) As String
Dim Fact As TreeManager
Dim Obj As SubjectNode
If Trim(strID) = "" Then Exit Function
Set Fact = QCConnection.TreeManager
Set Obj = Fact.NodeById(strID)
GetTestFolderPath = Obj.path
End Function

Public Function GetTestSetFolderPath(strID As String) As String
Dim Fact As TestSetFactory
Dim Obj As TestSet
If Trim(strID) = "" Then Exit Function
Set Fact = QCConnection.TestSetFactory
Set Obj = Fact.Item(strID)
GetTestSetFolderPath = Obj.TestSetFolder.path
End Function

Public Function GetTestInstanceFolderPath(strID As String) As String
Dim Fact As TSTestFactory
Dim Obj As tsTest
If Trim(strID) = "" Then Exit Function
Set Fact = QCConnection.TSTestFactory
Set Obj = Fact.Item(strID)
GetTestInstanceFolderPath = Obj.TestSet.TestSetFolder.path
End Function

Public Sub GetAllCheckedItems(objNode As Node)
    Dim objSiblingNode As Node
    Set objSiblingNode = objNode
Do
     If objSiblingNode.Checked = True Then
        CheckedItems(UBound(CheckedItems)) = objSiblingNode.Key
        ReDim Preserve CheckedItems(UBound(CheckedItems) + 1)
     End If
     If Not objSiblingNode.Child Is Nothing Then
         Call GetAllCheckedItems(objSiblingNode.Child)
     End If
     Set objSiblingNode = objSiblingNode.Next
Loop While Not objSiblingNode Is Nothing
End Sub

Public Sub SetHeader(Optional tmpForm As Form)
mdiMain.Caption = "RR-R1 Testing Team - SuperQC " & curInstance & " - " & curDomain & "-" & curProject & " ver." & App.Major & "." & App.Minor & "." & App.Revision
If IsNull(tmpForm) = True Then
    tmpForm.Caption = tmpForm.Tag
End If
End Sub

Public Sub GetAllBusinessComponents()
Dim tmp, i, FileFunct As New clsFiles, tmpStr, j, allPars
tmp = FileFunct.GetFolders(App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject)
ReDim GetAllBusinessComponents_MyComps(UBound(tmp))
FileFunct.FileDelete App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & curDomain & "-" & curProject & ".hxh"
For i = LBound(tmp) To UBound(tmp)
    tmpStr = FileFunct.ReadFromFile(tmp(i) & "\Params.txt")
    GetAllBusinessComponents_MyComps(i).BC_ID = GetContent("{", "}", tmpStr)
    GetAllBusinessComponents_MyComps(i).BC_Name = GetContent("|", "|", tmpStr)
    GetAllBusinessComponents_MyComps(i).BC_Path = GetContent("<", ">", tmpStr)
    GetAllBusinessComponents_GetParameters CStr(tmp(i) & "\Params.txt"), i
    allPars = ""
    For j = LBound(GetAllBusinessComponents_MyComps(i).BC_Parameters) To UBound(GetAllBusinessComponents_MyComps(i).BC_Parameters)
        allPars = allPars & GetAllBusinessComponents_MyComps(i).BC_Parameters(j).ParameterName & "�" & GetAllBusinessComponents_MyComps(i).BC_Parameters(j).ParameterValue & "�"
    Next
    allPars = Left(allPars, Len(allPars) - 1)
    FileFunct.WriteToEndOfFile App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & curDomain & "-" & curProject & ".hxh", GetAllBusinessComponents_MyComps(i).BC_ID & "�" & GetAllBusinessComponents_MyComps(i).BC_Name & "�" & GetAllBusinessComponents_MyComps(i).BC_Path & "�" & allPars
    Debug.Print i
Next
End Sub
' HERE!!!!
Public Sub GetAllBusinessComponent(BC_ID As String)
Dim tmp, i, FileFunct As New clsFiles, tmpStr, j, k, allPars
        tmpStr = FileFunct.ReadFromFile(App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & BC_ID & "\Params.txt")
        tmp = GetContent("{", "}", tmpStr)
        For k = LBound(GetAllBusinessComponents_MyComps) To UBound(GetAllBusinessComponents_MyComps)
            If GetAllBusinessComponents_MyComps(k).BC_ID = tmp Then
                Exit Sub
            End If
        Next
        ReDim Preserve GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps) + 1)
        i = UBound(GetAllBusinessComponents_MyComps)
        GetAllBusinessComponents_MyComps(i).BC_ID = GetContent("{", "}", tmpStr)
        GetAllBusinessComponents_MyComps(i).BC_Name = GetContent("|", "|", tmpStr)
        GetAllBusinessComponents_MyComps(i).BC_Path = GetContent("<", ">", tmpStr)
        GetAllBusinessComponents_GetParameters CStr(App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & BC_ID & "\Params.txt"), i
        allPars = ""
        For j = LBound(GetAllBusinessComponents_MyComps(i).BC_Parameters) To UBound(GetAllBusinessComponents_MyComps(i).BC_Parameters)
            allPars = allPars & GetAllBusinessComponents_MyComps(i).BC_Parameters(j).ParameterName & "�" & GetAllBusinessComponents_MyComps(i).BC_Parameters(j).ParameterValue & "�"
        Next
        allPars = Left(allPars, Len(allPars) - 1)
        FileFunct.WriteToEndOfFile App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & curDomain & "-" & curProject & ".hxh", GetAllBusinessComponents_MyComps(i).BC_ID & "�" & GetAllBusinessComponents_MyComps(i).BC_Name & "�" & GetAllBusinessComponents_MyComps(i).BC_Path & "�" & allPars
End Sub
'HERE!!!

Public Sub LoadAllComponentsFromQC()
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Set objCommand = QCConnection.Command
Dim SQL
SQL = "SELECT CO_ID, CO_NAME FROM COMPONENT, COMPONENT_FOLDER WHERE FC_ID = CO_FOLDER_ID AND (FC_PATH LIKE '" & SPEF & "' AND CO_NAME NOT LIKE '%" & "ZZZ" & "%')"
objCommand.CommandText = SQL
Set rs = objCommand.Execute
If rs.RecordCount > 0 Then
    ReDim All_BC_QC(0)
    Do While rs.EOR = False
        All_BC_QC(UBound(All_BC_QC)).BC_ID = rs.FieldValue("CO_ID")
        All_BC_QC(UBound(All_BC_QC)).BC_Name = rs.FieldValue("CO_NAME")
        ReDim Preserve All_BC_QC(UBound(All_BC_QC) + 1)
        rs.Next
    Loop
End If
End Sub

Public Sub LoadAllBusinessComponent(BC_ID, BC_Name)
Dim tmp, i, FileFunct As New clsFiles, BCDetails, Params, j
DumpBusinessComponent CStr(BC_ID), CStr(BC_Name)
ReDim GetAllBusinessComponents_MyComps(0)
tmp = FileFunct.ReadFromFileToArray(App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & curDomain & "-" & curProject & ".hxh")
On Error Resume Next
If tmp(LBound(tmp)) <> "" Then
ElseIf tmp = "No Data Found" Then
    Exit Sub
End If
If Err.Number = 13 Then Exit Sub
On Error GoTo 0
ReDim Preserve tmp(UBound(tmp) - 1)
For i = LBound(tmp) To UBound(tmp)
   If InStr(1, tmp(i), BC_ID & "�", vbTextCompare) <> 0 Then
        ReDim Preserve GetAllBusinessComponents_MyComps(i)
        BCDetails = Split(tmp(i), "�")
        GetAllBusinessComponents_MyComps(i).BC_ID = BCDetails(0)
        GetAllBusinessComponents_MyComps(i).BC_Name = BCDetails(1)
        GetAllBusinessComponents_MyComps(i).BC_Path = BCDetails(2)
        Params = Split(BCDetails(3), "�")
        For j = LBound(Params) To UBound(Params)
            ReDim Preserve GetAllBusinessComponents_MyComps(i).BC_Parameters(j)
            GetAllBusinessComponents_MyComps(i).BC_Parameters(j).ParameterName = Left(Params(j), InStr(1, Params(j), "�") - 1)
            GetAllBusinessComponents_MyComps(i).BC_Parameters(j).ParameterValue = Replace(Params(j), Left(Params(j), InStr(1, Params(j), "�")), "")
        Next
        Exit Sub
   End If
Next
End Sub

Public Sub LoadAllBusinessComponents()
Dim tmp, i, FileFunct As New clsFiles, BCDetails, Params, j
ReDim GetAllBusinessComponents_MyComps(0)
tmp = FileFunct.ReadFromFileToArray(App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & curDomain & "-" & curProject & ".hxh")
On Error Resume Next
If tmp(LBound(tmp)) <> "" Then
ElseIf tmp = "No Data Found" Then
    Exit Sub
End If
If Err.Number = 13 Then Exit Sub
On Error GoTo 0
ReDim Preserve tmp(UBound(tmp) - 1)
For i = LBound(tmp) To UBound(tmp)
   ReDim Preserve GetAllBusinessComponents_MyComps(i)
   BCDetails = Split(tmp(i), "�")
   GetAllBusinessComponents_MyComps(i).BC_ID = BCDetails(0)
   GetAllBusinessComponents_MyComps(i).BC_Name = BCDetails(1)
   GetAllBusinessComponents_MyComps(i).BC_Path = BCDetails(2)
   Params = Split(BCDetails(3), "�")
   For j = LBound(Params) To UBound(Params)
       ReDim Preserve GetAllBusinessComponents_MyComps(i).BC_Parameters(j)
       GetAllBusinessComponents_MyComps(i).BC_Parameters(j).ParameterName = Left(Params(j), InStr(1, Params(j), "�") - 1)
       GetAllBusinessComponents_MyComps(i).BC_Parameters(j).ParameterValue = Replace(Params(j), Left(Params(j), InStr(1, Params(j), "�")), "")
   Next
Next
End Sub

Public Function GetContent(StartD, EndD, txtString)
Dim tmp, startI, endI
startI = InStr(1, txtString, StartD)
endI = InStr(startI + 1, txtString, EndD)
GetContent = Trim(Mid(txtString, startI + 1, endI - startI - 1))
End Function

Public Sub GetAllBusinessComponents_GetParameters(txtString, Index)
Dim FileFunct As New clsFiles
Dim tmp, tmpComp(), i
tmp = FileFunct.ReadFromFileToArray(CStr(txtString))
ReDim GetAllBusinessComponents_MyComps(Index).BC_Parameters(0)
For i = LBound(tmp) To UBound(tmp)
    If InStr(1, tmp(i), "[", vbTextCompare) <> 0 And InStr(1, tmp(i), "]", vbTextCompare) <> 0 Then
        GetAllBusinessComponents_MyComps(Index).BC_Parameters(UBound(GetAllBusinessComponents_MyComps(Index).BC_Parameters)).ParameterName = GetContent("[", "]", tmp(i))
        GetAllBusinessComponents_MyComps(Index).BC_Parameters(UBound(GetAllBusinessComponents_MyComps(Index).BC_Parameters)).ParameterValue = Trim(Replace(tmp(i), "[" & GetAllBusinessComponents_MyComps(Index).BC_Parameters(UBound(GetAllBusinessComponents_MyComps(Index).BC_Parameters)).ParameterName & "]", ""))
        ReDim Preserve GetAllBusinessComponents_MyComps(Index).BC_Parameters(UBound(GetAllBusinessComponents_MyComps(Index).BC_Parameters) + 1)
    End If
Next
If UBound(GetAllBusinessComponents_MyComps(Index).BC_Parameters) > 0 Then ReDim Preserve GetAllBusinessComponents_MyComps(Index).BC_Parameters(UBound(GetAllBusinessComponents_MyComps(Index).BC_Parameters) - 1)
End Sub

'***************************************************************************
Public Function ExtractURI(ParentObject)
Dim uri_tmp
Dim tmp, fTOCollection, RepObjIndex, fTestObject, PropertiesColl, pIndex, ObjectProperty
'Get Objects by parent From loaded Repository
'If parent not specified all objects will be returned
On Error Resume Next
Set fTOCollection = RepositoryFrom.GetChildren(ParentObject)
    For RepObjIndex = 0 To fTOCollection.Count - 1
        uri_tmp = "label=LABELXXX01; type=TYPEXXX02; id=IDXXX03"
        tmp = uri_tmp
        'Get object by index
        Set fTestObject = fTOCollection.Item(RepObjIndex)
            'Check whether the object is having child objects
            If RepositoryFrom.GetChildren(fTestObject).Count <> 0 Then
                All_URI(UBound(All_URI)).LogicalName = RepositoryFrom.GetLogicalName(fTestObject)
                Debug.Print "Object Logical Name:= " & RepositoryFrom.GetLogicalName(fTestObject)
                Debug.Print "**********************************************************"
                    'Get TO Properties List
                    Set PropertiesColl = fTestObject.GetTOProperties
                    For pIndex = 0 To PropertiesColl.Count - 1
                                'debug.print Property and Value
                                Set ObjectProperty = PropertiesColl.Item(pIndex)
                                Debug.Print ObjectProperty.Name&; ":=" & ObjectProperty.Value
                                Select Case UCase(Trim(ObjectProperty.Name))
                                Case "NAME"
                                    tmp = Replace(tmp, "LABELXXX01", ObjectProperty.Value)
                                    All_URI(UBound(All_URI)).Name = ObjectProperty.Value
                                Case "TYPE"
                                    tmp = Replace(tmp, "TYPEXXX02", ObjectProperty.Value)
                                    All_URI(UBound(All_URI)).Type = ObjectProperty.Value
                                Case "ID"
                                    tmp = Replace(tmp, "IDXXX03", ObjectProperty.Value)
                                    All_URI(UBound(All_URI)).ID = ObjectProperty.Value
                                End Select
                    Next
                All_URI(UBound(All_URI)).URI = tmp
                ReDim Preserve All_URI(UBound(All_URI) + 1)
                Debug.Print tmp
                Debug.Print "**********************************************************"
'Calling Recursive Function
                ExtractURI fTestObject
            Else
                
                Debug.Print "**********************************************************"
                   Debug.Print "Object Logical Name:= " & RepositoryFrom.GetLogicalName(fTestObject)
                   All_URI(UBound(All_URI)).LogicalName = RepositoryFrom.GetLogicalName(fTestObject)
                Debug.Print "**********************************************************"
                    Set PropertiesColl = fTestObject.GetTOProperties
                    For pIndex = 0 To PropertiesColl.Count - 1
                                Set ObjectProperty = PropertiesColl.Item(pIndex)
                                Debug.Print ObjectProperty.Name&; ":=" & ObjectProperty.Value
                                Select Case UCase(Trim(ObjectProperty.Name))
                                Case "NAME"
                                    tmp = Replace(tmp, "LABELXXX01", ObjectProperty.Value)
                                    All_URI(UBound(All_URI)).Name = ObjectProperty.Value
                                Case "TYPE"
                                    tmp = Replace(tmp, "TYPEXXX02", ObjectProperty.Value)
                                    All_URI(UBound(All_URI)).Type = ObjectProperty.Value
                                Case "ID"
                                    tmp = Replace(tmp, "IDXXX03", ObjectProperty.Value)
                                    All_URI(UBound(All_URI)).ID = ObjectProperty.Value
                                End Select
                    Next
                    All_URI(UBound(All_URI)).URI = tmp
                    ReDim Preserve All_URI(UBound(All_URI) + 1)
                    Debug.Print tmp
            End If
    Next
End Function
'***************************************************************************

Public Sub GetURIs(ObjectRepositoryPath As String)
ReDim All_URI(0)
'ObjectRepositoryPath = "C:\Users\M14x\Documents\Mass Uploads\AUTOMATION\03-06-12\O3O_RT01\Action1\ObjectRepository.bdb"
'Creating Object Repository  utility Object
Set RepositoryFrom = Nothing
Set RepositoryFrom = CreateObject("Mercury.ObjectRepositoryUtil")
'Load Object Repository
RepositoryFrom.Load ObjectRepositoryPath
Call ExtractURI(ParentObject)
End Sub

Public Sub DumpBusinessComponent(Dump_BC_ID As String, Dump_BC_Name As String)
On Error Resume Next
Dim i As Integer, i2 As Integer
Dim k As Integer
Dim z As Integer, tmp, FileFunct As New clsFiles

Dim comp As Component
Dim compStorage As ExtendedStorage
Dim CompDownLoadPath As String
Dim NullList As List
Dim isFatalErr As Boolean
Dim compFact As ComponentFactory

Dim compParamFactory As ComponentParamFactory
Dim compParam As ComponentParam
Dim tmpList As List, FileStruct As New clsFiles

Dim IssueCount As Integer

tmp = ""

Set compFact = QCConnection.ComponentFactory
Set comp = compFact.Item(Dump_BC_ID)
Set compParamFactory = comp.ComponentParamFactory
Set tmpList = compParamFactory.NewList("")
tmp = ""
tmp = "{" & Dump_BC_ID & "}" & vbCrLf
'tmp = tmp & "~" & rs2.FieldValue(2) & "~" & vbCrLf
tmp = tmp & "<" & GetBusinessComponentFolderPath(Dump_BC_ID) & ">" & vbCrLf
tmp = tmp & "|" & Dump_BC_Name & "|" & vbCrLf
For i2 = 1 To tmpList.Count
    tmp = tmp & "[" & tmpList.Item(i2).Name & "] " & tmpList.Item(i2).Value & vbCrLf
Next
Set compStorage = comp.ExtendedStorage(0)
'compStorage.ClientPath = App.path & "\SQC Logs\bin\" & Format(Dump_BC_ID, "0000000000") & "-" & Dump_BC_name & "\" & Dump_BC_ID
compStorage.ClientPath = App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & Dump_BC_ID
CompDownLoadPath = compStorage.Load("Action1\Script.mts,Action1\Resource.mtr", True)
FileStruct.WriteNewFile App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & Dump_BC_ID & "\Params.txt", CStr(tmp)
'CompDownLoadPath = compStorage.SaveEx("\Action1\Script.mts, Action1\Resource.mtr", True, NullList) 'SAVE TO QC

If Err.Number = 0 Then
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Dumping BC Script (PASSED) " & Now & " " & Format(Dump_BC_ID, "0000000") & "-" & Dump_BC_Name
Else
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Dumping BC Script (FAILED) " & Now & " " & Format(Dump_BC_ID, "0000000") & "-" & Dump_BC_Name & " (" & Err.Description & ")"
    IssueCount = IssueCount + 1
End If
Err.Clear
GetAllBusinessComponent Dump_BC_ID 'HERE!!!
End Sub

Public Function ResolveParameter(ParName As String)
Dim tmp, i
    tmp = CleanTheString_PARAMS(Trim(ParName))
    Do While InStr(1, tmp, "__", vbTextCompare) <> 0
        tmp = Replace(tmp, "__", "_")
    Loop
    If Right(tmp, 1) = "_" Then tmp = Left(tmp, Len(tmp) - 1)
    ResolveParameter = Trim(tmp)
End Function

Public Function LatestVersionCheck() As Boolean
Dim UpdatedV As String
Dim curV As String
Dim AUTO_UPDATE_SQC_ID  As String
AUTO_UPDATE_SQC_ID = GetFromTable("'_AUTO_UPDATE_SQC_'", "TS_NAME", "TS_TEST_ID", "TEST")
If AUTO_UPDATE_SQC_ID = "" Then Exit Function
UpdatedV = CleanTheString(CleanHTML(GetFromTable(AUTO_UPDATE_SQC_ID, "TS_TEST_ID", "TS_DESCRIPTION", "TEST")))
curV = CleanTheString(CurVersion)
If Trim(UCase(UpdatedV)) = Trim(UCase(curV)) Then
  LatestVersionCheck = True
Else
  LatestVersionCheck = False
End If
End Function

Public Sub Patch()
Dim UpdateF As String, CurF As String, FileFunct As New clsFiles
UpdateF = DownloadAttachment(GetTest(GetFromTable("'_AUTO_UPDATE_SQC_'", "TS_NAME", "TS_TEST_ID", "TEST")))
CurF = App.path & "\" & App.EXEName & ".exe"
FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myPatch.hxh", "<PATCHPATH_UPDATE>", UpdateF
FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myPatch.hxh", "<PATCHPATH_CUR>", CurF
MsgBox "A new version of SuperQC was detected in the server. This superQC version will be updated automatically." & vbCrLf & "The tool will now restart", vbCritical
QCConnection.SendMail "user@companyemail.com", "", "[HPQC UPDATES] superQC version was updated by " & curUser, "superQC version was updated by " & curUser, "HTML"
MsgBox "The superQC latest patch(es) were installed successfully" & vbCrLf & "Press OK to launch superQC", vbInformation
Shell App.path & "\SQC DAT" & "\" & App.EXEName & ".exe"
End
End Sub

Public Function DownloadAttachment(theTest As TDAPIOLELib.Test) As String
    Dim TestAttachFact As AttachmentFactory
    Dim AttachList As List, TAttach As Attachment
    Dim TestAttachStorage As IExtendedStorage
    Dim TestStorage As IExtendedStorage
    Dim TestDownLoadPath As String, AttachDownLoadPath$
    Dim isFatalErr As Boolean
    Dim OwnerType As String, OwnerKey As Variant
    Dim filename As String
' Get the Attachments.
    Set TestAttachFact = theTest.Attachments
    TestAttachFact.FactoryProperties OwnerType, OwnerKey
    Debug.Print "OwnerType = " & OwnerType & ", " _
        & "OwnerKey = " & OwnerKey; ""
' OwnerType = TEST, OwnerKey = 98
'
'----------------------------
' Get the list of attachments and go through
' the list, downloading one at a time.
    Set AttachList = TestAttachFact.NewList("")
    For Each TAttach In AttachList
      With TAttach
        Debug.Print "----------------------------"
        Debug.Print "Download attachment" & vbCrLf
        Debug.Print "Before setting path"
        Debug.Print "The attachment name: " & .Name
        Debug.Print "The attachment server name: " & .ServerFileName
        Debug.Print "The filename: " & .filename
'----------------------------------------------------
' Use Attachment.AttachmentStorage to get the
' extended storage object.
        Set TestAttachStorage = .AttachmentStorage
        TestAttachStorage.ClientPath = _
            App.path & "\SQC DAT" & "\" & theTest.Name & "\attachStorage"
'----------------------------------------------------
' Use Attachment.Load to download the attachment files.
        TAttach.Load True, AttachDownLoadPath
' Note that the Attachment.FileName property changes as a result
' of setting the IExtendedStorage.ClientPath.
        Debug.Print vbCrLf & "After download"
        Debug.Print "Down load path: " & AttachDownLoadPath
        filename = .filename
        Debug.Print "The filename: " & .filename
' After download
      End With
      Exit For
    Next TAttach
    DownloadAttachment = filename
End Function

Public Function DownloadAttachments(theTest As TDAPIOLELib.Test) As String
    Dim TestAttachFact As AttachmentFactory
    Dim AttachList As List, TAttach As Attachment
    Dim TestAttachStorage As IExtendedStorage
    Dim TestStorage As IExtendedStorage
    Dim TestDownLoadPath As String, AttachDownLoadPath$
    Dim isFatalErr As Boolean
    Dim OwnerType As String, OwnerKey As Variant
    Dim filename As String
' Get the Attachments.
    Set TestAttachFact = theTest.Attachments
    TestAttachFact.FactoryProperties OwnerType, OwnerKey
    Debug.Print "OwnerType = " & OwnerType & ", " _
        & "OwnerKey = " & OwnerKey; ""
' OwnerType = TEST, OwnerKey = 98
'
'----------------------------
' Get the list of attachments and go through
' the list, downloading one at a time.
    Set AttachList = TestAttachFact.NewList("")
    For Each TAttach In AttachList
      With TAttach
        Debug.Print "----------------------------"
        Debug.Print "Download attachment" & vbCrLf
        Debug.Print "Before setting path"
        Debug.Print "The attachment name: " & .Name
        Debug.Print "The attachment server name: " & .ServerFileName
        Debug.Print "The filename: " & .filename
'----------------------------------------------------
' Use Attachment.AttachmentStorage to get the
' extended storage object.
        Set TestAttachStorage = .AttachmentStorage
        TestAttachStorage.ClientPath = _
            App.path & "\SQC DAT" & "\" & theTest.Name & "\attachStorage"
'----------------------------------------------------
' Use Attachment.Load to download the attachment files.
        If InStr(1, .Name, ".hxh", vbTextCompare) <> 0 Then
          TAttach.Load True, AttachDownLoadPath
        End If
' Note that the Attachment.FileName property changes as a result
' of setting the IExtendedStorage.ClientPath.
        Debug.Print vbCrLf & "After download"
        Debug.Print "Down load path: " & AttachDownLoadPath
        filename = .filename
        Debug.Print "The filename: " & .filename
' After download
      End With
    Next TAttach
    DownloadAttachments = filename
    DeleteAttachTestPlan
    RenameTheFiles ".hxh", ".jpg", App.path & "\SQC DAT" & "\" & theTest.Name & "\attachStorage"
End Function

Public Sub DeleteAttachTestPlan()
Dim Treemgr As TreeManager
Dim subjnode As SubjectNode
Dim tfact As TestFactory
Dim Test As Test
Dim tList As List
Dim tfilter As TDFilter
Dim tsname
Dim i
Dim fso As FileSystemObject
Dim Attachfile
Dim AttachFact As AttachmentFactory
Dim Attachment As Attachment
Dim AttachList As List

For i = 1 To 1
    Set Treemgr = QCConnection.TreeManager
    Set subjnode = Treemgr.NodeByPath("Subject\BPT Resources\SuperQC\")
    Set tfact = subjnode.TestFactory
    Set tfilter = tfact.Filter
    Attachfile = ".hxh"
    tsname = "_AUTO_UPDATE_SQC_"
    tsname = Replace(tsname, " ", "*")
    tsname = Replace(tsname, "(", "*")
    tsname = Replace(tsname, ")", "*")
    tfilter.Filter("TS_NAME") = tsname
    Set tList = tfact.NewList(tfilter.Text)
    Set Test = tList.Item(1)
    Set AttachFact = Test.Attachments
    Set AttachList = AttachFact.NewList("")
    For Each Attachment In AttachList
        If InStr(1, Attachment.Name, Attachfile) <> 0 Then
            AttachFact.RemoveItem Attachment
        End If
    Next
Next
End Sub

Public Sub RenameTheFiles(OldTxt As String, NewString As String, curPath As String)
Dim i As Integer
Dim from_str As String
Dim to_str As String
Dim dir_path As String
Dim old_name As String
Dim new_name As String

    On Error GoTo RenameError

    from_str = OldTxt
    to_str = NewString
    dir_path = curPath
    
    If Right$(dir_path, 1) <> "\" Then dir_path = dir_path _
        & "\"

    old_name = Dir$(dir_path & "*.*", vbNormal)
    Do While Len(old_name) > 0
        ' Rename this file.
        new_name = Replace$(old_name, from_str, to_str)
        If new_name <> old_name Then
            Name dir_path & old_name As dir_path & Right(new_name, 17)
            i = i + 1
        End If

        ' Get the next file.
        old_name = Dir$()
    Loop
    Exit Sub
RenameError:
    MsgBox Err.Description
End Sub

Public Function GetTest(TestID As String)
Dim tfact As TestFactory
Dim mytest As Test
Set tfact = QCConnection.TestFactory
Set mytest = tfact.Item(TestID)
Set GetTest = mytest
End Function
