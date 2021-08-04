VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************************************************************************************************
'*      LIBRARY NAME: clsDatabase.cls
'*      PURPOSE: Compilation of functions that deals with ADODB
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
        'Force the definition of the variables
        Option Explicit
        'This value will be return if the password isn't found
        Private Const sNoPassword As String = "-No password found-"
        'Use by GetPreviousNextID function
        Public Enum eDirection
            ePrevious = 0
            eNext = 1
        End Enum
        'Use to know which kind of connection we need
        Public Enum DBFormat
            eJET = 0
            eODBC = 1 'SQL
            eFoxPro = 2
        End Enum
        'Use in the CreateDatabase function
        Public Enum AccessVersion
            Access2000 = 0
            Access97 = 1
            AccessXP = 2
        End Enum
        'Store regesty information
        Public dbPath As String
        Public dbFileName As String
        'Create an instance of the object connection
        Public cnn As ADODB.Connection
'===============================================================================


' @Function SetAppInfo
' -----------------------------
'@Author Roland Ross Hadi
'@Description Set the registry key for the current application
'@Comments
' Parameter:
'       DbName - Database name
Public Sub SetAppInfo(DbName As String)
    If GetSetting(App.Title, "Control", "Title") = "" Then
        Call SaveSetting(App.Title, "Control", "Title", App.Title)
        Call SaveSetting(App.Title, "Control", "Version", App.Major & "." & App.Minor & "." & App.Revision)
        Call SaveSetting(App.Title, "Install", "Date", Now())
        Call SaveSetting(App.Title, "Install", "AppPath", App.path)
        Call SaveSetting(App.Title, "NCS", "Copyright", "(c)" & Year(date))
        Call SaveSetting(App.Title, "Data", "DBName", DbName)
        Call SaveSetting(App.Title, "Data", "DBPath", App.path & "\Data")
    End If
End Sub
' Function SetAppInfo
' -----------------------------

' @Function GetDBInfo
' -----------------------------
'@Author Roland Ross Hadi
'@Description Set the default path and name of the database
'@Comments
' Parameter:
'
Public Sub GetDBInfo()
    dbPath = GetSetting(App.Title, "Data", "DBPath")
    dbFileName = GetSetting(App.Title, "Data", "DBName")
End Sub
' Function GetDBInfo
' -----------------------------

' @Function OpenDBADO
' -----------------------------
'@Author Roland Ross Hadi
'@Description Create a new connection and open the database through the Jet engine
'@Comments
' Parameter:
'
Public Sub OpenDBADO(Optional DbName As String, Optional DatabaseEngine As DBFormat = 0, Optional sSourcePath As String, Optional isExclusive As Boolean)
On Error GoTo ErrorHandler
    Dim sCnnString As String
    Dim sExclusive As String
   
    'Create a new instance of a ADO Connection
    Set cnn = New ADODB.Connection
   
    'If the open is call implicitly
    If DbName = "" Then
        Call GetDBInfo
        DbName = dbPath & "\" & dbFileName
    End If
   
    If isExclusive Then
        sExclusive = "Oui"
    Else
        sExclusive = "Non"
    End If
   
    Select Case DatabaseEngine
        Case DBFormat.eJET
            'Connection with JET 4.0 (VB with service pack 5)
            sCnnString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source="
        Case DBFormat.eFoxPro
            sCnnString = "UID=;" & _
                        "PWD=;" & _
                        "SourceDB=;" & _
                        "SourceType=DBF;" & _
                        "Exclusive=No;" & _
                        "BackgroundFetch=Yes;" & _
                        "Collate=Machine;" & _
                        "Null=No;" & _
                        "Deleted=No;" & _
                        "DSN="
        Case DBFormat.eODBC
            sCnnString = "Provider=MSDASQL.1;" & _
                        "Persist Security Info=False;" & _
                        "Data Source="
    End Select
    cnn.Open sCnnString & DbName
Exit_Sub:
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
        MsgBox "Connection problem with SQL Server. The server may be down, " & _
            "check with the network administrator", vbCritical
    Else
        MsgBox Err.Number & " - " & Err.Description
    End If
    Err.Clear
    Resume Exit_Sub
End Sub
' @Function OpenDBADO
' -----------------------------

' @Function CloseDBADO
' -----------------------------
'@Author Roland Ross Hadi
'@Description Close the connection and Destroy the cnn object
'@Comments
' Parameter:
'
Public Sub CloseDBADO()
On Error GoTo ErrorHandler
    'Close the connection
    If cnn.State = 1 Then cnn.Close
   
Exit_Sub:
    'Destroy the cnn object
'    Set cnn = Nothing
    Exit Sub
   
ErrorHandler:
    'If the connection object is already closed
    If Err.Number = 91 Then
        Err.Clear
        Resume Exit_Sub
    End If
    MsgBox Err.Number & "-" & Err.Description
    Err.Clear
    Resume Exit_Sub
End Sub
' @Function CloseDBADO
' -----------------------------

' @Function HowManyData
' -----------------------------
'@Author Roland Ross Hadi
'@Description Use to return the amount of data found, based on the SQL criteria
'@Comments
' Parameter:
'        tblName - Table name
'        sWhere - Condition (Ex:City_Name="Montreal")
Public Function HowManyData(tblName As String, Optional sWhere As String) As Long
    On Error GoTo ErrorHandler
    Dim sSQL As String
    Dim rs As ADODB.Recordset
   
    'Create a recorset object
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenStatic
   
    'If there is a condtion, append it to the SQL statement
    If Len(sWhere) > 0 Then
        If InStr(1, UCase(sWhere), "WHERE") <> 0 Then
            sSQL = "SELECT * FROM " & tblName & " " & sWhere
        Else
            sSQL = "SELECT * FROM " & tblName & " WHERE " & sWhere
        End If
    Else
        sSQL = "SELECT * FROM " & tblName
    End If
   
    'Open the recordset based on the sSQL
'    Debug.Print sSQL
    rs.Open sSQL, cnn
   
    'Assign the number of recordset found
    HowManyData = rs.RecordCount
   
    'Close the recordset
    rs.Close
   
Exit_Function:
    'Destroy the recorset object
    Set rs = Nothing
    Exit Function
   
ErrorHandler:
    Call HandleError(Err.Number)
    Err.Clear
    Resume Exit_Function
End Function
' Function HowManyData
' -----------------------------

' @Function HowManyDataSQL
' -----------------------------
'@Author Roland Ross Hadi
'@Description Use to return the amount of data found, based on the SQL criteria
'@Comments
' Parameter:
'        sSQLtoCount - SQL statement
Public Function HowManyDataSQL(sSQLtoCount As String) As Long
    On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
   
    'Create a recorset object
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenStatic
    'Debug.Print sSQLtoCount
    'Open the recordset based on the sSQLtoCount
    rs.Open sSQLtoCount, cnn
   
    'Assign the number of recordset found
    HowManyDataSQL = rs.RecordCount
   
    'Close the recordset
    rs.Close
   
Exit_Function:
    'Destroy the recorset object
    Set rs = Nothing
    Exit Function
   
ErrorHandler:
    Call HandleError(Err.Number)
    Err.Clear
    Resume Exit_Function
End Function
' Function HowManyDataSQL
' -----------------------------

' @Function GetDataSQL
' -----------------------------
'@Author Roland Ross Hadi
'@Description Based on a SQL statement, return the required value. If nothing is
'               found, the function return -2 (we use -2 insted of -1, because if
'               the funtion return a boolean value (True/False, Yes/No) the numeric
'@Comments
' Parameter:
'         sSQL : SQL statement (Ex.:Select City_Name From City Where City=1)
Public Function GetDataSQL(sSQL As String) As Variant
    On Error GoTo ErrorHandler
   
    Dim rs As ADODB.Recordset
   
    'Create a new instance or a recordset
    Set rs = New ADODB.Recordset
   
    'Open the recordset based on the sSQL
    rs.Open sSQL, cnn
   
    'If the data was found, return the first field. Otherwise
    'return -2
    If Not isRSEmpty(rs) Then
        GetDataSQL = rs.Fields(0)
    Else
        GetDataSQL = -2
    End If
   
Exit_Function:
    'Destroy the recordset object
    Set rs = Nothing
    Exit Function
ErrorHandler:
MsgBox Err.Description
    Call HandleError(Err.Number)
    Err.Clear
    Resume Exit_Function
End Function
' Function GetDataSQL
' -----------------------------

' @Function GetDataSQL
' -----------------------------
'@Author Roland Ross Hadi
'@Description Based on a SQL statement, return the required value. If nothing is
'               found, the function return -2 (we use -2 insted of -1, because if
'               the funtion return a boolean value (True/False, Yes/No) the numeric
'@Comments
' Parameter:
'         sSQL : SQL statement (Ex.:Select City_Name From City Where City=1)
Public Function GetRecordSetSQL(sSQL As String) As Variant
    On Error GoTo ErrorHandler
   
    Dim rs As ADODB.Recordset
   
    'Create a new instance or a recordset
    Set rs = New ADODB.Recordset
   
    'Open the recordset based on the sSQL
    rs.Open sSQL, cnn
   
    'If the data was found, return the first field. Otherwise
    'return -2
    
        Set GetRecordSetSQL = rs
   
Exit_Function:
    'Destroy the recordset object
    Set rs = Nothing
    Exit Function
ErrorHandler:
MsgBox Err.Description
    Call HandleError(Err.Number)
    Err.Clear
    Resume Exit_Function
End Function
' Function GetDataSQL
' -----------------------------

' @Function FilllstcboBox
' -----------------------------
'@Author Roland Ross Hadi
'@Description This procedure use a SQL statement to fill a ListBox or a ComboBox
'@Comments
' Parameter:
'        frm    - Form holding the [ListBox] or [ComboBox]
'        ctrlName - Name of the [ListBox] or [ComboBox]
'        sSql    - SQL Statement. The statement must have two fields, an ID
'                    number and a value to display on the screen. For example:
'                    "Select City, City_Name From City"
Public Sub FilllstcboBox(frm As Form, ctrlName As String, sSQL As String, Optional cboIndex As Integer = -1)
    On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim iField0 As Integer
    Dim sField1 As String
   
    'Create a new instance or a recordset
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenStatic
    'Open the recordset based on the sSQL
    rs.Open sSQL, cnn
   
    'Be sure to be at the beggining of the recordset
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
   
    'Until you reach the end of the recordset, fill the control
    While Not rs.EOF
   
        If rs.Fields.Count = 1 Then
            iField0 = 1
            sField1 = IIf(IsNull(rs.Fields(0)), "", rs.Fields(0))
        Else
            iField0 = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
            sField1 = IIf(IsNull(rs.Fields(1)), "", rs.Fields(1))
        End If
   
   
        If Len(sField1) > 1 Then
            If cboIndex < 0 Then
                frm.Controls(ctrlName).AddItem sField1
                frm.Controls(ctrlName).ItemData(frm.Controls(ctrlName).NewIndex) = iField0
            Else
                'same procedure for an array of combo/list box
                frm.Controls(ctrlName)(cboIndex).AddItem sField1
                frm.Controls(ctrlName)(cboIndex).ItemData(frm.Controls(ctrlName)(cboIndex).NewIndex) = iField0
            End If
        End If
        rs.MoveNext
    Wend
   
    'Close the recordset
    rs.Close
   
Exit_Sub:
    'Destroy the recordset object
    Set rs = Nothing
    Exit Sub
   
ErrorHandler:
    Debug.Print "*** FILLLSTCBOBOX (clsDatabase) ************************************************"
    Debug.Print Time() & ": ERROR # " & Err.Number
    Debug.Print "Description : " & Err.Description
    Debug.Print "SQL = " & sSQL
    Debug.Print "***************************************************************************"
    Call HandleError(Err.Number)
    Err.Clear
    Resume Exit_Sub
End Sub
' Function FilllstcboBox
' -----------------------------

' @Function isRSEmpty
' -----------------------------
'@Author Roland Ross Hadi
'@Description Use to check if a recordset is empty or not
'@Comments
' Parameter:
'        RS: Recordset Object
Public Function isRSEmpty(rs As ADODB.Recordset) As Boolean
    isRSEmpty = ((rs.BOF = True) And (rs.EOF = True))
End Function
' Function isRSEmpty
' -----------------------------

' @Function ExecuteSQL
' -----------------------------
'@Author Roland Ross Hadi
'@Description Run a sql statement
'@Comments
' Parameter:
'        sSQL : Sql statement (use for Insert, Update or Delete)
Public Function ExecuteSQL(sSQL As String) As Long
On Error GoTo ErrorHandler
   
    If GetSetting(App.Title, "Control", "DebugMode") = "TRUE" Then
        Debug.Print sSQL
    End If
       
    Call cnn.Execute(sSQL)
    'Debug.Print cnn.ConnectionString
    'Assing the error number as return value, if there is no error, it will
    'be 0
    ExecuteSQL = Err.Number
Exit_Sub:
    Exit Function
   
ErrorHandler:
    Debug.Print "***************************************************************************"
    Debug.Print Time() & ": ERROR # " & Err.Number
    Debug.Print "Description : " & Err.Description
    Debug.Print "SQL = " & sSQL
    Debug.Print "***************************************************************************"
    ExecuteSQL = Err.Number
    Err.Clear
    Resume Exit_Sub
End Function
' Function ExecuteSQL
' -----------------------------

' @Function HandleError
' -----------------------------
'@Author Roland Ross Hadi
'@Description Process the errors
'@Comments
' Parameter:
'        ErrNumber : The system error number
Private Sub HandleError(ErrNumber As Long)
    Dim sMsg As String
    'Based on the error number
    Select Case ErrNumber
        Case -2147467259
            sMsg = "Cannont access the table, probably locked or open in design monde"
        Case Else
            sMsg = Err.Description
            Debug.Print Time() & ":" & Err.Number & "-" & sMsg
            sMsg = ""
    End Select
   
    'If the sMsg is not empty
    If sMsg <> "" Then MsgBox sMsg
End Sub
' Function HandleError
' -----------------------------

' @Function GetSQL
' -----------------------------
'@Author Roland Ross Hadi
'@Description Read from the disk SQL statements saved in a files
'@Comments
' Parameter:
'        sqlFileName : Name of the SQL File
Public Function GetSQL(sqlFileName As String) As String
    Dim sTmp, sSQL As String
    sqlFileName = "SQL\" & sqlFileName & ".SQL"
    Open sqlFileName For Input As #1
        While Not EOF(1)
            Line Input #1, sTmp
            sSQL = sSQL & Chr(32) & sTmp
        Wend
    Close #1
    GetSQL = Trim(sSQL)
End Function
' Function GetSQL
' -----------------------------

' @Function isAdmin
' -----------------------------
'@Author Roland Ross Hadi
'@Description Is the user have administration rights
'@Comments
' Parameter:
'        UserName : User name to check
Public Function isAdmin(UserName As String) As Boolean
    isAdmin = GetDataSQL("Select [ADM_Right] FROM [User] WHERE [Usr_UserName]=" & _
            PutQuotes(UserName)) = True
End Function
' Function isAdmin
' -----------------------------

' @Function GetUserPassword
' -----------------------------
'@Author Roland Ross Hadi
'@Description Get the user password
'@Comments
' Parameter:
'        UserName : User name
Public Function GetUserPassword(UserName As String, Optional fldName As String, Optional fldPass As String, Optional tblName As String) As String
    On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim strsql As String
   
    If fldName <> "" Then
        strsql = "SELECT [" & fldPass & "] " & _
                "FROM [" & tblName & "] " & _
                "WHERE [" & fldName & "] = " & PutQuotes(UserName)
    Else
        strsql = "SELECT [Usr_Password] " & _
                "FROM [USER] " & _
                "WHERE [Usr_UserName] = " & PutQuotes(UserName)
    End If
    Debug.Print strsql
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenStatic
    rs.Open strsql, cnn
   
    If rs.RecordCount <= 0 Then
        GetUserPassword = sNoPassword
    Else
        GetUserPassword = rs.Fields(0)
    End If
   
Exit_Function:
    rs.Close
    'Call CloseDBADO
    Set rs = Nothing
    Exit Function
   
ErrorHandler:
    MsgBox Err.Number & vbCrLf & Err.Description
    Err.Clear
    Resume Exit_Function
End Function

Private Function GetNoPasswordMessage() As String
    GetNoPasswordMessage = sNoPassword
End Function
' Function GetUserPassword
' -----------------------------

' @Function DropTable
' -----------------------------
'@Author Roland Ross Hadi
'@Description Delete a table
'@Comments
' Parameter:
'        tblName : Name of the table to delete
Public Sub DropTable(tblName As String)
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim strsql As String
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    strsql = "DROP TABLE " & tblName
    rs.Open strsql, cnn
Exit_Sub:
    'RS.Close
    Set rs = Nothing
    Exit Sub
ErrorHandler:
    'MsgBox Err.Number & " " & Err.Description
    Err.Clear
End Sub
' Function DropTable
' -----------------------------

'' @Function CreateDatabase
'' -----------------------------
''@Author Roland Ross Hadi
''@Description Create a new MDB file
''@Comments
'' Parameter:
''        dbName    : Name of the MDB file
''        DBVersion : Enum showing the possibility to create the MDB for
''                    Access 2000 or Access 97
'Public Sub CreateDatabase(DbName As String, DBVersion As AccessVersion)
'    Dim newDB As New ADOX.Catalog
'    If DBVersion = Access2000 Then
'        'For ACCESS 2000
'        newDB.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'                    "Data Source=" & App.Path & "\" & DbName
'    Else
'        'For ACCESS 97
'        newDB.Create "Provider=Microsoft.Jet.OLEDB.3.51;" & _
'                    "Data Source=" & App.Path & "\" & DbName
'    End If
'    Set newDB = Nothing
'End Sub
'' Function CreateDatabase
'' -----------------------------

' @Function CreateTable
' -----------------------------
'@Author Roland Ross Hadi
'@Description reate a new table in the database
'@Comments
' Parameter:
'        dbName    : Name of the MDB file
'        tblName : Name of the new table
'        fldName : Name of the fields and their DataType (fldName,DataType
'                    ex.:Cust_Name, Text)
'                    Access 2000 or Access 97
Public Sub CreateTable(DbName As String, tblName As String, fldName() As String, Optional PrimaryKey As String)
    Dim adoc As New clsDatabase
    Dim sSQL As String
    Dim iCount As Integer
    Dim sField As String
    Dim sDataType As String

    sSQL = "CREATE TABLE " & tblName & " ("
    For iCount = 0 To UBound(fldName()) - 1
        Call SplitFieldDataType(fldName(iCount), sField, sDataType)
        sSQL = sSQL & sField & " " & sDataType
        If iCount < UBound(fldName()) - 1 Then
            sSQL = sSQL & ","
        End If
    Next iCount
    If PrimaryKey = "" Then
        sSQL = sSQL & ")"
    Else
        sSQL = sSQL & "," & PrimaryKey & ")"
    End If
    Call adoc.OpenDBADO(App.path & "\" & DbName)
    'Debug.Print sSQL
    Call adoc.ExecuteSQL(sSQL)
    Call adoc.CloseDBADO
    Set adoc = Nothing
End Sub
' Function CreateTable
' -----------------------------

' @Function SplitFieldDataType
' -----------------------------
'@Author Roland Ross Hadi
'@Description Split a string in two part, the field name and his data type
'@Comments
' Parameter:
'        sFldDT    : The string to separate (ex:"Cust_Name|String(20)")
'        sField    : Storing in the address of the calling variable(ByRef),
'                    the name of the field ("Cust_Name")
'        sDataType : Storing in the address of the calling variable(ByRef),
'                    the data type of the field ("String(20)")
Private Sub SplitFieldDataType(sFldDT As String, ByRef sField As String, ByRef sDataType As String)
    sField = Left(sFldDT, InStr(1, sFldDT, "|") - 1)
    sDataType = Mid(sFldDT, InStr(1, sFldDT, "|") + 1)
End Sub
' Function SplitFieldDataType
' -----------------------------

' @Function CompactDatabase
' -----------------------------
'@Author Roland Ross Hadi
'@Description Compact a database (MDB)
'@Comments
' Parameter:
'        dbPathAndName : The path and name of the DB to compact. If nothing is
'                        set for this variable, the system will take the current
'                        DB (set in GetDBInfo)
'        MakeBackup    : If its true make a copy of the DB
Public Sub CompactDatabase(Optional dbPathAndName As String, Optional MakeBackup As Boolean)
    Dim JRO As New JRO.JetEngine
    If dbPathAndName <> "" Then
        If MakeBackup Then Call FileCopy(dbPathAndName, dbPathAndName & ".bak")
        JRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                            "Data Source=" & dbPathAndName, _
                            "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                            "Data Source=" & App.path & "\dtnetCompact.mdb;" & _
                            "Jet OLEDB:Engine Type=4"
        Call Kill(dbPathAndName)
        Name App.path & "\BakCompact.mdb" As dbPathAndName
    Else
        Call GetDBInfo
        If MakeBackup Then Call FileCopy(dbPath & dbFileName, dbPath & dbFileName & ".bak")
        JRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                            "Data Source=" & dbPath & dbFileName, _
                            "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                            "Data Source=" & App.path & "\dtnetCompact.mdb;" & _
                            "Jet OLEDB:Engine Type=4"
        Call Kill(dbPath & dbFileName)
        Name dbPath & "\BakCompact.mdb" As dbPath & dbFileName
    End If
    Set JRO = Nothing
End Sub
' Function CompactDatabase
' -----------------------------

' @Function PutQuotes
' -----------------------------
'@Author Roland Ross Hadi
'@Description This function surround a string by quotes. You need to call
'               this procedure each time you want to put a string value inside
'@Comments
' Parameter:
'        txtToQuote : String to surround with quotes
Public Function PutQuotes(ByRef txtToQuote As String, Optional isSQLServer As Boolean) As String
    Dim iASCII As Integer
   
    If isSQLServer Then
        iASCII = 39
    Else
        iASCII = 34
    End If
   
    PutQuotes = Chr(iASCII) & txtToQuote & Chr(iASCII)
End Function
' Function PutQuotes
' -----------------------------

' @Function ExportToCSV
' -----------------------------
'@Author Roland Ross Hadi
'@Description Export a SQL result to a Comma-Separated Value File (CSV)
'@Comments
' Parameter:
'        sPath    : Path of the exported CSV file
'        sFileName : The CSV filename
'        sSQL    : The query result to export
Public Sub ExportToCSV(sPath As String, sFileName As String, sSQL As String)
    Dim iPortFile As Integer
    Dim rs As ADODB.Recordset
    Dim iCount As Integer
    Dim sField As String
   
    'Create a new instance or a recordset
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenStatic
    'Open the recordset based on the sSQL
    rs.Open sSQL, cnn
   
    iPortFile = FreeFile
   
    'Be sure to be at the beggining of the recordset
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
   
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    If InStr(1, LCase(sFileName), ".csv") = 0 Then sFileName = sFileName & ".csv"
   
    Open sPath & sFileName For Output As #iPortFile
   
    For iCount = 0 To rs.Fields.Count - 1
        sField = "" & rs.Fields(iCount).Name
        Write #iPortFile, sField;
    Next iCount
   
    Write #iPortFile,
    rs.MoveFirst
   
    While Not rs.EOF
        For iCount = 0 To rs.Fields.Count - 1
            sField = "" & rs.Fields(iCount)
            Write #iPortFile, sField;
        Next iCount
        Write #iPortFile,
        rs.MoveNext
    Wend
   
    rs.Close
    Close #iPortFile
   
    Set rs = Nothing
End Sub
' @Function ExportToCSV
' -----------------------------

' @Function GetPreviousNextID
' -----------------------------
'@Author Roland Ross Hadi
'@Description Get the previous or the next ID of the current one
'@Comments
' Parameter:
'        sSQL    : The query
'        lCurrent : The current ID
'        iDirection : If you need the Previous or the Next
Public Function GetPreviousNextID(sSQL As String, lCurrentID As Long, iDirection As eDirection) As Long
    Dim rs As ADODB.Recordset
   
    'Create a new instance or a recordset
    Set rs = New ADODB.Recordset
   
    'Allow the recordset to be searched in any direction
    rs.CursorType = adOpenKeyset
   
    'Open the recordset based on the sSQL
    rs.Open sSQL, cnn
   
   
    If isRSEmpty(rs) Then
        'If the data wasn't found, return -2
        GetPreviousNextID = -2
    Else
        rs.find rs.Fields(0).Name & "=" & lCurrentID
        Select Case iDirection
            Case eDirection.ePrevious
                rs.MovePrevious
                If Not rs.BOF Then
                    GetPreviousNextID = rs.Fields(0)
                Else
                    GetPreviousNextID = lCurrentID
                End If
            Case eDirection.eNext
                If Not rs.EOF Then rs.MoveNext
                If Not rs.EOF Then
                    GetPreviousNextID = rs.Fields(0)
                Else
                    GetPreviousNextID = lCurrentID
                End If
        End Select
       
    End If
   
    rs.Close
    Set rs = Nothing
End Function
' Function GetPreviousNextID
' -----------------------------

' @Function DoubleSingleQuote
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns Double Single Quote text
'@Comments
' Parameter:
'        sText    : The query
Public Function DoubleSingleQuote(sText As String) As String
    Dim iCount As Integer
    Dim sTmp As String
   
    For iCount = 1 To Len(sText)
        If Mid(sText, iCount, 1) = Chr(39) Then
            sTmp = sTmp & Mid(sText, iCount, 1) & Chr(39)
        Else
            sTmp = sTmp & Mid(sText, iCount, 1)
        End If
       
    Next iCount
    DoubleSingleQuote = sTmp
End Function
' Function DoubleSingleQuote
' -----------------------------

' @Function ReplaceCharBy
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns ReplaceCharBy
'@Comments
' Parameter:
'        sText    :   The query
'        sSource    : source
'        sByWhat    : replace value
Public Function ReplaceCharBy(sText As String, sSource As String, sByWhat As String) As String
    Dim iCount As Integer
    For iCount = 1 To Len(sText)
        If Mid(sText, iCount, 1) = sSource Then Mid(sText, iCount, 1) = sByWhat
    Next iCount
    ReplaceCharBy = sText
End Function
' Function ReplaceCharBy
' -----------------------------

' @Function CreateNextID
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns CreateNextID
'@Comments
' Parameter:
'        sTbl    : Table
'        sFld    : Field
'        cAdoLocal    : Adodb object
Public Function CreateNextID(sTbl As String, sFld As String, cAdoLocal As clsDatabase) As Long
On Error GoTo ErrorHandler
    Dim sSQL As String
    Dim lID As Long
   
    'Check if the connection was establish
    Debug.Print cnn.State
   
    sSQL = "SELECT TOP 1 " & sFld & " FROM " & sTbl & " ORDER BY " & sFld & " DESC"
    lID = cAdoLocal.GetDataSQL(sSQL)
    If lID <= 0 Then
        lID = 1
    Else
        lID = lID + 1
    End If
   
    CreateNextID = lID
   
Exit_Function:
    Exit Function
ErrorHandler:
    If Err.Number = 91 Then
        MsgBox "Aucune connection n'a été établi!"
        Err.Clear
        Resume Exit_Function
    End If
End Function
' Function CreateNextID
' -----------------------------

' @Function isTableExist
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns CreateNextID
'@Comments
' Parameter:
'        cAdotmp    : Adodb object
'        sTable    :  Table
Public Function isTableExist(cAdotmp As clsDatabase, sTable As String) As Boolean
On Error GoTo ErrorHandler
    Dim sSQL As String
    Dim Rstmp As New ADODB.Recordset
    sSQL = "SELECT TOP 1 * FROM " & sTable
    Rstmp.Open sSQL, cAdotmp.cnn
    isTableExist = True
Exit_Function:
    If Rstmp.State = 1 Then Rstmp.Close
    Set Rstmp = Nothing
    Exit Function
   
ErrorHandler:
    If Err.Number = -2147217865 Then isTableExist = False
    Err.Clear
    Resume Exit_Function
End Function
' Function isTableExist
' -----------------------------

' @Function ConnectToSQLDB
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns CreateNextID
'@Comments
' Parameter:
'        strUserID    :  Database user id
'        strPassword    :  User password
'        strDataSource    :  DNS datasource
Public Function ConnectToSQLDB(strUserID As String, strPassword As String, strDataSource As String) As Boolean
  On Error GoTo Err1
  Set cnn = New ADODB.Connection
  cnn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=true;user id=" & strUserID & ";password=" & strPassword & ";Data Source=" & strDataSource & ""
  cnn.CursorLocation = ADODB.CursorLocationEnum.adUseClient
  cnn.Open
  ConnectToSQLDB = True
  Exit Function
Err1:
  ConnectToSQLDB = False
End Function
' Function ConnectToSQLDB
' -----------------------------

' @Function ConnectToMDB
' -----------------------------
'@Author Roland Ross Hadi
'@Description Returns CreateNextID
'@Comments
' Parameter:
'        cAdotmp    : Adodb object
'        sTable    :  Table
Public Function ConnectToMDB(strDataSource As String, Optional strUserID As String, Optional strPassword As String) As Boolean
  On Error GoTo Err1
  Set cnn = New ADODB.Connection
  If Trim(strUserID) <> "" And Trim(strPassword) <> "" Then
    cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=true;user id=" & strUserID & ";password=" & strPassword & ";Data Source=" & strDataSource
  ElseIf Trim(strPassword) <> "" Then
    cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=" & strPassword & ";Data Source=" & strDataSource
  Else
    cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDataSource
  End If
  cnn.CursorLocation = ADODB.CursorLocationEnum.adUseClient
  cnn.Open
  ConnectToMDB = True
  Exit Function
Err1:
  ConnectToMDB = False
End Function
' Function ConnectToMDB
' -----------------------------
