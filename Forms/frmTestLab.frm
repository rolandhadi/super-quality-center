VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestLab 
   Caption         =   "Data Scripting Module"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   12675
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLoadExcel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Picture         =   "frmTestLab.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Import step description and expected results from an excel file"
      Top             =   1620
      Width           =   2205
   End
   Begin VB.ListBox lstUpdateFields 
      Columns         =   4
      Height          =   735
      Left            =   4560
      Style           =   1  'Checkbox
      TabIndex        =   3
      ToolTipText     =   "Select Values to update"
      Top             =   780
      Width           =   7995
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdRefresh"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdGenerate"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdOutput"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdUpload"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6060
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11940
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestLab.frx":07A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestLab.frx":0A38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView QCTree 
      Height          =   5355
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9446
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   3915
      Left            =   4560
      TabIndex        =   4
      Top             =   2040
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   6906
      _Version        =   393216
      Cols            =   10
      WordWrap        =   -1  'True
      AllowUserResizing=   3
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   11400
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestLab.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestLab.frx":13DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestLab.frx":1AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestLab.frx":2200
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgOpenExcel 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Microsoft Excel File | *.xls*"
   End
   Begin VB.Label Label1 
      Caption         =   "Fields to Update"
      Height          =   195
      Left            =   4560
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmTestLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GenerateOutput()
Dim rs As TDAPIOLELib.Recordset
Dim AllScript
Dim objCommand
Dim i
Dim strPath
Dim iterd
Dim NewVal
Dim iter
Dim TimeSt
Dim AllF
Dim Decompile
Dim p
    
    TimeSt = Format(Now, "mmddyyyy HHMM AMPM") & "-"
    AllF = "Data Scripting"
    
    If Left(QCTree.SelectedItem.Key, 1) = "F" Then
        strPath = "CF_ITEM_PATH LIKE '" & GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%'"
    Else
        strPath = "TC_CYCLE_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
    End If
    
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT TC_TESTCYCL_ID, RTP_ID, '1' AS ""Iteration"",  RTP_ORDER, RTP_NAME, RTP_BPTA_LONG_VALUE, '' AS ""RTP_ACTUAL_VALUE"", CF_ITEM_NAME, CY_CYCLE, TS_NAME, TC_DATA_OBJ FROM RUNTIME_PARAM, TEST, TESTCYCL, CYCLE, CYCL_FOLD WHERE TC_TEST_ID = TS_TEST_ID AND RTP_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " ORDER BY CF_ITEM_ID, TC_CYCLE_ID, TC_TEST_ORDER, RTP_TEST_ID, RTP_ORDER "
    Set rs = objCommand.Execute
    AllScript = "Test Instance ID" & vbTab & "RunTime ID" & vbTab & "Iteration" & vbTab & "Parameter Order" & vbTab & "Parameter Name" & vbTab & "Default Value" & vbTab & "Actual Value" & vbTab & "Folder Name" & vbTab & "Test Set" & vbTab & "Test Instance"
    FileWrite "C:\ALL SCRIPTS\" & TimeSt & AllF & ".xls", AllScript
    AllScript = ""
    ClearTable
    flxImport.Rows = rs.RecordCount + 1
    For i = 1 To rs.RecordCount
        If i = 1 Then FileWrite "C:\ALL SCRIPTS\" & TimeSt & AllF & ".xls", AllScript
                stsBar.SimpleText = "Processing " & i & " out of " & rs.RecordCount
                AllScript = AllScript & vbCrLf & rs.FieldValue("TC_TESTCYCL_ID") & vbTab & rs.FieldValue("RTP_ID") & vbTab & 1 & vbTab & rs.FieldValue("RTP_ORDER") & vbTab & rs.FieldValue("RTP_NAME") & vbTab & rs.FieldValue("RTP_BPTA_LONG_VALUE") & vbTab & "" & vbTab & rs.FieldValue("CF_ITEM_NAME") & vbTab & rs.FieldValue("CY_CYCLE") & vbTab & rs.FieldValue("TS_NAME")
                flxImport.TextMatrix(i, 0) = rs.FieldValue("TC_TESTCYCL_ID")
                flxImport.TextMatrix(i, 1) = rs.FieldValue("RTP_ID")
                flxImport.TextMatrix(i, 2) = 1
                flxImport.TextMatrix(i, 3) = rs.FieldValue("RTP_ORDER")
                flxImport.TextMatrix(i, 4) = rs.FieldValue("RTP_NAME")
                flxImport.TextMatrix(i, 5) = rs.FieldValue("RTP_BPTA_LONG_VALUE")
                flxImport.TextMatrix(i, 6) = ""
                flxImport.TextMatrix(i, 7) = rs.FieldValue("CF_ITEM_NAME")
                flxImport.TextMatrix(i, 8) = rs.FieldValue("CY_CYCLE")
                flxImport.TextMatrix(i, 9) = rs.FieldValue("TS_NAME")
        If i Mod 2500 = 0 Then
            FileAppend "C:\ALL SCRIPTS\" & TimeSt & AllF & ".xls", AllScript
            AllScript = ""
        End If
        rs.Next
    Next
FileAppend "C:\ALL SCRIPTS\" & TimeSt & AllF & ".xls", AllScript
stsBar.SimpleText = "Ready"
End Sub

Private Sub cmdLoadExcel_Click()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim fname As String
Dim LastRow
On Error Resume Next
    xlWB.Close
    xlObject.Application.Quit
On Error GoTo 0
    dlgOpenExcel.ShowOpen
    fname = dlgOpenExcel.filename
    If fname = "" Then Exit Sub
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Open(fname) 'Open your book here
                
    Clipboard.Clear
    With xlObject.ActiveWorkbook.ActiveSheet
         LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
        .Range("A1:J" & LastRow).Copy 'Set selection to Copy
    End With
       
    With flxImport
        .Clear
        .Redraw = False     'Dont draw until the end, so we avoid that flash
        .Row = 0            'Paste from first cell
        .Col = 0
        .Rows = LastRow
        .Cols = 10
        .RowSel = LastRow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
        .ColSel = 10 - 1
    End With
    
     With flxImport
        .Clear
        .Redraw = False     'Dont draw until the end, so we avoid that flash
        .Row = 0            'Paste from first cell
        .Col = 0
        .Rows = LastRow
        .Cols = 10
        .RowSel = LastRow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
        .ColSel = 10 - 1
        .Clip = Replace(Clipboard.GetText, vbNewLine, vbCr)   'Replace carriage return with the correct one
        .Col = 1            'Just to remove that blue selection from Flexgrid
        .Redraw = True      'Now draw
    End With
        
    xlObject.DisplayAlerts = False 'To avoid "Save woorkbook" messagebox
    
    'Close Excel
    xlWB.Close
    xlObject.Application.Quit
    Set xlWB = Nothing
    Set xlObject = Nothing
End Sub

Private Sub flxImport_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 67 And Shift = vbCtrlMask Then
    Clipboard.Clear
    Clipboard.SetText flxImport.Clip
End If
End Sub

Private Sub Form_Load()
ClearForm
End Sub

Private Sub Form_Resize()
QCTree.Height = stsBar.Top - 550
lstUpdateFields.Width = Me.Width - lstUpdateFields.Left - 350
flxImport.Height = stsBar.Top - flxImport.Top - 250
flxImport.Width = Me.Width - flxImport.Left - 350
End Sub

Private Sub QCTree_DblClick()
'Dim rs As TDAPIOLELib.Recordset
'Dim objCommand
'Dim i As Long
'Dim nodx As Node
'    If QCTree.SelectedItem.Children <> 0 Then Exit Sub
'    Set objCommand = QCConnection.Command
'    objCommand.CommandText = "SELECT FC_ID, FC_NAME FROM COMPONENT_FOLDER WHERE FC_FATHER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY FC_NAME"
'    Set rs = objCommand.Execute
'    For i = 1 To rs.RecordCount
'        QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("F" & rs.FieldValue("FC_ID")), rs.FieldValue("FC_NAME")
'        rs.Next
'    Next
'
'    Set objCommand = QCConnection.Command
'    objCommand.CommandText = "SELECT DISTINCT CO_NAME, CO_ID FROM COMPONENT, COMPONENT_FOLDER WHERE CO_FOLDER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY CO_NAME"
'    Set rs = objCommand.Execute
'    For i = 1 To rs.RecordCount
'        QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("C" & rs.FieldValue("CO_ID")), rs.FieldValue("CO_NAME")
'        rs.Next
'    Next

Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
Dim nodx As Node

    If QCTree.SelectedItem.Children <> 0 Then Exit Sub
    
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT DISTINCT CY_CYCLE, CY_CYCLE_ID FROM CYCLE, CYCL_FOLD WHERE CY_FOLDER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY CY_CYCLE"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("C" & rs.FieldValue("CY_CYCLE_ID")), rs.FieldValue("CY_CYCLE"), 2
        rs.Next
    Next
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "cmdRefresh"
    stsBar.SimpleText = "Preparing the process..."
    ClearForm
    stsBar.SimpleText = "Ready"
Case "cmdGenerate"
    stsBar.SimpleText = "Preparing the process..."
    GenerateOutput
Case "cmdOutput"
    If flxImport.Rows <= 1 Then
        MsgBox "Nothing to output", vbInformation
    Else
        If GetEditableFields() = "IV:IV" Then
            If MsgBox("You have selected nothing to update on this sheet. The whole sheet will be read-only. Do you want to proceed?", vbYesNo) = vbYes Then
                stsBar.SimpleText = "Preparing the process..."
                OutputTable
            End If
        Else
            stsBar.SimpleText = "Preparing the process..."
            OutputTable
        End If
    End If
Case "cmdUpload"
If flxImport.TextMatrix(0, 0) = "Test Instance ID" And flxImport.TextMatrix(0, 6) = "Actual Value" Then
    If MsgBox("Are you sure you want to mass update " & flxImport.Rows - 1 & " record(s) to the Run Tim Parameter?", vbYesNo) = vbYes Then
            stsBar.SimpleText = "Preparing the process..."
    End If
Else
    MsgBox "Invalid or Incorrect upload sheet file selected", vbCritical
End If
End Select
End Sub

Private Sub ClearForm()
ClearTable
QCTree.Nodes.Clear
'Dim rs As TDAPIOLELib.Recordset
'Dim objCommand
'Dim i As Long
'    QCTree.Nodes.Add , , "Root", "Root"
'    Set objCommand = QCConnection.Command
'    objCommand.CommandText = "SELECT FC_ID, FC_NAME FROM COMPONENT_FOLDER WHERE FC_FATHER_ID = 1 ORDER BY FC_NAME"
'    Set rs = objCommand.Execute
'    For i = 1 To rs.RecordCount
'        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("FC_ID")), rs.FieldValue("FC_NAME")
'        rs.Next
'    Next

Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
    QCTree.Nodes.Add , , "Root", "1Company", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = 1 ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    
    lstUpdateFields.Clear
    lstUpdateFields.AddItem "Actual Value"
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.TextMatrix(0, 0) = "Test Instance ID"
flxImport.TextMatrix(0, 1) = "RunTime ID"
flxImport.TextMatrix(0, 2) = "Iteration"
flxImport.TextMatrix(0, 3) = "Parameter Order"
flxImport.TextMatrix(0, 4) = "Parameter Name"
flxImport.TextMatrix(0, 5) = "Default Value"
flxImport.TextMatrix(0, 6) = "Actual Value"
flxImport.TextMatrix(0, 7) = "Folder Name"
flxImport.TextMatrix(0, 8) = "Test Set"
flxImport.TextMatrix(0, 9) = "Test Instance"
flxImport.Rows = 2
End Sub

Private Sub OutputTable()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim i, Protections
Dim curTab
Dim w


Set xlObject = New Excel.Application

On Error Resume Next
For Each w In xlObject.Workbooks
   w.Close savechanges:=False
Next w
On Error GoTo 0

Set xlWB = xlObject.Workbooks.Add

  'xlObject.Visible = True
  curTab = "RUN_TIME01"
  xlObject.Sheets("Sheet1").Name = curTab
  flxImport.FixedCols = 0
  flxImport.FixedRows = 0
  flxImport.RowSel = flxImport.Rows - 1
  flxImport.ColSel = flxImport.Cols - 1
  Clipboard.Clear
  Clipboard.SetText flxImport.Clip
  flxImport.FixedCols = 1
  flxImport.FixedRows = 1
  
  xlObject.Sheets(curTab).Range("A1").Select
  xlObject.Sheets(curTab).Paste

'On Error Resume Next
     xlObject.Sheets(curTab).Range("A:J").Select
     xlObject.Sheets(curTab).Range("A2:J" & flxImport.Rows).Sort Key1:=xlObject.Sheets(curTab).Range("A2"), Order1:=xlAscending, Key2:=xlObject.Sheets(curTab).Range("C2") _
        , Order2:=xlAscending, Key3:=xlObject.Sheets(curTab).Range("D2"), Order3:=xlAscending, Header:= _
        xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, DataOption3:= _
        xlSortNormal
        
    xlObject.Sheets(curTab).Range("A:J").Borders(xlDiagonalDown).LineStyle = xlNone
    xlObject.Sheets(curTab).Range("A:J").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlObject.Sheets(curTab).Range("A:J").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:J").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:J").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:J").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:J").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:J").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    xlObject.Sheets(curTab).Rows("1:1").Select
    With xlObject.Sheets(curTab).Rows("1:1").Interior
        .ColorIndex = 6
        .Pattern = xlSolid
    End With
    xlObject.Sheets(curTab).Rows("1:1").Font.Bold = True
    xlObject.Sheets(curTab).Range("A:J").Select
    xlObject.Sheets(curTab).Range("A:J").EntireColumn.AutoFit
    xlObject.Sheets(curTab).Range("A1").Select
    
    xlObject.Sheets(curTab).Range("A1").AddComment
    xlObject.Sheets(curTab).Range("A1").Comment.Visible = False
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & mdiMain.Caption & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""
  
  xlObject.Sheets(curTab).Protection.AllowEditRanges.Add Title:="Range1", Range:=xlObject.Sheets(curTab).Columns(GetEditableFields)
  xlObject.Sheets(curTab).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  xlObject.Workbooks(1).SaveAs "RUN_TIME01-" & Format(Now, "mmddyyyy HHMM AMPM")
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate
  
  Set xlWB = Nothing
  Set xlObject = Nothing
  
  stsBar.SimpleText = "Ready"
On Error GoTo 0
End Sub

Private Function GetEditableFields()
Dim i
Dim j
Dim tmp
For i = 0 To lstUpdateFields.ListCount - 1
   If lstUpdateFields.Selected(i) = True Then
       For j = 0 To flxImport.Cols - 1
          If lstUpdateFields.List(i) = flxImport.TextMatrix(0, j) Then
             tmp = tmp & Chr(65 + j) & ":" & Chr(65 + j) & ", "
          End If
       Next
   End If
Next
tmp = Trim(tmp)
If tmp <> "" Then
   GetEditableFields = Left(tmp, Len(tmp) - 1)
Else
   GetEditableFields = "IV:IV"
End If
End Function
