VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataTableManager 
   Caption         =   "DataTable Manager"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   14520
   Tag             =   "DataTable Manager"
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
      Left            =   4200
      Picture         =   "frmDataTableManager.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Import step description and expected results from an excel file"
      Top             =   600
      Width           =   2205
   End
   Begin VB.FileListBox fileExcel 
      Height          =   1455
      Left            =   60
      Pattern         =   "*.xls*"
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   600
      Width           =   4035
   End
   Begin VB.DirListBox dirFolders 
      Height          =   4815
      Left            =   60
      TabIndex        =   3
      Top             =   960
      Width           =   4035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdRefresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "cmdAdd"
            Object.ToolTipText     =   "Add New Library"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdGenerate"
            Object.ToolTipText     =   "Generate"
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "F_GET"
                  Text            =   "Get From Source File"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdOutput"
            Object.ToolTipText     =   "Output to Excel"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdUpload"
            Object.ToolTipText     =   "Upload to System"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   30
         Left            =   1080
         TabIndex        =   1
         Top             =   300
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   53
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   15240
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataTableManager.frx":07A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataTableManager.frx":0EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataTableManager.frx":12F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataTableManager.frx":1A08
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataTableManager.frx":211A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList_Sts 
      Left            =   15840
      Top             =   120
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
            Picture         =   "frmDataTableManager.frx":282C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataTableManager.frx":2B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataTableManager.frx":305F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataTableManager.frx":35B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6045
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmDataTableManager.frx":3B01
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   24386
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   4755
      Left            =   4200
      TabIndex        =   4
      Top             =   1020
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   8387
      _Version        =   393216
      Cols            =   4
      WordWrap        =   -1  'True
      AllowUserResizing=   3
   End
   Begin MSComDlg.CommonDialog dlgOpenExcel 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Microsoft Excel File | *.xls*"
   End
End
Attribute VB_Name = "frmDataTableManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type DataTableEntries
    DataTablePath As String
    ParameterName() As String
    ParameterValue() As String
End Type

Private Sub GenerateOutput()
Dim i, sArray() As String, tmp As String
ReDim sArray(0)
tmp = fileExcel.path
Call DirWalk("*.XLS", fileExcel.path, sArray)
dirFolders.path = tmp
        For i = LBound(sArray) To UBound(sArray) - 1
            ReadDataTableFiles sArray(i)
        Next
stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = flxImport.Rows - 1 & " record(s) generated"
End Sub

Function GetCommentText(rCommentCell As Range)
     Dim strGotIt As String
         On Error Resume Next
         strGotIt = WorksheetFunction.Clean _
             (rCommentCell.Comment.Text)
         GetCommentText = strGotIt
         On Error GoTo 0
End Function

Private Sub cmdLoadExcel_Click()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim fname As String
Dim lastrow
Dim i, j
On Error Resume Next
    xlWB.Close
    xlObject.Application.Quit
On Error GoTo 0
On Error GoTo ErrLoad
    dlgOpenExcel.filename = "": dlgOpenExcel.ShowOpen
    fname = dlgOpenExcel.filename
    If fname = "" Then Exit Sub Else Me.Caption = Me.Caption & " (" & dlgOpenExcel.FileTitle & ")"
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Open(fname) 'Open your book here
                
    Clipboard.Clear
    
    With xlObject.ActiveWorkbook.ActiveSheet
         If UCase(Trim(curDomain & "-" & curProject)) <> UCase(Trim(xlObject.ActiveWorkbook.Sheets(2).Range("B7").Value)) Then
            MsgBox "The spreadsheet is from a different Domain or Project"
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
         If InStr(1, GetCommentText(.Range("A1")), "DataTable") = 0 Then
            MsgBox "Import file is invalid. Please use only sheets generated by the SuperQC"
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
         lastrow = .Range("A" & .Rows.Count).End(xlUp).row
        .Range("A1:D" & lastrow).Copy 'Set selection to Copy
    End With
       
    With flxImport
        .Clear
        .Redraw = False     'Dont draw until the end, so we avoid that flash
        .row = 0            'Paste from first cell
        .col = 0
        .Rows = lastrow
        .Cols = 4
        .RowSel = lastrow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
        .ColSel = 4 - 1
    End With
    
     With flxImport
        .Clear
        .Redraw = False     'Dont draw until the end, so we avoid that flash
        .row = 0            'Paste from first cell
        .col = 0
        .Rows = lastrow
        .Cols = 4
        .RowSel = lastrow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
        .ColSel = 4 - 1
        .Clip = Replace(Clipboard.GetText, vbNewLine, vbCr)   'Replace carriage return with the correct one
        .col = 1            'Just to remove that blue selection from Flexgrid
        .Redraw = True      'Now draw
    End With
        
    xlObject.DisplayAlerts = False 'To avoid "Save woorkbook" messagebox
    
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = flxImport.Rows - 1 & " record(s) loaded"
    
    xlObject.DisplayAlerts = False
    xlWB.Close
    xlObject.Quit
    Set xlObject = Nothing
    Set xlWB = Nothing
    mdiMain.pBar.Max = 100
    mdiMain.pBar.Value = 100
Exit Sub
ErrLoad:
MsgBox "There was an error while importing the file. Please refresh and close all excel and try again" & vbCrLf & Err.Description, vbCritical
    xlObject.DisplayAlerts = False
    xlWB.Close
    xlObject.Quit
    Set xlObject = Nothing
    Set xlWB = Nothing
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpR
Select Case Button.Key
Case "cmdRefresh"
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
    ClearForm
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
Case "cmdGenerate"
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
    On Error GoTo OutputErr
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
    ClearTable
    GenerateOutput
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = flxImport.Rows - 1 & " parameters loaded successfully"
    Exit Sub
OutputErr:
    MsgBox "Data was been truncated because of an error." & vbCrLf & Err.Description
Case "cmdOutput"
    If flxImport.Rows <= 1 Then
        MsgBox "Nothing to output", vbInformation
    Else
            If flxImport.TextMatrix(1, 0) = "" Then Exit Sub
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
            OutputTable
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export to excel completed"
    End If
Case "cmdUpload"
If flxImport.TextMatrix(1, 0) = "" Then Exit Sub
If IncorrectHeaderDetails = False Then
        If MsgBox("Are you sure you want to mass update " & flxImport.Rows - 1 & " record(s) to the " & dirFolders.List(dirFolders.ListIndex) & "?", vbYesNo) = vbYes Then
            Randomize: tmpR = CInt(Rnd(1000) * 10000)
            If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
'                On Error Resume Next
'                ZipFolder dirFolders.List(dirFolders.ListIndex), App.path & "\SQC Logs" & "\" & "DataTable_LastRun"
'                On Error GoTo 0
                LoadToDataTables
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = flxImport.Rows - 1 & " parameters saved successfully"
                QCConnection.SendMail "user@companyemail.com", "", "[HPQC UPDATES] DataTable data records loaded successfully by " & curUser & " in " & curDomain & "-" & curProject, flxImport.Rows - 1 & " DataTable data records loaded successfully" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
                QCConnection.SendMail curUser, "", "[HPQC UPDATES] DataTable data records loaded successfully by " & curUser & " in " & curDomain & "-" & curProject, flxImport.Rows - 1 & " DataTable data records loaded successfully" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
            Else
                MsgBox "Invalid pass key", vbCritical
            End If
        End If
    Else
        MsgBox "No fields to update", vbCritical
    End If
End Select
End Sub

Private Sub LoadToDataTables()
Dim tmp() As DataTableEntries
Dim i, x, Y
ReDim tmp(x)
ReDim tmp(x).ParameterName(Y)
ReDim tmp(x).ParameterValue(Y)
x = -1
Y = -1
For i = 1 To flxImport.Rows - 1
    If flxImport.TextMatrix(i, 1) <> flxImport.TextMatrix(i - 1, 1) Then
        x = x + 1
        Y = 0
        ReDim Preserve tmp(x)
        tmp(x).DataTablePath = flxImport.TextMatrix(i, 1)
        ReDim Preserve tmp(x).ParameterName(Y)
        ReDim Preserve tmp(x).ParameterValue(Y)
        tmp(x).ParameterName(Y) = flxImport.TextMatrix(i, 2)
        tmp(x).ParameterValue(Y) = flxImport.TextMatrix(i, 3)
    Else
        Y = Y + 1
        ReDim Preserve tmp(x).ParameterName(Y)
        ReDim Preserve tmp(x).ParameterValue(Y)
        tmp(x).ParameterName(Y) = flxImport.TextMatrix(i, 2)
        tmp(x).ParameterValue(Y) = flxImport.TextMatrix(i, 3)
    End If
Next
For i = LBound(tmp) To UBound(tmp)
    WriteDataTableFiles tmp(i)
Next
End Sub

Private Sub ClearForm()
ClearTable
Me.Caption = Me.Tag
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.TextMatrix(0, 0) = "ID Number"
flxImport.TextMatrix(0, 1) = "Excel Path"
flxImport.TextMatrix(0, 2) = "Parameter Name"
flxImport.TextMatrix(0, 3) = "Parameter Value"
flxImport.Rows = 2
End Sub

Private Sub OutputTable()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim i, Protections
Dim curTab
Dim w, j


Set xlObject = New Excel.Application

On Error Resume Next
For Each w In xlObject.Workbooks
   w.Close savechanges:=False
Next w
On Error GoTo 0

On Error GoTo OutErr: Set xlWB = xlObject.Workbooks.Add
    xlObject.Sheets("Sheet2").Range("A1").Value = "1 - Only edit values in the column(s) colored green"
    xlObject.Sheets("Sheet2").Range("A2").Value = "2 - Do not Add, Delete or Modify Rows and Column's Position, Color or Order"
    xlObject.Sheets("Sheet2").Range("A3").Value = "3 - The same sheet will be uploaded using SuperQC tools"
    xlObject.Sheets("Sheet2").Range("A4").Value = flxImport.TextMatrix(0, 0) & " First Entry:"
    xlObject.Sheets("Sheet2").Range("A5").Value = flxImport.TextMatrix(0, 0) & " Last Entry:"
    xlObject.Sheets("Sheet2").Range("A6").Value = "Total Records:"
    xlObject.Sheets("Sheet2").Range("B4").Value = flxImport.TextMatrix(1, 0)
    xlObject.Sheets("Sheet2").Range("B5").Value = flxImport.TextMatrix(flxImport.Rows - 1, 0)
    xlObject.Sheets("Sheet2").Range("B6").Value = flxImport.Rows - 1
    
    xlObject.Sheets("Sheet2").Range("A7").Value = "Environment:"
    xlObject.Sheets("Sheet2").Range("B7").Value = curDomain & "-" & curProject
    
    xlObject.Sheets("Sheet2").Columns("A:A").EntireColumn.AutoFit
    With xlObject.Sheets("Sheet2").Columns("A:A").Font
        .Name = "Arial"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
  xlObject.Sheets("Sheet2").Name = "INSTRUCTIONS"
  'xlObject.Visible = True
  curTab = "DT_MAN01"
  xlObject.Sheets("Sheet1").Name = curTab
  flxImport.FixedCols = 0
  flxImport.FixedRows = 0
  flxImport.col = 0
  flxImport.row = 0
  Pause 1
  flxImport.RowSel = flxImport.Rows - 1
  flxImport.ColSel = flxImport.Cols - 1
    For i = 0 To flxImport.Rows - 1
            For j = 0 To flxImport.Cols - 1
                xlObject.Sheets(curTab).Range(Chr(j + 64 + 1) & i + 1).Formula = "'" & flxImport.TextMatrix(i, j)
            Next
    Next
  flxImport.FixedCols = 1
  flxImport.FixedRows = 1
'On Error Resume Next
    xlObject.Sheets(curTab).Range("A:D").Select
        
    xlObject.Sheets(curTab).Range("A:D").Borders(xlDiagonalDown).LineStyle = xlNone
    xlObject.Sheets(curTab).Range("A:D").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlObject.Sheets(curTab).Range("A:D").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:D").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:D").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:D").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:D").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:D").Borders(xlInsideHorizontal)
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
    xlObject.Sheets(curTab).Range("A:D").Select
    xlObject.Sheets(curTab).Range("A:D").EntireColumn.AutoFit
    xlObject.Sheets(curTab).Range("A1").Select
    
    xlObject.Sheets(curTab).Range("A1").AddComment
    xlObject.Sheets(curTab).Range("A1").Comment.Visible = False
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & mdiMain.Caption & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""
    
    xlObject.Sheets(curTab).Range("D:D").Interior.ColorIndex = 35
    xlObject.Sheets(curTab).Protection.AllowEditRanges.Add Title:="Range1", Range:=xlObject.Sheets(curTab).Range("D:D")
    xlObject.Sheets(curTab).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  
  xlObject.Workbooks(1).SaveAs "DT_MAN01-" & "" & "-" & Format(Now, "mmddyyyy HHMM AMPM")
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate
  FXGirl.EZPlay FXExportToExcel
  Set xlWB = Nothing
  Set xlObject = Nothing
  
  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export to MS Excel completed.": Exit Sub:
OutErr:     MsgBox Err.Description, vbCritical: xlObject.Visible = True: xlObject.ActiveWindow.Activate: Set xlWB = Nothing: Set xlObject = Nothing
On Error GoTo 0
End Sub

Private Function IncorrectHeaderDetails() As Boolean
    If flxImport.TextMatrix(0, 0) <> "ID Number" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 1) <> "Excel Path" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 2) <> "Parameter Name" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 3) <> "Parameter Value" Then IncorrectHeaderDetails = True
End Function

Private Sub dirFolders_Change()
On Error Resume Next
fileExcel.path = dirFolders.path
stsBar.Panels(2).Text = fileExcel.ListCount & " excel files found (sub folders not included)."
If Err.Number <> 0 Then
    MsgBox Err.Description
    dirFolders.ListIndex = 0
End If
End Sub

Private Sub drvDrive_Change()
On Error Resume Next
dirFolders.path = drvDrive.List(drvDrive.ListIndex)
If Err.Number <> 0 Then
    MsgBox Err.Description
    drvDrive.ListIndex = 0
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
dirFolders.height = stsBar.Top - 1150
flxImport.height = stsBar.Top - flxImport.Top - 225
flxImport.width = Me.width - flxImport.Left - 350
End Sub

Private Sub ReadDataTableFiles(ExcelPath As String)
On Error GoTo errHandler:
    Dim xlsApp As Object
    Dim xlsWB1 As Object
    Dim xlsWS1 As Object
    'Opening the file to parse now
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.Visible = False
    xlsApp.DisplayAlerts = False
    Set xlsWB1 = xlsApp.Workbooks.Open(ExcelPath)
    Set xlsWS1 = xlsWB1.Worksheets(1)
    Dim col As Integer
    Dim row As Integer
    Dim str As String
    Dim i As Integer, MaxRow, MaxCol, CaseArray
    
    str = ""
    MaxRow = 1
    MaxCol = 255
    For i = 1 To 255
        If Trim(xlsWS1.Cells(1, i).Value) = "" Then
            MaxCol = i - 2
            Exit For
        End If
    Next
    ReDim CaseArray(MaxRow, MaxCol)
    For row = 0 To MaxRow
        For col = 0 To MaxCol
            If Left(xlsWS1.Cells(row + 1, col + 1).Formula, 1) = "=" Then
                CaseArray(row, col) = CStr("'" & xlsWS1.Cells(row + 1, col + 1).Formula)
            Else
                CaseArray(row, col) = CStr(xlsWS1.Cells(row + 1, col + 1).Formula)
            End If
        Next
    Next
    xlsWB1.Close
    xlsApp.Quit
    Set xlsApp = Nothing
    Set xlsWB1 = Nothing
    Set xlsWS1 = Nothing
    If flxImport.TextMatrix(flxImport.Rows - 1, 0) <> "" Then
        flxImport.Rows = flxImport.Rows + 1
    End If
    For i = 0 To MaxCol
        flxImport.TextMatrix(flxImport.Rows - 1, 0) = flxImport.Rows - 1
        flxImport.TextMatrix(flxImport.Rows - 1, 1) = ExcelPath
        flxImport.TextMatrix(flxImport.Rows - 1, 2) = CaseArray(0, i)
        flxImport.TextMatrix(flxImport.Rows - 1, 3) = CaseArray(1, i)
        flxImport.Rows = flxImport.Rows + 1
    Next
    flxImport.Rows = flxImport.Rows - 1
    Exit Sub
errHandler:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub WriteDataTableFiles(ExcelValues As DataTableEntries)
On Error GoTo errHandler:
    Dim xlsApp As Object
    Dim xlsWB1 As Object
    Dim xlsWS1 As Object
    'Opening the file to parse now
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.Visible = False
    xlsApp.DisplayAlerts = False
    Set xlsWB1 = xlsApp.Workbooks.Open(ExcelValues.DataTablePath)
    Set xlsWS1 = xlsWB1.Worksheets(1)
    Dim col As Integer
    Dim row As Integer
    Dim str As String
    Dim i As Integer, j, MaxRow, MaxCol
    str = ""
    MaxRow = 1
    MaxCol = 255
    For i = 1 To 255
        If Trim(xlsWS1.Cells(1, i).Formula) = "" Then
            MaxCol = i
            Exit For
        End If
    Next
    For i = 1 To MaxCol
        For j = 0 To UBound(ExcelValues.ParameterName)
            If UCase(Trim(xlsWS1.Cells(1, i).Formula)) = Trim(UCase(ExcelValues.ParameterName(j))) Then
                If Left(ExcelValues.ParameterValue(j), 2) = "'=" Then
                    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading DataTable Records " & Now & " " & xlsWS1.Cells(1, i).Formula & " (" & xlsWS1.Cells(2, i).Formula & " - " & ExcelValues.ParameterValue(j) & ")"
                    xlsWS1.Cells(2, i).Formula = Right(ExcelValues.ParameterValue(j), Len(ExcelValues.ParameterValue(j)) - 1)
                Else
                    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading DataTable Records " & Now & " " & xlsWS1.Cells(1, i).Formula & " (" & xlsWS1.Cells(2, i).Formula & " - " & ExcelValues.ParameterValue(j) & ")"
                    xlsWS1.Cells(2, i).Formula = "'" & CStr(ExcelValues.ParameterValue(j))
                End If
                Exit For
            End If
        Next
    Next
    xlsApp.DisplayAlerts = False
    xlsWB1.Save
    xlsWB1.Close
    xlsApp.Quit
    Set xlsApp = Nothing
    Set xlsWB1 = Nothing
    Set xlsWS1 = Nothing
    Exit Sub
errHandler:
On Error GoTo 0
On Error Resume Next
    xlsApp.DisplayAlerts = False
    xlsWB1.Close
    xlsApp.Quit
    Set xlsApp = Nothing
    Set xlsWB1 = Nothing
    Set xlsWS1 = Nothing
End Sub


Sub DirWalk(ByVal sPattern As String, ByVal CurrDir As String, sFound() As String)
Dim i As Integer
Dim sCurrPath As String
Dim sFile As String
Dim ii As Integer
Dim iFiles As Integer
Dim iLen As Integer

If Right$(CurrDir, 1) <> "\" Then
    dirFolders.path = CurrDir & "\"
Else
    dirFolders.path = CurrDir
End If
For i = 0 To dirFolders.ListCount
    If dirFolders.List(i) <> "" Then
        DoEvents
        Call DirWalk(sPattern, dirFolders.List(i), sFound)
    Else
        If Right$(dirFolders.path, 1) = "\" Then
            sCurrPath = Left(dirFolders.path, Len(dirFolders.path) - 1)
        Else
            sCurrPath = dirFolders.path
        End If
        fileExcel.path = sCurrPath
        fileExcel.Pattern = sPattern
        If fileExcel.ListCount > 0 Then 'matching files found in the Directory
            For ii = 0 To fileExcel.ListCount - 1
                ReDim Preserve sFound(UBound(sFound) + 1)
                sFound(UBound(sFound) - 1) = sCurrPath & "\" & fileExcel.List(ii)
            Next ii
        End If
        iLen = Len(dirFolders.path)
        Do While Mid(dirFolders.path, iLen, 1) <> "\"
            iLen = iLen - 1
        Loop
        dirFolders.path = Mid(dirFolders.path, 1, iLen)
    End If
Next i
End Sub
