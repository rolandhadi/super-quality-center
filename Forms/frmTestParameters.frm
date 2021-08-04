VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTestParameters 
   Caption         =   "Test Plan Parameters Module"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   Icon            =   "frmTestParameters.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   10125
   Tag             =   "Test Parameters Module"
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
      Picture         =   "frmTestParameters.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Import step description and expected results from an excel file"
      Top             =   1800
      Width           =   2205
   End
   Begin VB.ListBox lstUpdateFields 
      Columns         =   4
      Height          =   960
      ItemData        =   "frmTestParameters.frx":1070
      Left            =   4560
      List            =   "frmTestParameters.frx":1072
      Style           =   1  'Checkbox
      TabIndex        =   2
      ToolTipText     =   "Select Values to update"
      Top             =   780
      Width           =   7995
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10125
      _ExtentX        =   17859
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
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdGenerate"
            Object.ToolTipText     =   "Generate"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdOutput"
            Object.ToolTipText     =   "Export to Excel"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdUpload"
            Object.ToolTipText     =   "Upload to HPQC"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin VB.CheckBox chkCSV 
         Caption         =   "Download to CSV"
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   120
         Width           =   2655
      End
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
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestParameters.frx":1074
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestParameters.frx":1306
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestParameters.frx":1598
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
      Checkboxes      =   -1  'True
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
      Height          =   3735
      Left            =   4560
      TabIndex        =   3
      Top             =   2220
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
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
            Picture         =   "frmTestParameters.frx":1826
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestParameters.frx":1F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestParameters.frx":264A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestParameters.frx":2D5C
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
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   6060
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmTestParameters.frx":346E
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList_Sts 
      Left            =   11940
      Top             =   0
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
            Picture         =   "frmTestParameters.frx":39BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestParameters.frx":3CA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestParameters.frx":41F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestParameters.frx":4743
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Fields to Update"
      Height          =   195
      Left            =   4560
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmTestParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GenerateOutput()
Dim rs As TDAPIOLELib.Recordset
Dim AllScript
Dim objCommand
Dim i, j
Dim strPath
Dim iterd
Dim NewVal
Dim iter
Dim TimeSt
Dim AllF
Dim Decompile
Dim P
Dim k, a0, a1, a2, a3, a4, a5, last_a1
    
    TimeSt = Format(Now, "mmm-dd-yyyy hhmmss") & "-"
    If chkCSV.Value = Checked Then
      AllF = InputBox("Enter file name", "File name", "[Test Parameters] ")
    Else
      AllF = "[Test Parameters] "
    End If
    
    ReDim CheckedItems(0): strPath = ""
    GetAllCheckedItems QCTree.Nodes(1)
    For j = LBound(CheckedItems) To UBound(CheckedItems) - 1
        If Left(CheckedItems(j), 1) = "F" Then
            strPath = strPath & "AL_ABSOLUTE_PATH LIKE '" & GetFromTable(Right(CheckedItems(j), Len(CheckedItems(j)) - 1), "AL_ITEM_ID", "AL_ABSOLUTE_PATH", "ALL_LISTS") & "%'" & " OR "
        ElseIf Left(CheckedItems(j), 1) = "C" Then
            strPath = strPath & "TS_TEST_ID = " & Right(CheckedItems(j), Len(CheckedItems(j)) - 1) & " OR "
        End If
    Next
    If Trim(strPath) <> "" Then
        strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
    Else
        MsgBox "Please select and check source(s) in the HPQC folder tree"
        stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
        Exit Sub
    End If
    
    Set objCommand = QCConnection.Command
    
    objCommand.CommandText = "SELECT BPIP_ID, TS_SUBJECT, TS_NAME, CO_NAME, FP_NAME, BPIP_VALUE FROM  BP_ITER_PARAM, BP_PARAM, FRAMEWORK_PARAM, COMPONENT, TEST, BPTEST_TO_COMPONENTS, ALL_LISTS WHERE  BPIP_BPP_ID = BPP_ID AND    BPP_PARAM_ID = FP_ID AND    FP_COMPONENT_ID = CO_ID AND    TS_TEST_ID = BC_BPT_ID AND    BC_CO_ID = CO_ID AND    BC_ID = BPP_BPC_ID AND    TS_SUBJECT = AL_ITEM_ID AND " & strPath & " ORDER BY TS_SUBJECT, BC_BPT_ID, BC_ORDER, FP_ORDER"
    Debug.Print Me.Caption & "-" & objCommand.CommandText
    
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[SQL] " & Now & " " & objCommand.CommandText
    Set rs = objCommand.Execute                                                                                                                                                                                                                                                 'HERE!!!!!! <<<-------------
    If rs.RecordCount > 10000 And chkCSV.Value = Unchecked Then '***
        MsgBox "The records found exceeds 2500 records. It will be automatically generated as a CSV file.", vbOKOnly
        chkCSV.Value = Checked
        Exit Sub
        GenerateOutput
    End If '***
    AllScript = """" & "BPIP ID" & """" & "," & """" & "Test Folder" & """" & "," & """" & "Test Name" & """" & "," & """" & "Component Name" & """" & "," & """" & "Parameter Name" & """" & "," & """" & "Parameter Value" & """"
    a1 = ""
    last_a1 = " "
    ClearTable
    If chkCSV.Value = Unchecked Then '***
        flxImport.Rows = rs.RecordCount + 1
    End If '***
    k = 0
    mdiMain.pBar.Max = rs.RecordCount + 3
    For i = 1 To rs.RecordCount
        'If i = 1 Then FileWrite App.path & "\SQC DAT" & "\" & QCTree.SelectedItem.Key & "-" & TimeSt & AllF & ".xls", AllScript
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Processing " & i & " out of " & rs.RecordCount
                If chkCSV.Value = Unchecked Then
                  k = k + 1
                  flxImport.Rows = k + 1
                  flxImport.TextMatrix(k, 0) = rs.FieldValue("BPIP_ID")
                  If rs.FieldValue("BPIP_ID") <> flxImport.TextMatrix(k - 1, 0) Then
                      flxImport.TextMatrix(k, 1) = GetTestFolderPath(rs.FieldValue("TS_SUBJECT"))  'rs.FieldValue("TS_SUBJECT") '"Test Folder"
                  Else
                      flxImport.TextMatrix(k, 1) = flxImport.TextMatrix(k - 1, 1)
                  End If
                  flxImport.TextMatrix(k, 2) = rs.FieldValue("TS_NAME")
                  flxImport.TextMatrix(k, 3) = rs.FieldValue("CO_NAME")
                  flxImport.TextMatrix(k, 4) = rs.FieldValue("FP_NAME")
                  flxImport.TextMatrix(k, 5) = rs.FieldValue("BPIP_VALUE")
                Else
                  a0 = rs.FieldValue("BPIP_ID")
                  a2 = rs.FieldValue("TS_NAME")
                  a3 = rs.FieldValue("CO_NAME")
                  a4 = rs.FieldValue("FP_NAME")
                  a5 = rs.FieldValue("BPIP_VALUE")
                  If a1 <> last_a1 Then
                  a1 = GetTestFolderPath(rs.FieldValue("TS_SUBJECT"))
                  last_a1 = a1
                  Else
                  a1 = last_a1
                  End If
                  If Trim(AllScript) <> "" Then
                        AllScript = AllScript & vbCrLf & """" & a0 & """" & "," & """" & a1 & """" & "," & """" & a2 & """" & "," & """" & a3 & """" & "," & """" & a4 & """" & "," & """" & a5 & """"
                  Else
                        AllScript = AllScript & """" & a0 & """" & "," & """" & a1 & """" & "," & """" & a2 & """" & "," & """" & a3 & """" & "," & """" & a4 & """" & "," & """" & a5 & """"
                  End If
                End If
            If chkCSV.Value = Checked Then '***
              If i Mod 500 = 0 Then
                  FileAppend App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv", AllScript
                  AllScript = ""
              End If
            End If '***
            mdiMain.pBar.Value = k + 1
            If mdiMain.pBar.Max > 300 Then
                Select Case GlobalStrings.Percentage(mdiMain.pBar.Value, mdiMain.pBar.Max)
                Case 25 To 25.3
                    FXGirl.EZPlay FX25
                Case 50 To 50.3
                    FXGirl.EZPlay FX50
                Case 75 To 75.3
                    FXGirl.EZPlay FX75
                End Select
            End If
            rs.Next
    Next
    FXGirl.EZPlay FXSQCExtractCompleted
    mdiMain.pBar.Value = mdiMain.pBar.Max
If chkCSV.Value = Checked Then '***
    FileAppend App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv", AllScript: If MsgBox("Successfully exported to " & App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv" & vbCrLf & "Do you want to launch the extracted file?", vbYesNo) = vbYes Then Shell "explorer.exe " & App.path & "\SQC Logs" & "\", vbNormalFocus
    AllScript = vbCrLf & " ,"
    AllScript = AllScript & """" & "SQL Code:" & """" & "," & """" & Replace(objCommand.CommandText, """", "'") & """"
    FileAppend App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv", AllScript
End If '***
stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = mdiMain.pBar.Max - 1 & " record(s) generated"
End Sub

Function GetCommentText(rCommentCell As Range)
     Dim strGotIt As String
         On Error Resume Next
         strGotIt = WorksheetFunction.Clean _
             (rCommentCell.Comment.Text)
         GetCommentText = strGotIt
         On Error GoTo 0
End Function

Private Sub chkCSV_Click()
If chkCSV.Value = Checked Then
    If MsgBox("Are you sure you want to download directly to CSV?", vbYesNo) = vbYes Then
        chkCSV.Value = Checked
    Else
        chkCSV.Value = Unchecked
    End If
End If
End Sub

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
    
    For i = 0 To lstUpdateFields.ListCount - 1
        lstUpdateFields.Selected(i) = False
    Next
    
    With xlObject.ActiveWorkbook.ActiveSheet
         If UCase(Trim(curDomain & "-" & curProject)) <> UCase(Trim(xlObject.ActiveWorkbook.Sheets(2).Range("B7").Value)) Then
            MsgBox "The spreadsheet is from a different Domain or Project"
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
         If InStr(1, GetCommentText(.Range("A1")), "Test Parameters Module") = 0 Then
            MsgBox "Import file is invalid. Please use only sheets generated by the SuperQC"
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
         For i = 1 To 6
            If .Range(ColumnLetter(CInt(i)) & 1).Interior.ColorIndex = 35 Then
                For j = 0 To lstUpdateFields.ListCount - 1
                    If .Range(ColumnLetter(CInt(i)) & 1).Value = lstUpdateFields.List(j) Then
                        lstUpdateFields.Selected(j) = True
                    End If
                Next
            End If
         Next
         lastrow = .Range("A" & .Rows.Count).End(xlUp).row
        .Range("A1:Y" & lastrow).COPY 'Set selection to Copy
    End With
       
    With flxImport
        .Clear
        .Redraw = False     'Dont draw until the end, so we avoid that flash
        .row = 0            'Paste from first cell
        .col = 0
        .Rows = lastrow
        .Cols = 6
        .RowSel = lastrow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
        .ColSel = 6 - 1
    End With
    
     With flxImport
        .Clear
        .Redraw = False     'Dont draw until the end, so we avoid that flash
        .row = 0            'Paste from first cell
        .col = 0
        .Rows = lastrow
        .Cols = 6
        .RowSel = lastrow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
        .ColSel = 6 - 1
        .Clip = Replace(Clipboard.GetText, vbNewLine, vbCr)   'Replace carriage return with the correct one
        .col = 1            'Just to remove that blue selection from Flexgrid
        .Redraw = True      'Now draw
    End With
        
    xlObject.DisplayAlerts = False 'To avoid "Save woorkbook" messagebox
    
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = flxImport.Rows - 1 & " record(s) loaded"
    
    'Close Excel
    xlWB.Close
    xlObject.Application.Quit
    Set xlWB = Nothing
    Set xlObject = Nothing
    mdiMain.pBar.Max = 100
    mdiMain.pBar.Value = 100
Exit Sub
ErrLoad:
MsgBox "There was an error while importing the file. Please refresh and close all excel and try again" & vbCrLf & Err.Description, vbCritical
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
On Error Resume Next
QCTree.height = stsBar.Top - 550
lstUpdateFields.width = Me.width - lstUpdateFields.Left - 350
flxImport.height = stsBar.Top - flxImport.Top - 250
flxImport.width = Me.width - flxImport.Left - 350
End Sub

Private Sub Label1_Click()
Dim tmpPath, tmpID: On Error Resume Next
tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "AL_ITEM_ID", "AL_ABSOLUTE_PATH", "ALL_LISTS") & "%"
frmLogs.txtLogs.Text = "Test ID: " & tmpID & vbCrLf & "AL_ABSOLUTE_PATH: " & tmpPath & vbCrLf & "Folder Path: " & QCTree.SelectedItem.FullPath
frmLogs.Show 1
End Sub

Private Sub QCTree_DblClick()
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
Dim nodx As Node

    If QCTree.SelectedItem.Children <> 0 Then Exit Sub
    
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT AL_ITEM_ID, AL_DESCRIPTION FROM ALL_LISTS WHERE AL_FATHER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY AL_DESCRIPTION"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("F" & rs.FieldValue("AL_ITEM_ID")), rs.FieldValue("AL_DESCRIPTION"), 1
        rs.Next
    Next
        
    If Left(QCTree.SelectedItem.Key, 1) = "F" Then
        Set objCommand = QCConnection.Command
        objCommand.CommandText = "SELECT DISTINCT TS_NAME, TS_TEST_ID FROM TEST, ALL_LISTS WHERE TS_SUBJECT = AL_ITEM_ID AND AL_ITEM_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY TS_NAME"
        Set rs = objCommand.Execute
        For i = 1 To rs.RecordCount
            QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("C" & rs.FieldValue("TS_TEST_ID")), rs.FieldValue("TS_NAME"), 2
            rs.Next
        Next
    End If

End Sub

Private Sub QCTree_NodeCheck(ByVal Node As MSComctlLib.Node)
Node.Selected = True
End Sub

Private Sub stsBar_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
frmLogs.txtLogs.Text = stsBar.Panels(2).Text: frmLogs.Show 1
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
    GenerateOutput
    Exit Sub
OutputErr:
    MsgBox "Data was been truncated because of an error." & vbCrLf & Err.Description
Case "cmdOutput"
    If flxImport.Rows <= 1 Then
        MsgBox "Nothing to output", vbInformation
    Else
        If GetEditableFields() = "IV:IV" Then
            If MsgBox("You have selected nothing to update on this sheet. The whole sheet will be read-only. Do you want to proceed?", vbYesNo) = vbYes Then
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
                OutputTable
            End If
        Else
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
            OutputTable
        End If
    End If
Case "cmdUpload"
If IncorrectHeaderDetails = False Then
    If GetEditableFields <> "IV:IV" Then
        If MsgBox("Are you sure you want to mass update " & flxImport.Rows - 1 & " record(s) to the Test Plan Parameters?", vbYesNo) = vbYes Then
            Randomize: tmpR = CInt(Rnd(1000) * 10000)
            If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
                If lstUpdateFields.Selected(0) = True Then Upload_Test_Parameters
            Else
                MsgBox "Invalid pass key", vbCritical
            End If
        End If
    Else
        MsgBox "No fields to update", vbCritical
    End If
Else
    MsgBox "Invalid or Incorrect upload sheet file selected", vbCritical
End If
End Select
End Sub

Private Sub ClearForm()
ClearTable
QCTree.Nodes.Clear

Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
    QCTree.Nodes.Add , , "Root", "Subject", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT AL_ITEM_ID, AL_DESCRIPTION FROM ALL_LISTS WHERE AL_FATHER_ID = 2 ORDER BY AL_DESCRIPTION"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("AL_ITEM_ID")), rs.FieldValue("AL_DESCRIPTION"), 1
        rs.Next
    Next
    
    lstUpdateFields.Clear
    lstUpdateFields.AddItem "Parameter Value"
    Me.Caption = Me.Tag
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.TextMatrix(0, 0) = "BPIP ID"
flxImport.TextMatrix(0, 1) = "Test Folder"
flxImport.TextMatrix(0, 2) = "Test Name"
flxImport.TextMatrix(0, 3) = "Component Name"
flxImport.TextMatrix(0, 4) = "Parameter Name"
flxImport.TextMatrix(0, 5) = "Parameter Value"
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
  curTab = "TESTPARAM01"
  xlObject.Sheets("Sheet1").Name = curTab
  flxImport.FixedCols = 0
  flxImport.FixedRows = 0
  flxImport.col = 0
  flxImport.row = 0
  Pause 1
  flxImport.RowSel = flxImport.Rows - 1
  flxImport.ColSel = flxImport.Cols - 1
  Clipboard.Clear
  Clipboard.SetText flxImport.Clip
'  flxImport.Clip = Replace(flxImport.Clip, vbCrLf, "")
'  flxImport.Clip = Replace(flxImport.Clip, Chr(10) + Chr(13), "")
'  flxImport.Clip = Replace(flxImport.Clip, Chr(10), "")
'  flxImport.Clip = Replace(flxImport.Clip, Chr(13), "")
  flxImport.FixedCols = 1
  flxImport.FixedRows = 1
  
  xlObject.Sheets(curTab).Range("A1").Select
  xlObject.Sheets(curTab).Paste

'On Error Resume Next
    xlObject.Sheets(curTab).Range("A:F").Select
        
    xlObject.Sheets(curTab).Range("A:F").Borders(xlDiagonalDown).LineStyle = xlNone
    xlObject.Sheets(curTab).Range("A:F").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlObject.Sheets(curTab).Range("A:F").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:F").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:F").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:F").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:F").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:F").Borders(xlInsideHorizontal)
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
    xlObject.Sheets(curTab).Range("A:F").Select
    xlObject.Sheets(curTab).Range("A:F").EntireColumn.AutoFit
    xlObject.Sheets(curTab).Range("A1").Select
    
    xlObject.Sheets(curTab).Range("A1").AddComment
    xlObject.Sheets(curTab).Range("A1").Comment.Visible = False
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & mdiMain.Caption & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""
    
    xlObject.Sheets(curTab).Range(GetEditableFields).Interior.ColorIndex = 35
    xlObject.Sheets(curTab).Range("A:B").Interior.ColorIndex = 3
  
  xlObject.Sheets(curTab).Protection.AllowEditRanges.Add Title:="Range1", Range:=xlObject.Sheets(curTab).Range(GetEditableFields)
  xlObject.Sheets(curTab).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  xlObject.Workbooks(1).SaveAs "TESTPARAM01-" & CleanTheString(QCTree.SelectedItem.Text) & "-" & Format(Now, "mmddyyyy HHMM AMPM")
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate
  
  Set xlWB = Nothing
  Set xlObject = Nothing
  FXGirl.EZPlay FXExportToExcel
  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export to MS Excel completed.": Exit Sub:
OutErr:     MsgBox Err.Description, vbCritical: xlObject.Visible = True: xlObject.ActiveWindow.Activate: Set xlWB = Nothing: Set xlObject = Nothing
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
             tmp = tmp & ColumnLetter(CInt(j + 1)) & ":" & ColumnLetter(CInt(j + 1)) & ", "
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

Private Function Upload_Test_Parameters()
Dim lastrowofsheet As Integer, HasUploadIssue
Dim i, X, j, objCommand, rs
    lastrowofsheet = flxImport.Rows
    mdiMain.pBar.Max = lastrowofsheet + 2
    For i = 1 To lastrowofsheet - 1
        On Error Resume Next
        Set objCommand = QCConnection.Command
        objCommand.CommandText = "UPDATE BP_ITER_PARAM SET BPIP_VALUE = '" & flxImport.TextMatrix(i, 5) & "' WHERE BPIP_ID = " & flxImport.TextMatrix(i, 0)
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[SQL] " & Now & " " & objCommand.CommandText
        Set rs = objCommand.Execute
        If Err.Number <> 0 Then
            FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[UPDATE TEST PARAM: (FAILED) " & Now & " " & flxImport.TextMatrix(i, 0) & "-" & flxImport.TextMatrix(i, 5) & "] " & Err.Description
            HasUploadIssue = HasUploadIssue + 1
        Else
            FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[UPDATE TEST PARAM: (PASSED) " & Now & " " & flxImport.TextMatrix(i, 0) & "-" & flxImport.TextMatrix(i, 5) & "]"
        End If
        stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Loading Test Parameters - " & i + 1 & " out of " & lastrowofsheet + 1 & " (" & flxImport.TextMatrix(i, 0) & ")"
        Err.Clear
        On Error GoTo 0
        mdiMain.pBar.Value = i + 1
        If mdiMain.pBar.Max > 300 Then
            Select Case GlobalStrings.Percentage(mdiMain.pBar.Value, mdiMain.pBar.Max)
            Case 25 To 25.3
                FXGirl.EZPlay FX25
            Case 50 To 50.3
                FXGirl.EZPlay FX50
            Case 75 To 75.3
                FXGirl.EZPlay FX75
            End Select
        End If
    Next
    mdiMain.pBar.Value = mdiMain.pBar.Max
    FXGirl.EZPlay FXDataUploadCompleted
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = i - 1 & " of " & lastrowofsheet - 1 & " Parameter(s) in Test Plan updated"
    QCConnection.SendMail "user@companyemail.com", "", "[HPQC UPDATES] Parameter(s) in Test Plan updated by " & curUser & " in " & curDomain & "-" & curProject, i - 1 & " of " & lastrowofsheet - 1 & " Parameters) in Test Plan updated" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
    QCConnection.SendMail curUser, "", "[HPQC UPDATES] Parameter(s) in Test Plan updated by " & curUser & " in " & curDomain & "-" & curProject, i - 1 & " of " & lastrowofsheet - 1 & " Parameter(s) in Test Plan updated" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
End Function

Private Function GetFieldNameinDB(X As String)
Select Case X
    Case "BPIP ID"
        GetFieldNameinDB = "BPIP_ID"
    Case "Test Folder"
        GetFieldNameinDB = "TS_SUBJECT"
    Case "Test Name"
        GetFieldNameinDB = "TS_NAME"
    Case "Component Name"
        GetFieldNameinDB = "CO_NAME"
    Case "Parameter Name"
        GetFieldNameinDB = "FP_NAME"
    Case "Parameter Value"
        GetFieldNameinDB = "BPIP_VALUE"
End Select
End Function

Private Function GetFieldNumberinDB(X As Integer)
Select Case X
    Case 1
        GetFieldNumberinDB = "BPIP_ID"
    Case 2
        GetFieldNumberinDB = "TS_SUBJECT"
    Case 3
        GetFieldNumberinDB = "TS_NAME"
    Case 4
        GetFieldNumberinDB = "CO_NAME"
    Case 5
        GetFieldNumberinDB = "FP_NAME"
    Case 6
        GetFieldNumberinDB = "BPIP_VALUE"
End Select
End Function

Private Function GetFieldNumberinDB_byName(X As String)
Select Case X
    Case "BPIP ID"
        GetFieldNumberinDB_byName = 0
    Case "Test Folder"
        GetFieldNumberinDB_byName = 1
    Case "Test Name"
        GetFieldNumberinDB_byName = 2
    Case "Component Name"
        GetFieldNumberinDB_byName = 3
    Case "Parameter Name"
        GetFieldNumberinDB_byName = 4
    Case "Parameter Value"
        GetFieldNumberinDB_byName = 5
        
    Case "BPIP_ID"
        GetFieldNumberinDB_byName = 0
    Case "TS_SUBJECT"
        GetFieldNumberinDB_byName = 1
    Case "TS_NAME"
        GetFieldNumberinDB_byName = 2
    Case "CO_NAME"
        GetFieldNumberinDB_byName = 3
    Case "FP_NAME"
        GetFieldNumberinDB_byName = 4
    Case "BPIP_VALUE"
        GetFieldNumberinDB_byName = 5
End Select
End Function

Private Function IncorrectHeaderDetails() As Boolean
    If flxImport.TextMatrix(0, 0) <> "BPIP ID" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 1) <> "Test Folder" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 2) <> "Test Name" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 3) <> "Component Name" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 4) <> "Parameter Name" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 5) <> "Parameter Value" Then IncorrectHeaderDetails = True
End Function

Private Function GetTestFolderPath(strID As String) As String
Dim Fact As TreeManager
Dim Obj As SubjectNode
If Trim(strID) = "" Then Exit Function
Set Fact = QCConnection.TreeManager
Set Obj = Fact.NodeById(strID)
GetTestFolderPath = Obj.path
End Function

'########################### Promote Test Plan Parameters ###########################
Private Function PromoteParamBPTTest(Test_Plan_ID As String, BC_ID As String)
'On Error GoTo Exp
Dim comp As Component
Dim compFact As ComponentFactory
Dim targetFolderID
Dim myBPComponent As BPComponent
Dim myBPTest As BusinessProcess
Dim tfact As TestFactory
Dim mytest As Test
Dim bpcount As Long
Dim com As Command
Dim recset As Recordset
Dim mycurrentTestID As String

Dim tstFactory As TestFactory
Dim tst As Test
Dim fl As IBusinessProcess2
Dim bpComp As BPComponent
Dim bpParam As BPParameter
Dim newFlowOutputParam As ComponentParam
Dim iter As BPIteration
Dim myTempIteration As BPIteration
Dim myTempIterationParam As BPIterationParam
Dim myBPGroup As BPGroup

Dim tsfilter As TDFilter
Dim tList As List
Dim comfol As ComponentFolderFactory
Dim X, myRTParam As RTParam

        Set tfact = QCConnection.TestFactory
        Set mytest = tfact.Item(Test_Plan_ID)
        Set myBPTest = mytest
        myBPTest.Load
        For Each bpComp In myBPTest.BPComponents
            If bpComp.Component.ID = BC_ID Then
                For Each iter In bpComp.Iterations
                    X = 1
                    For Each bpParam In bpComp.BPParams
                            For Each myTempIteration In bpComp.Iterations
                                On Error Resume Next
                                Set myTempIterationParam = myTempIteration.IterationParams.Item(X)
                                If Err.Number = 0 Then
                                    If myTempIterationParam.Value = "" Then
                                        Set myRTParam = myBPTest.AddRTParam
                                        myRTParam.Name = myTempIterationParam.BPParameter.ComponentParamName
                                        myRTParam.ValueType = "String"
                                        myTempIterationParam.Value = "{" & myRTParam.Name & "}"
                                    End If
                                     FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Promoting Test Parameters (PASSED) " & Now & " (TEST ID:" & Test_Plan_ID & "-BC ID:" & BC_ID & ") - " & myTempIterationParam.Value
                                Else
                                     FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Promoting Test Parameters (FAILED) " & Now & " (TEST ID:" & Test_Plan_ID & "-BC ID:" & BC_ID & ") - " & myTempIterationParam.Value & " " & Err.Description
                                End If
                                Err.Clear
                                On Error GoTo 0
                            Next myTempIteration
                            X = X + 1
                    Next bpParam
                    myBPTest.Refresh
                    myBPTest.Save
                    X = 0
                Next iter
            End If
        Next bpComp

        Set comp = Nothing
        Set mytest = Nothing
        Set myBPTest = Nothing
        Set myBPComponent = Nothing
        Set myTempIterationParam = Nothing
        Set myRTParam = Nothing

PromoteParamBPTTest = True

Set comp = Nothing
Set mytest = Nothing
Set myBPTest = Nothing
Set myBPComponent = Nothing
Set myTempIterationParam = Nothing
Set myRTParam = Nothing
End Function
'########################### End Of Promote Test Plan Parameters ###########################
