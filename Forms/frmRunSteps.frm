VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRunSteps 
   Caption         =   "Run Steps Module"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12675
   Icon            =   "frmRunSteps.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   12675
   Tag             =   "Run Steps Module"
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkAutoDates 
      Caption         =   "Auto Compute Dates"
      Height          =   315
      Left            =   8100
      TabIndex        =   10
      Top             =   1860
      Width           =   1875
   End
   Begin VB.CheckBox chkInvertDate 
      Caption         =   "Invert Date"
      Height          =   315
      Left            =   6840
      TabIndex        =   9
      Top             =   1860
      Width           =   1155
   End
   Begin VB.CheckBox chkUpload 
      Caption         =   "Upload from XML file"
      Height          =   315
      Left            =   4080
      TabIndex        =   8
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtFilter 
      Height          =   375
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Filter by Test Instance (L4)"
      Top             =   600
      Width           =   4335
   End
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
      Picture         =   "frmRunSteps.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Import step description and expected results from an excel file"
      Top             =   1800
      Width           =   2205
   End
   Begin VB.ListBox lstUpdateFields 
      Columns         =   4
      Height          =   960
      ItemData        =   "frmRunSteps.frx":1070
      Left            =   4560
      List            =   "frmRunSteps.frx":1072
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
         Left            =   1980
         TabIndex        =   11
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
            Picture         =   "frmRunSteps.frx":1074
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunSteps.frx":1306
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunSteps.frx":1598
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView QCTree 
      Height          =   4935
      Left            =   60
      TabIndex        =   0
      Top             =   1020
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8705
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
      Cols            =   14
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
            Picture         =   "frmRunSteps.frx":1826
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunSteps.frx":1F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunSteps.frx":264A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunSteps.frx":2D5C
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
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmRunSteps.frx":346E
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   21132
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
            Picture         =   "frmRunSteps.frx":39BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunSteps.frx":3CA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunSteps.frx":41F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunSteps.frx":4743
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgOpenXML 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "TAO Logs | *.xml*"
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
Attribute VB_Name = "frmRunSteps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UpdateList()
Dim RN_HOST As String
Dim RN_STATUS As String

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
Dim k
Dim LastCYID
Dim a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, last_a1, CYID

    TimeSt = Format(Now, "mmm-dd-yyyy hhmmss") & "-"
    If chkCSV.Value = Checked Then
      AllF = InputBox("Enter file name", "File name", "[Run Steps] ")
    Else
      AllF = "[Run Steps] "
    End If
    
    ReDim CheckedItems(0): strPath = ""
    GetAllCheckedItems QCTree.Nodes(1)
    For j = LBound(CheckedItems) To UBound(CheckedItems) - 1
        If Left(CheckedItems(j), 1) = "F" Then
            strPath = strPath & "CF_ITEM_PATH LIKE '" & GetFromTable(Right(CheckedItems(j), Len(CheckedItems(j)) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%'" & " OR "
        ElseIf Left(CheckedItems(j), 1) = "T" Then
            strPath = strPath & "TC_TESTCYCL_ID = " & Right(CheckedItems(j), Len(CheckedItems(j)) - 1) & " OR "
        Else
            strPath = strPath & "TC_CYCLE_ID = " & Right(CheckedItems(j), Len(CheckedItems(j)) - 1) & " OR "
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
    
    If Trim(txtFilter.Text) <> "" Then
      objCommand.CommandText = "SELECT ST_ID, CY_CYCLE_ID, TC_TESTCYCL_ID, RN_RUN_ID, CF_ITEM_NAME, CY_CYCLE, TS_NAME, RN_RUN_NAME, ST_STEP_ORDER, ST_STEP_NAME, ST_DESCRIPTION, ST_ACTUAL, ST_EXPECTED, ST_STATUS, ST_EXECUTION_DATE, ST_EXECUTION_TIME, ST_LEVEL  FROM COMPONENT, TEST, CYCLE, TESTCYCL, RUN, STEP, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE RN_TEST_ID = TS_TEST_ID AND RN_RUN_ID = ST_RUN_ID AND  RN_TESTCYCL_ID = TC_TESTCYCL_ID AND RN_CYCLE_ID = CY_CYCLE_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & _
                              strPath & " AND " & Trim(txtFilter.Text) & " ORDER BY CY_CYCLE_ID, TC_TESTCYCL_ID, RN_RUN_ID, ST_STEP_ORDER "
    Else
      objCommand.CommandText = "SELECT ST_ID, CY_CYCLE_ID, TC_TESTCYCL_ID, RN_RUN_ID, CF_ITEM_NAME, CY_CYCLE, TS_NAME, RN_RUN_NAME, ST_STEP_ORDER, ST_STEP_NAME, ST_DESCRIPTION, ST_ACTUAL, ST_EXPECTED, ST_STATUS, ST_EXECUTION_DATE, ST_EXECUTION_TIME, ST_LEVEL  FROM COMPONENT, TEST, CYCLE, TESTCYCL, RUN, STEP, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE RN_TEST_ID = TS_TEST_ID AND RN_RUN_ID = ST_RUN_ID AND  RN_TESTCYCL_ID = TC_TESTCYCL_ID AND RN_CYCLE_ID = CY_CYCLE_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & _
                              strPath & " ORDER BY CY_CYCLE_ID, TC_TESTCYCL_ID, RN_RUN_ID, ST_STEP_ORDER "
    End If
    
    Debug.Print Me.Caption & "-" & objCommand.CommandText
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[SQL] " & Now & " " & objCommand.CommandText
    Set rs = objCommand.Execute
    If rs.RecordCount > 10000 And chkCSV.Value = Unchecked Then '***
        MsgBox "The records found exceeds 2500 records. It will be automatically generated as a CSV file.", vbOKOnly
        chkCSV.Value = Checked
        Exit Sub
        GenerateOutput
    End If '***
    AllScript = """" & "Step ID" & """" & "," & """" & "Test Set Folder" & """" & "," & """" & "Test Set" & """" & "," & """" & "Test Instance" & """" & "," & """" & "Run" & """" & "," & """" & "Step Order" & """" & "," & """" & "Step Name" & """" & "," & """" & "Description" & """" & "," & """" & "Expected" & """" & "," & """" & "Actual" & """" & "," & """" & "Status" & """" & "," & """" & "Exec Date" & """" & "," & """" & "Exec Time" & """" & "," & """" & "Level" & """"
    ClearTable
    If chkCSV.Value = Unchecked Then '***
        flxImport.Rows = rs.RecordCount + 1
    End If '***
    k = 0
    mdiMain.pBar.Max = rs.RecordCount + 3
    For i = 1 To rs.RecordCount
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Processing " & i & " out of " & rs.RecordCount
                If chkCSV.Value = Unchecked Then
                  k = k + 1
                  flxImport.Rows = k + 1
                  flxImport.TextMatrix(k, 0) = rs.FieldValue("ST_ID")
                  If rs.FieldValue("CY_CYCLE_ID") <> LastCYID Then
                      flxImport.TextMatrix(k, 1) = GetTestSetFolderPath(rs.FieldValue("CY_CYCLE_ID"))
                  Else
                      flxImport.TextMatrix(k, 1) = flxImport.TextMatrix(k - 1, 1)
                  End If
                  LastCYID = rs.FieldValue("CY_CYCLE_ID")
                  flxImport.TextMatrix(k, 2) = ReplaceAllEnter(rs.FieldValue("CY_CYCLE"))
                  flxImport.TextMatrix(k, 3) = rs.FieldValue("TS_NAME")
                  flxImport.TextMatrix(k, 4) = ReplaceAllEnter(rs.FieldValue("RN_RUN_NAME"))
                  flxImport.TextMatrix(k, 5) = rs.FieldValue("ST_STEP_ORDER")
                  flxImport.TextMatrix(k, 6) = ReplaceAllEnter(rs.FieldValue("ST_STEP_NAME"))
                  flxImport.TextMatrix(k, 7) = ReplaceAllEnter(rs.FieldValue("ST_DESCRIPTION"))
                  flxImport.TextMatrix(k, 7) = Replace(flxImport.TextMatrix(k, 7), "<html><body>", "")
                  flxImport.TextMatrix(k, 7) = Replace(flxImport.TextMatrix(k, 7), "</body></html>", "")
                  flxImport.TextMatrix(k, 7) = Replace(flxImport.TextMatrix(k, 7), "<b>", "")
                  flxImport.TextMatrix(k, 7) = Replace(flxImport.TextMatrix(k, 7), "</b>", "")
                  flxImport.TextMatrix(k, 7) = Replace(flxImport.TextMatrix(k, 7), "<u>", "")
                  flxImport.TextMatrix(k, 7) = Replace(flxImport.TextMatrix(k, 7), "</u>", "")
                  flxImport.TextMatrix(k, 7) = Replace(flxImport.TextMatrix(k, 7), "<br>", "")
                  flxImport.TextMatrix(k, 7) = Replace(flxImport.TextMatrix(k, 7), "<font color=#800000>", "")
                  flxImport.TextMatrix(k, 7) = Replace(flxImport.TextMatrix(k, 7), "</font>", "")
                  flxImport.TextMatrix(k, 7) = CleanHTML(flxImport.TextMatrix(k, 7))
                  flxImport.TextMatrix(k, 8) = ReplaceAllEnter(rs.FieldValue("ST_EXPECTED"))
                  flxImport.TextMatrix(k, 8) = Replace(flxImport.TextMatrix(k, 8), "<html><body>", "")
                  flxImport.TextMatrix(k, 8) = Replace(flxImport.TextMatrix(k, 8), "</body></html>", "")
                  flxImport.TextMatrix(k, 8) = Replace(flxImport.TextMatrix(k, 8), "<b>", "")
                  flxImport.TextMatrix(k, 8) = Replace(flxImport.TextMatrix(k, 8), "</b>", "")
                  flxImport.TextMatrix(k, 8) = Replace(flxImport.TextMatrix(k, 8), "<u>", "")
                  flxImport.TextMatrix(k, 8) = Replace(flxImport.TextMatrix(k, 8), "</u>", "")
                  flxImport.TextMatrix(k, 8) = Replace(flxImport.TextMatrix(k, 8), "<br>", "")
                  flxImport.TextMatrix(k, 8) = Replace(flxImport.TextMatrix(k, 8), "<font color=#800000>", "")
                  flxImport.TextMatrix(k, 8) = Replace(flxImport.TextMatrix(k, 8), "</font>", "")
                  flxImport.TextMatrix(k, 8) = CleanHTML(flxImport.TextMatrix(k, 8))
                  flxImport.TextMatrix(k, 9) = ReplaceAllEnter(rs.FieldValue("ST_ACTUAL"))
                  flxImport.TextMatrix(k, 10) = rs.FieldValue("ST_STATUS")
                  flxImport.TextMatrix(k, 11) = Format(rs.FieldValue("ST_EXECUTION_DATE"), "dd/mmm/yyyy")
                  flxImport.TextMatrix(k, 12) = Format(rs.FieldValue("ST_EXECUTION_TIME"), "hh:mm:ss")
                  flxImport.TextMatrix(k, 13) = rs.FieldValue("ST_LEVEL")
                Else
                  a0 = rs.FieldValue("ST_ID")
                  CYID = rs.FieldValue("CY_CYCLE_ID")
                  If CYID <> LastCYID Then
                    a1 = GetTestSetFolderPath(rs.FieldValue("CY_CYCLE_ID"))
                    last_a1 = a1
                    LastCYID = CYID
                  Else
                    a1 = last_a1
                  End If
                  a2 = ReplaceAllEnter(rs.FieldValue("CY_CYCLE"))
                  a3 = rs.FieldValue("TS_NAME")
                  a4 = ReplaceAllEnter(rs.FieldValue("RN_RUN_NAME"))
                  a5 = rs.FieldValue("ST_STEP_ORDER")
                  a6 = ReplaceAllEnter(rs.FieldValue("ST_STEP_NAME"))
                  a7 = ReplaceAllEnter(rs.FieldValue("ST_DESCRIPTION"))
                  a7 = Replace(a7, "<html><body>", "")
                  a7 = Replace(a7, "</body></html>", "")
                  a7 = Replace(a7, "<b>", "")
                  a7 = Replace(a7, "</b>", "")
                  a7 = Replace(a7, "<u>", "")
                  a7 = Replace(a7, "</u>", "")
                  a7 = Replace(a7, "<br>", "")
                  a7 = Replace(a7, "<font color=#800000>", "")
                  a7 = Replace(a7, "</font>", "")
                  a7 = CleanHTML(CStr(a7))
                  a8 = ReplaceAllEnter(rs.FieldValue("ST_EXPECTED"))
                  a8 = Replace(a8, "<html><body>", "")
                  a8 = Replace(a8, "</body></html>", "")
                  a8 = Replace(a8, "<b>", "")
                  a8 = Replace(a8, "</b>", "")
                  a8 = Replace(a8, "<u>", "")
                  a8 = Replace(a8, "</u>", "")
                  a8 = Replace(a8, "<br>", "")
                  a8 = Replace(a8, "<font color=#800000>", "")
                  a8 = Replace(a8, "</font>", "")
                  a8 = CleanHTML(CStr(a8))
                  a9 = ReplaceAllEnter(rs.FieldValue("ST_ACTUAL"))
                  a10 = rs.FieldValue("ST_STATUS")
                  a11 = Format(rs.FieldValue("ST_EXECUTION_DATE"), "dd/mmm/yyyy")
                  a12 = Format(rs.FieldValue("ST_EXECUTION_TIME"), "hh:mm:ss")
                  a13 = rs.FieldValue("ST_LEVEL")
                  If Trim(AllScript) <> "" Then
                        AllScript = AllScript & vbCrLf & """" & a0 & """" & "," & """" & a1 & """" & "," & """" & a2 & """" & "," & """" & a3 & """" & "," & """" & a4 & """" & "," & """" & a5 & """" & "," & """" & a6 & """" & "," & """" & a7 & """" & "," & """" & a8 & """" & "," & """" & a9 & """" & "," & """" & a10 & """" & "," & """" & a11 & """" & "," & """" & a12 & """" & "," & """" & a13 & """"
                  Else
                        AllScript = AllScript & """" & a0 & """" & "," & """" & a1 & """" & "," & """" & a2 & """" & "," & """" & a3 & """" & "," & """" & a4 & """" & "," & """" & a5 & """" & "," & """" & a6 & """" & "," & """" & a7 & """" & "," & """" & a8 & """" & "," & """" & a9 & """" & "," & """" & a10 & """" & "," & """" & a11 & """" & "," & """" & a12 & """" & "," & """" & a13 & """"
                  End If
                End If
                If chkCSV.Value = Checked Then
                    If i Mod 500 = 0 Then
                        FileAppend App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv", AllScript
                        AllScript = ""
                    End If
                End If
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
        rs.Next
    Next
    mdiMain.pBar.Value = mdiMain.pBar.Max
    FXGirl.EZPlay FXSQCExtractCompleted
If chkCSV.Value = Checked Then '***
    FileAppend App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv", AllScript: If MsgBox("Successfully exported to " & App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv" & vbCrLf & "Do you want to launch the extracted file?", vbYesNo) = vbYes Then Shell "explorer.exe " & App.path & "\SQC Logs" & "\", vbNormalFocus
    AllScript = vbCrLf & " ,"
    AllScript = AllScript & """" & "SQL Code:" & """" & "," & """" & Replace(objCommand.CommandText, """", "'") & """"
    FileAppend App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv", AllScript
End If '***
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

Private Sub chkAutoDates_Click()
If chkAutoDates.Value = Checked Then
    If MsgBox("Are you sure you want to auto compute the dates?", vbYesNo) = vbYes Then
        chkAutoDates.Value = Checked
    Else
        chkAutoDates.Value = Unchecked
    End If
Else
    chkAutoDates.Value = Unchecked
End If
End Sub

Private Sub chkCSV_Click()
If chkCSV.Value = Checked Then
    If MsgBox("Are you sure you want to download directly to CSV?", vbYesNo) = vbYes Then
        chkCSV.Value = Checked
    Else
        chkCSV.Value = Unchecked
    End If
End If
End Sub

Private Sub chkInvertDate_Click()
If chkInvertDate.Value = Checked Then
    If MsgBox("Are you sure you want to invert the dates?", vbYesNo) = vbYes Then
        chkInvertDate.Value = Checked
    Else
        chkInvertDate.Value = Unchecked
    End If
Else
    chkInvertDate.Value = Unchecked
End If
End Sub

Private Sub chkUpload_Click()
If chkUpload.Value = Checked Then
    If MsgBox("Are you sure you want to do a new upload?", vbYesNo) = vbYes Then
        chkUpload.Value = Checked
        lstUpdateFields.Visible = False
        Label1.Visible = False
        ClearForm 'Uncomment
        Toolbar1.Buttons(2).Visible = False
    Else
        chkUpload.Value = Unchecked
        lstUpdateFields.Visible = True
        Label1.Visible = True
        ClearForm 'Uncomment
        Toolbar1.Buttons(2).Visible = True
    End If
Else
    chkUpload.Value = Unchecked
    lstUpdateFields.Visible = True
    Label1.Visible = True
    ClearForm
    Toolbar1.Buttons(2).Visible = True
End If
End Sub

Private Sub cmdLoadExcel_Click()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim fname As String
Dim lastrow
Dim i, j, FileFunct As New clsFiles, stringFunct As New clsStrings
Dim tmpSplit, X, tmpDate As String

If chkUpload.Value = Unchecked Then
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
             If InStr(1, GetCommentText(.Range("A1")), "Run Steps") = 0 Then
                MsgBox "Import file is invalid. Please use only sheets generated by the SuperQC"
                xlWB.Close
                xlObject.Application.Quit
                Set xlWB = Nothing
                Set xlObject = Nothing
                Exit Sub
             End If
             For i = 1 To 14
                If .Range(ColumnLetter(CInt(i)) & 1).Interior.ColorIndex = 35 Then
                    For j = 0 To lstUpdateFields.ListCount - 1
                        If .Range(ColumnLetter(CInt(i)) & 1).Value = lstUpdateFields.List(j) Then
                            lstUpdateFields.Selected(j) = True
                        End If
                    Next
                End If
             Next
             lastrow = .Range("A" & .Rows.Count).End(xlUp).row
            .Range("A1:AB" & lastrow).COPY 'Set selection to Copy
        End With
           
        With flxImport
            .Clear
            .Redraw = False     'Dont draw until the end, so we avoid that flash
            .row = 0            'Paste from first cell
            .col = 0
            .Rows = lastrow
            .Cols = 14
            .RowSel = lastrow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
            .ColSel = 14 - 1
        End With
        
         With flxImport
            .Clear
            .Redraw = False     'Dont draw until the end, so we avoid that flash
            .row = 0            'Paste from first cell
            .col = 0
            .Rows = lastrow
            .Cols = 14
            .RowSel = lastrow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
            .ColSel = 14 - 1
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
        FXGirl.EZPlay FXExportToExcel
        mdiMain.pBar.Max = 100
        mdiMain.pBar.Value = 100
Else
    On Error GoTo ErrLoad
    If Left(QCTree.SelectedItem.Key, 1) <> "T" Then MsgBox "Select Test Instance from the HPQC tree": Exit Sub
    dlgOpenXML.filename = "": dlgOpenXML.ShowOpen
    fname = dlgOpenXML.filename
    If Trim(fname) = "" Then Exit Sub
    FileFunct.LoadXMLDocument fname
    tmpSplit = Split(FileFunct.LoadedXMLData, vbCrLf)
    flxImport.Clear
    flxImport.Cols = 8
    flxImport.TextMatrix(0, 0) = "Step Number"
    flxImport.TextMatrix(0, 1) = "Execution Time"
    flxImport.TextMatrix(0, 2) = "Elapsed Time"
    flxImport.TextMatrix(0, 3) = "Step Result"
    flxImport.TextMatrix(0, 4) = "Step Summary"
    flxImport.TextMatrix(0, 5) = "Component Name"
    flxImport.TextMatrix(0, 6) = "Step Description"
    flxImport.TextMatrix(0, 7) = "Image Path"
    flxImport.Rows = UBound(tmpSplit) + 1
    X = 1
    For i = LBound(tmpSplit) To UBound(tmpSplit) - 1
        flxImport.TextMatrix(i + 1, 0) = i + 1
        If InStr(1, tmpSplit(i), "PC:") <> 0 Then
            RN_HOST = Trim(Replace(tmpSplit(i), "PC:", ""))
        ElseIf InStr(1, tmpSplit(i), "STATUS:") <> 0 Then
            RN_STATUS = Trim(Replace(tmpSplit(i), "STATUS:", ""))
        ElseIf InStr(1, tmpSplit(i), "EXECUTION_TIME:") <> 0 Then
            If chkInvertDate.Value = Checked Then
                tmpDate = stringFunct.Switch_Month_Day_DateFormat(CDate((Trim(Replace(tmpSplit(i), "EXECUTION_TIME:", "")))), "/", "dd/mmm/yyyy")
            Else
                tmpDate = Format(CDate((Trim(Replace(tmpSplit(i), "EXECUTION_TIME:", "")))), "dd/mmm/yyyy")
            End If
            flxImport.TextMatrix(X, 1) = tmpDate & " " & Format(CDate((Trim(Replace(tmpSplit(i), "EXECUTION_TIME:", "")))), "hh:mm:ss")
        ElseIf InStr(1, tmpSplit(i), "ELAPSED_TIME:") <> 0 Then
            flxImport.TextMatrix(X, 2) = Trim(Replace(tmpSplit(i), "ELAPSED_TIME:", ""))
        ElseIf InStr(1, tmpSplit(i), "STEP_RESULT:") <> 0 Then
            flxImport.TextMatrix(X, 3) = Trim(Replace(tmpSplit(i), "STEP_RESULT:", ""))
        ElseIf InStr(1, tmpSplit(i), "STEP_SUMMARY:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "STEP_SUMMARY:", ""))
        ElseIf InStr(1, tmpSplit(i), "COMPONENT_NAME:") <> 0 Then
            flxImport.TextMatrix(X, 5) = Trim(Replace(tmpSplit(i), "COMPONENT_NAME:", ""))
        ElseIf InStr(1, tmpSplit(i), "STEP_DESCRIPTION:") <> 0 Then
            flxImport.TextMatrix(X, 6) = Trim(Replace(tmpSplit(i), "STEP_DESCRIPTION:", ""))
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "IMAGE_PATH:") <> 0 Then
            flxImport.TextMatrix(X - 1, 7) = Trim(flxImport.TextMatrix(X, 2) & " Captured Image Stored in: " & Trim(Replace(tmpSplit(i), "IMAGE_PATH:", "")))
        End If
        If chkAutoDates.Value = Checked Then
            If X <> "1" Then
                flxImport.TextMatrix(X, 1) = Format(DateAdd("s", Val(flxImport.TextMatrix(X, 2)), flxImport.TextMatrix(X - 1, 1)), "dd/mmm/yyyy hh:mm:ss")
            End If
        End If
    Next
    flxImport.Rows = X + 1
    flxImport.TextMatrix(X, 1) = Format(DateAdd("s", 10, flxImport.TextMatrix(X - 1, 1)), "dd/mmm/yyyy hh:mm:ss")
    flxImport.TextMatrix(X, 4) = "End Business Component"
    On Error GoTo 0
    On Error Resume Next
    flxImport.TextMatrix(X, 6) = "End Business Component " & QCTree.SelectedItem.Text 'Uncomment
    If Err.Number <> 0 Then MsgBox "Select Test Instance from the HPQC tree"
    flxImport.TextMatrix(X, 2) = "10"
    flxImport.TextMatrix(X, 3) = "DONE"
End If
Exit Sub
ErrLoad:
MsgBox "There was an error while importing the file. Please refresh and close all excel and try again" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub cmdLoadExcel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim fname As String
Dim lastrow
Dim i, j, FileFunct As New clsFiles, stringFunct As New clsStrings

If Button = 2 Then
    If chkUpload.Value = Checked Then
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
                 If InStr(1, GetCommentText(.Range("A1")), "Run Steps") = 0 Then
                    MsgBox "Import file is invalid. Please use only sheets generated by the SuperQC"
                    xlWB.Close
                    xlObject.Application.Quit
                    Set xlWB = Nothing
                    Set xlObject = Nothing
                    Exit Sub
                 End If
                 For i = 1 To 8
                    If .Range(ColumnLetter(CInt(i)) & 1).Interior.ColorIndex = 35 Then
                        For j = 0 To lstUpdateFields.ListCount - 1
                            If .Range(ColumnLetter(CInt(i)) & 1).Value = lstUpdateFields.List(j) Then
                                lstUpdateFields.Selected(j) = True
                            End If
                        Next
                    End If
                 Next
                 lastrow = .Range("A" & .Rows.Count).End(xlUp).row
                .Range("A1:H" & lastrow).COPY 'Set selection to Copy
            End With
               
            With flxImport
                .Clear
                .Redraw = False     'Dont draw until the end, so we avoid that flash
                .row = 0            'Paste from first cell
                .col = 0
                .Rows = lastrow
                .Cols = 8
                .RowSel = lastrow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
                .ColSel = 8 - 1
            End With
            
             With flxImport
                .Clear
                .Redraw = False     'Dont draw until the end, so we avoid that flash
                .row = 0            'Paste from first cell
                .col = 0
                .Rows = lastrow
                .Cols = 8
                .RowSel = lastrow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
                .ColSel = 8 - 1
                .Clip = Replace(Clipboard.GetText, vbNewLine, vbCr)   'Replace carriage return with the correct one
                .col = 1            'Just to remove that blue selection from Flexgrid
                .FixedCols = 1
                .FixedRows = 1
                .Redraw = True      'Now draw
            End With
                
            xlObject.DisplayAlerts = False 'To avoid "Save woorkbook" messagebox
            
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = flxImport.Rows - 1 & " record(s) loaded"
            
            'Close Excel
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            FXGirl.EZPlay FXExportToExcel
            mdiMain.pBar.Max = 100
            mdiMain.pBar.Value = 100
    End If
End If
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
ClearForm 'Uncomment
End Sub

Private Sub Form_Resize()
On Error Resume Next
QCTree.height = stsBar.Top - 1150
lstUpdateFields.width = Me.width - lstUpdateFields.Left - 350
flxImport.height = stsBar.Top - flxImport.Top - 250
flxImport.width = Me.width - flxImport.Left - 350
End Sub

Private Sub Label1_Click()
Dim tmpPath, tmpID: On Error Resume Next
tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%"
frmLogs.txtLogs.Text = "Test Instace ID: " & tmpID & vbCrLf & "CF_ITEM_PATH: " & tmpPath & vbCrLf & "Folder Path: " & QCTree.SelectedItem.FullPath
frmLogs.Show 1
End Sub

Private Sub QCTree_DblClick()
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
    
    If Left(QCTree.SelectedItem.Key, 1) = "F" Then
        Set objCommand = QCConnection.Command
        objCommand.CommandText = "SELECT DISTINCT CY_CYCLE, CY_CYCLE_ID FROM CYCLE, CYCL_FOLD WHERE CY_FOLDER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY CY_CYCLE"
        Set rs = objCommand.Execute
        For i = 1 To rs.RecordCount
            QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("C" & rs.FieldValue("CY_CYCLE_ID")), rs.FieldValue("CY_CYCLE"), 2
            rs.Next
        Next
    End If
    
    If Left(QCTree.SelectedItem.Key, 1) = "C" Then
        Set objCommand = QCConnection.Command
        objCommand.CommandText = "SELECT DISTINCT TS_NAME, TC_TESTCYCL_ID, TC_TEST_ORDER FROM TEST, TESTCYCL WHERE TC_CYCLE_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " AND TC_TEST_ID = TS_TEST_ID ORDER BY TC_TEST_ORDER "
        Set rs = objCommand.Execute
        For i = 1 To rs.RecordCount
            QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("T" & rs.FieldValue("TC_TESTCYCL_ID")), "[" & rs.FieldValue("TC_TEST_ORDER") & "]" & rs.FieldValue("TS_NAME"), 3
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
    chkUpload.Value = Unchecked
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
    If chkUpload.Value = Unchecked Then
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
    Else
        If flxImport.Rows <= 1 Then
            MsgBox "Nothing to output", vbInformation
        Else
            If GetEditableFields() = "IV:IV" Then
                If MsgBox("You have selected nothing to update on this sheet. The whole sheet will be read-only. Do you want to proceed?", vbYesNo) = vbYes Then
                    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
                    OutputTable_XML
                End If
            Else
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
                OutputTable_XML
            End If
        End If
    End If
Case "cmdUpload"
If IncorrectHeaderDetails = False Or chkUpload.Value = Unchecked Then
    If GetEditableFields <> "IV:IV" Or chkUpload.Value = Unchecked Then
        If MsgBox("Are you sure you want to mass update " & flxImport.Rows - 1 & " record(s) of Run Steps?", vbYesNo) = vbYes Then
            Randomize: tmpR = CInt(Rnd(1000) * 10000)
            If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
                If chkUpload.Value = Unchecked Then
                    GetUpdateList
                    Upload_Test_Instance
                Else
                    If Left(QCTree.SelectedItem.Key, 1) <> "T" Then
                        MsgBox "Select a Test Instance from the QC tree then try again"
                        Exit Sub
                    Else
                        If IncorrectHeaderDetails_XML = False Then
                            If MsgBox("Are you sure you want to create a new report log for " & QCTree.SelectedItem.Text & "?", vbYesNo) = vbYes Then
                                AddRunSteps Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), QCTree.SelectedItem.Text
                            End If
                        Else
                            MsgBox "Invalid or Incorrect upload sheet file selected", vbCritical
                        End If
                    End If
                End If
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

QCTree.Nodes.Clear

Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
    QCTree.Nodes.Add , , "Root", "Root", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = 0 ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    QCTree.Nodes(1).Selected = True
     lstUpdateFields.Clear
     lstUpdateFields.AddItem "Step Order"
     lstUpdateFields.AddItem "Step Name"
     lstUpdateFields.AddItem "Description"
     lstUpdateFields.AddItem "Expected"
     lstUpdateFields.AddItem "Actual"
     lstUpdateFields.AddItem "Status"
     lstUpdateFields.AddItem "Exec Date"
     lstUpdateFields.AddItem "Exec Time"
     lstUpdateFields.AddItem "Level"
     Me.Caption = Me.Tag
     txtFilter.Text = ""
     txtFilter.Locked = True
     If chkUpload.Value = Checked Then
        lstUpdateFields.Visible = False
        Label1.Visible = False
        ClearTable_XML
     Else
        lstUpdateFields.Visible = True
        Label1.Visible = True
        ClearTable
     End If
     Toolbar1.Buttons(2).Visible = True
     RN_HOST = ""
     RN_STATUS = "FAILED"
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.Cols = 14
flxImport.TextMatrix(0, 0) = "Step ID"
flxImport.TextMatrix(0, 1) = "Test Set Folder"
flxImport.TextMatrix(0, 2) = "Test Set"
flxImport.TextMatrix(0, 3) = "Test Instance"
flxImport.TextMatrix(0, 4) = "Run"
flxImport.TextMatrix(0, 5) = "Step Order"
flxImport.TextMatrix(0, 6) = "Step Name"
flxImport.TextMatrix(0, 7) = "Description"
flxImport.TextMatrix(0, 8) = "Expected"
flxImport.TextMatrix(0, 9) = "Actual"
flxImport.TextMatrix(0, 10) = "Status"
flxImport.TextMatrix(0, 11) = "Exec Date"
flxImport.TextMatrix(0, 12) = "Exec Time"
flxImport.TextMatrix(0, 13) = "Level"
flxImport.Rows = 2
flxImport.FixedCols = 1
flxImport.FixedRows = 1
End Sub

Private Sub ClearTable_XML()
flxImport.Clear
flxImport.Cols = 8
flxImport.TextMatrix(0, 0) = "Step Number"
flxImport.TextMatrix(0, 1) = "Execution Time"
flxImport.TextMatrix(0, 2) = "Elapsed Time"
flxImport.TextMatrix(0, 3) = "Step Result"
flxImport.TextMatrix(0, 4) = "Step Summary"
flxImport.TextMatrix(0, 5) = "Component Name"
flxImport.TextMatrix(0, 6) = "Step Description"
flxImport.TextMatrix(0, 7) = "Image Path"
flxImport.Rows = 2
flxImport.FixedCols = 1
flxImport.FixedRows = 1
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
    xlObject.Sheets("Sheet2").Range("A7").Value = "4 - Planned Execution Start and End Date format is dd-mmm-yyyy"
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
  curTab = "RUN_STEPS01"
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
  flxImport.FixedCols = 1
  flxImport.FixedRows = 1
  
  xlObject.Sheets(curTab).Range("A1").Select
  xlObject.Sheets(curTab).Paste

'On Error Resume Next
    xlObject.Sheets(curTab).Range("A:N").Select
        
    xlObject.Sheets(curTab).Range("A:N").Borders(xlDiagonalDown).LineStyle = xlNone
    xlObject.Sheets(curTab).Range("A:N").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlObject.Sheets(curTab).Range("A:N").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:N").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:N").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:N").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:N").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:N").Borders(xlInsideHorizontal)
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
    xlObject.Sheets(curTab).Range("A:N").Select
    xlObject.Sheets(curTab).Range("A:N").EntireColumn.AutoFit
    xlObject.Sheets(curTab).Range("A1").Select
    
    xlObject.Sheets(curTab).Range("A1").AddComment
    xlObject.Sheets(curTab).Range("A1").Comment.Visible = False
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & mdiMain.Caption & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""
    
    xlObject.Sheets(curTab).Range(GetEditableFields).Interior.ColorIndex = 35
    xlObject.Sheets(curTab).Protection.AllowEditRanges.Add Title:="Range1", Range:=xlObject.Sheets(curTab).Range(GetEditableFields)
    xlObject.Sheets(curTab).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
  xlObject.Workbooks(1).SaveAs "RUN_STEPS01-" & CleanTheString(QCTree.SelectedItem.Text) & "-" & Format(Now, "mmddyyyy HHMM AMPM")
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate
  
  Set xlWB = Nothing
  Set xlObject = Nothing
  
  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export to MS Excel completed.": Exit Sub:
OutErr:     MsgBox Err.Description, vbCritical: xlObject.Visible = True: xlObject.ActiveWindow.Activate: Set xlWB = Nothing: Set xlObject = Nothing
On Error GoTo 0
End Sub

Private Sub OutputTable_XML()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim i, Protections
Dim curTab
Dim w, tmp


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
    xlObject.Sheets("Sheet2").Range("A7").Value = "4 - Planned Execution Start and End Date format is dd-mmm-yyyy"
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
  curTab = "RUN_STEPS01"
  xlObject.Sheets("Sheet1").Name = curTab
  flxImport.FixedCols = 0
  flxImport.FixedRows = 0
  flxImport.col = 0
  flxImport.row = 0
  Pause 1
  flxImport.RowSel = flxImport.Rows - 1
  flxImport.ColSel = flxImport.Cols - 1
  For i = 0 To flxImport.Rows - 1
    xlObject.Sheets(curTab).Range("A" & i + 1).Value = flxImport.TextMatrix(i, 0)
    xlObject.Sheets(curTab).Range("B" & i + 1).Value = flxImport.TextMatrix(i, 1)
    xlObject.Sheets(curTab).Range("C" & i + 1).Value = flxImport.TextMatrix(i, 2)
    xlObject.Sheets(curTab).Range("D" & i + 1).Value = flxImport.TextMatrix(i, 3)
    xlObject.Sheets(curTab).Range("E" & i + 1).Value = flxImport.TextMatrix(i, 4)
    xlObject.Sheets(curTab).Range("F" & i + 1).Value = flxImport.TextMatrix(i, 5)
    xlObject.Sheets(curTab).Range("G" & i + 1).Value = flxImport.TextMatrix(i, 6)
    xlObject.Sheets(curTab).Range("H" & i + 1).Value = flxImport.TextMatrix(i, 7)
  Next
  flxImport.FixedCols = 1
  flxImport.FixedRows = 1

'On Error Resume Next
    xlObject.Sheets(curTab).Range("A:H").Select
        
    xlObject.Sheets(curTab).Range("A:H").Borders(xlDiagonalDown).LineStyle = xlNone
    xlObject.Sheets(curTab).Range("A:H").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlObject.Sheets(curTab).Range("A:H").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:H").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:H").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:H").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:H").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:H").Borders(xlInsideHorizontal)
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
    xlObject.Sheets(curTab).Range("A:H").Select
    xlObject.Sheets(curTab).Range("A:H").EntireColumn.AutoFit
    xlObject.Sheets(curTab).Range("A1").Select
    
    xlObject.Sheets(curTab).Range("A1").AddComment
    xlObject.Sheets(curTab).Range("A1").Comment.Visible = False
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & mdiMain.Caption & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""
    
  xlObject.Workbooks(1).SaveAs "RUN_STEPS01-" & CleanTheString(QCTree.SelectedItem.Text) & "-" & Format(Now, "mmddyyyy HHMM AMPM")
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate
  
  Set xlWB = Nothing
  Set xlObject = Nothing
  
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
   If chkUpload.Value = Unchecked Then
    GetEditableFields = "IV:IV"
   Else
    GetEditableFields = "A:A"
   End If
End If
End Function

Private Function Upload_Test_Instance()
Dim i
Dim objCommand
Dim rs
Dim numErr
Dim tsTestF As TSTestFactory, tmp
Dim tsTest As tsTest, tmpTestID, tmpTestSetID, tmpOrder, tmpStartDate, tmpEndDate, tmpScripter, tmpQA, tmpGroup, tmpDataScript

Dim tfact As TestFactory
Dim mytest As Test
Dim TestSetFact As TestSetFactory
Dim mytestset As TestSet
Dim tsttestsetFact As TSTestFactory
Dim mytsttestset As tsTest

numErr = 0
    mdiMain.pBar.Max = UBound(UpdateList) + 3
    For i = LBound(UpdateList) To UBound(UpdateList)
        On Error Resume Next
        If UpdateList(i) = "" Then mdiMain.pBar.Value = mdiMain.pBar.Max: Exit Function
            Set objCommand = QCConnection.Command
            objCommand.CommandText = UpdateList(i)
            Debug.Print objCommand.CommandText
            Set rs = objCommand.Execute
            Debug.Print i
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Updating Test Instance " & i & " of " & UBound(UpdateList) + 1 & " (" & numErr & ") errors found"
            If Err.Number <> 0 Then
                FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Test Instance (FAILED) " & Now & " " & objCommand.CommandText & " " & Err.Description
                numErr = numErr + 1
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(3).Picture: stsBar.Panels(2).Text = "Updating Test Instance " & i & " of " & UBound(UpdateList) + 1 & " (" & numErr & ") errors found"
            Else
                FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Test Instance (PASSED) " & Now & " " & objCommand.CommandText
            End If
            Err.Clear
            On Error GoTo 0
            Set objCommand = Nothing
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
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = "Updated Test Instance " & i & " of " & UBound(UpdateList) + 1 & " (" & numErr & ") errors found"
    QCConnection.SendMail "user@companyemail.com", "", "[HPQC UPDATES] Updated Test Instance  by " & curUser & " in " & curDomain & "-" & curProject, "Updated Test Instance " & i & " of " & UBound(UpdateList) + 1 & " (" & numErr & ") errors found" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
    QCConnection.SendMail curUser, "", "[HPQC UPDATES] Updated Test Instance  by " & curUser & " in " & curDomain & "-" & curProject, "Updated Test Instance " & i & " of " & UBound(UpdateList) + 1 & " (" & numErr & ") errors found" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
    If numErr <> 0 Then
      Dim tmpFile As New clsFiles
      frmLogs.Caption = App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log"
      frmLogs.txtLogs.Text = tmpFile.ReadFromFile_FAILED(App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log")
      frmLogs.Show 1
    End If
End Function

Private Function GetUpdateList()
Dim i, X, j, tmpColVal
Dim objCommand
Dim rs
Dim curField As String
ReDim UpdateList(0)
X = -1
For i = 1 To flxImport.Rows - 1
        X = X + 1
        ReDim Preserve UpdateList(X)
        For j = 0 To lstUpdateFields.ListCount - 1
            If lstUpdateFields.Selected(j) = True Then
                curField = GetFieldNameinDB(lstUpdateFields.List(j))
                If curField = "ST_EXECUTION_DATE" Then
                    tmpColVal = tmpColVal & curField & " = '" & Format(flxImport.TextMatrix(i, GetFieldNumberinDB_byName(curField)), "dd/mmm/yyyy") & "', "
                ElseIf curField = "ST_EXECUTION_TIME" Then
                    tmpColVal = tmpColVal & curField & " = '" & Format(flxImport.TextMatrix(i, GetFieldNumberinDB_byName(curField)), "hh:mm:ss") & "', "
                Else
                    tmpColVal = tmpColVal & curField & " = '" & flxImport.TextMatrix(i, GetFieldNumberinDB_byName(curField)) & "', "
                End If
            End If
        Next
        tmpColVal = Left(tmpColVal, Len(tmpColVal) - 2)
        UpdateList(X) = "UPDATE STEP SET " & tmpColVal & " WHERE ST_ID = " & flxImport.TextMatrix(i, 0)
        tmpColVal = ""
Next
End Function
'========<><><><><><><> DATES SHOULD BE ENCLOSED WITH # # OR ???

Private Function GetFieldNameinDB(X As String)
Select Case X
    Case "Step ID"
        GetFieldNameinDB = "ST_ID"
    Case "Test Set Folder"
        GetFieldNameinDB = "CF_ITEM_NAME"
    Case "Test Set"
        GetFieldNameinDB = "CY_CYCLE"
    Case "Test Instance"
        GetFieldNameinDB = "TS_NAME"
    Case "Run"
        GetFieldNameinDB = "RN_RUN_NAME"
    Case "Step Order"
        GetFieldNameinDB = "ST_STEP_ORDER"
    Case "Step Name"
        GetFieldNameinDB = "ST_STEP_NAME"
    Case "Description"
        GetFieldNameinDB = "ST_DESCRIPTION"
    Case "Expected"
        GetFieldNameinDB = "ST_EXPECTED"
    Case "Actual"
        GetFieldNameinDB = "ST_ACTUAL"
    Case "Status"
        GetFieldNameinDB = "ST_STATUS"
    Case "Exec Date"
        GetFieldNameinDB = "ST_EXECUTION_DATE"
    Case "Exec Time"
        GetFieldNameinDB = "ST_EXECUTION_TIME"
    Case "Level"
        GetFieldNameinDB = "ST_LEVEL"
End Select
End Function

Private Function GetFieldNumberinDB(X As Integer)
Select Case X
    Case 0
        GetFieldNumberinDB = "ST_ID"
    Case 1
        GetFieldNumberinDB = "CF_ITEM_NAME"
    Case 2
        GetFieldNumberinDB = "CY_CYCLE"
    Case 3
        GetFieldNumberinDB = "TS_NAME"
    Case 4
        GetFieldNumberinDB = "RN_RUN_NAME"
    Case 5
        GetFieldNumberinDB = "ST_STEP_ORDER"
    Case 6
        GetFieldNumberinDB = "ST_STEP_NAME"
    Case 7
        GetFieldNumberinDB = "ST_DESCRIPTION"
    Case 8
        GetFieldNumberinDB = "ST_EXPECTED"
    Case 9
        GetFieldNumberinDB = "ST_ACTUAL"
    Case 10
        GetFieldNumberinDB = "ST_STATUS"
    Case 11
        GetFieldNumberinDB = "ST_EXECUTION_DATE"
    Case 12
        GetFieldNumberinDB = "ST_EXECUTION_TIME"
    Case 13
        GetFieldNumberinDB = "ST_LEVEL"
End Select
End Function

Private Function GetFieldNumberinDB_byName(X As String)
Select Case X
    Case "Step ID"
        GetFieldNumberinDB_byName = 0
    Case "Test Set Folder"
        GetFieldNumberinDB_byName = 1
    Case "Test Set"
        GetFieldNumberinDB_byName = 2
    Case "Test Instance"
        GetFieldNumberinDB_byName = 3
    Case "Run"
        GetFieldNumberinDB_byName = 4
    Case "Step Order"
        GetFieldNumberinDB_byName = 5
    Case "Step Name"
        GetFieldNumberinDB_byName = 6
    Case "Description"
        GetFieldNumberinDB_byName = 7
    Case "Expected"
        GetFieldNumberinDB_byName = 8
    Case "Actual"
        GetFieldNumberinDB_byName = 9
    Case "Status"
        GetFieldNumberinDB_byName = 10
    Case "Exec Date"
        GetFieldNumberinDB_byName = 11
    Case "Exec Time"
        GetFieldNumberinDB_byName = 12
    Case "Level"
        GetFieldNumberinDB_byName = 13
        
    Case "ST_ID"
        GetFieldNumberinDB_byName = 0
    Case "CF_ITEM_NAME"
        GetFieldNumberinDB_byName = 1
    Case "CY_CYCLE"
        GetFieldNumberinDB_byName = 2
    Case "TS_NAME"
        GetFieldNumberinDB_byName = 3
    Case "RN_RUN_NAME"
        GetFieldNumberinDB_byName = 4
    Case "ST_STEP_ORDER"
        GetFieldNumberinDB_byName = 5
    Case "ST_STEP_NAME"
        GetFieldNumberinDB_byName = 6
    Case "ST_DESCRIPTION"
        GetFieldNumberinDB_byName = 7
    Case "ST_EXPECTED"
        GetFieldNumberinDB_byName = 8
    Case "ST_ACTUAL"
        GetFieldNumberinDB_byName = 9
    Case "ST_STATUS"
        GetFieldNumberinDB_byName = 10
    Case "ST_EXECUTION_DATE"
        GetFieldNumberinDB_byName = 11
    Case "ST_EXECUTION_TIME"
        GetFieldNumberinDB_byName = 12
    Case "ST_LEVEL"
        GetFieldNumberinDB_byName = 13
End Select
End Function

Private Function IncorrectHeaderDetails() As Boolean
On Error Resume Next
    If flxImport.TextMatrix(0, 0) <> "Step ID" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 1) <> "Test Set Folder" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 2) <> "Test Set" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 3) <> "Test Instance" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 4) <> "Run" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 5) <> "Step Order" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 6) <> "Step Name" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 7) <> "Description" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 8) <> "Expected" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 9) <> "Actual" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 10) <> "Status" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 11) <> "Exec Date" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 12) <> "Exec Time" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 13) <> "Level" Then IncorrectHeaderDetails = True
If Err.Number <> 0 Then IncorrectHeaderDetails = False
End Function

Private Function IncorrectHeaderDetails_XML() As Boolean
On Error Resume Next
    If flxImport.TextMatrix(0, 0) <> "Step Number" Then IncorrectHeaderDetails_XML = True
    If flxImport.TextMatrix(0, 1) <> "Execution Time" Then IncorrectHeaderDetails_XML = True
    If flxImport.TextMatrix(0, 2) <> "Elapsed Time" Then IncorrectHeaderDetails_XML = True
    If flxImport.TextMatrix(0, 3) <> "Step Result" Then IncorrectHeaderDetails_XML = True
    If flxImport.TextMatrix(0, 4) <> "Step Summary" Then IncorrectHeaderDetails_XML = True
    If flxImport.TextMatrix(0, 5) <> "Component Name" Then IncorrectHeaderDetails_XML = True
    If flxImport.TextMatrix(0, 6) <> "Step Description" Then IncorrectHeaderDetails_XML = True
    If flxImport.TextMatrix(0, 7) <> "Image Path" Then IncorrectHeaderDetails_XML = True
If Err.Number <> 0 Then IncorrectHeaderDetails_XML = False
End Function

Private Function GetTestSetFolderPath(strID As String) As String
Dim Fact As TestSetFactory
Dim Obj As TestSet
Set Fact = QCConnection.TestSetFactory
Set Obj = Fact.Item(strID)
GetTestSetFolderPath = Obj.TestSetFolder.path
End Function

Private Sub txtFilter_DblClick()
txtFilter.Locked = False
End Sub

'########################### Add Run Steps to Test Instance ###########################
Private Sub AddRunSteps(tsTestID, tsTestName)
Dim tsttestsetFact As TSTestFactory
Dim tsttestset As tsTest
Dim tstRunFact As RunFactory
Dim tstRun As Run
Dim tstStepFact As StepFactory
Dim tstStep As Step, i, numErr
'Get the Test Factory
Err.Clear
Set tsttestsetFact = QCConnection.TSTestFactory
Set tsttestset = tsttestsetFact.Item(tsTestID)
Set tstRunFact = tsttestset.RunFactory
Set tstRun = tstRunFact.AddItem("Run_" & Format(flxImport.TextMatrix(1, 1), "m-d-h-m-s"))
tstRun.Status = RN_STATUS
tstRun.Post
mdiMain.pBar.Max = flxImport.Rows
    Set tstStepFact = tstRun.StepFactory
    Set tstStep = tstStepFact.AddItem(Null)
    tstStep.Field("ST_EXECUTION_DATE") = Format(flxImport.TextMatrix(1, 1), "dd/mmm/yyyy")
    tstStep.Field("ST_EXECUTION_TIME") = Format(flxImport.TextMatrix(1, 1), "hh:mm:ss")
    tstStep.Field("ST_STEP_NAME") = "Iteration 1"
    tstStep.Field("ST_DESCRIPTION") = ""
    tstStep.Field("ST_OBJ_ID") = "3"
    tstStep.Field("ST_LEVEL") = "3"
    tstStep.Status = RN_STATUS
    tstStep.Post
    tstStep.Field("ST_EXECUTION_DATE") = Format(flxImport.TextMatrix(1, 1), "dd/mmm/yyyy")
    tstStep.Field("ST_EXECUTION_TIME") = Format(flxImport.TextMatrix(1, 1), "hh:mm:ss")
    tstStep.Post
    
    Set tstStep = tstStepFact.AddItem(Null)
    tstStep.Field("ST_EXECUTION_DATE") = Format(flxImport.TextMatrix(1, 1), "dd/mmm/yyyy")
    tstStep.Field("ST_EXECUTION_TIME") = Format(flxImport.TextMatrix(1, 1), "hh:mm:ss")
    tstStep.Field("ST_STEP_NAME") = tsTestName
    tstStep.Field("ST_DESCRIPTION") = ""
    tstStep.Field("ST_OBJ_ID") = "2"
    tstStep.Field("ST_LEVEL") = "2"
    tstStep.Status = RN_STATUS
    tstStep.Post
    tstStep.Field("ST_EXECUTION_DATE") = Format(flxImport.TextMatrix(1, 1), "dd/mmm/yyyy")
    tstStep.Field("ST_EXECUTION_TIME") = Format(flxImport.TextMatrix(1, 1), "hh:mm:ss")
    tstStep.Post
    
For i = 2 To flxImport.Rows - 1
    Set tstStep = tstStepFact.AddItem(Null)
    tstStep.Field("ST_EXECUTION_DATE") = Format(flxImport.TextMatrix(i, 1), "dd/mmm/yyyy")
    tstStep.Field("ST_EXECUTION_TIME") = Format(flxImport.TextMatrix(i, 1), "hh:mm:ss")
    tstStep.Field("ST_STEP_NAME") = flxImport.TextMatrix(i, 5) & " " & flxImport.TextMatrix(i, 4)
    tstStep.Field("ST_DESCRIPTION") = flxImport.TextMatrix(i, 6)
    If Trim(flxImport.TextMatrix(i, 7)) <> "" Then tstStep.Field("ST_DESCRIPTION") = tstStep.Field("ST_DESCRIPTION") & " " & flxImport.TextMatrix(i, 7)
    tstStep.Field("ST_LEVEL") = ""
    tstStep.Field("ST_OBJ_ID") = ""
    If UCase(Trim(flxImport.TextMatrix(i, 3))) = "INFO" Or UCase(Trim(flxImport.TextMatrix(i, 3))) = "DONE" Then
        tstStep.Status = "Done"
    Else
        tstStep.Status = Trim(flxImport.TextMatrix(i, 3))
    End If
    tstStep.Post
    tstStep.Field("ST_EXECUTION_DATE") = Format(flxImport.TextMatrix(i, 1), "dd/mmm/yyyy")
    tstStep.Field("ST_EXECUTION_TIME") = Format(flxImport.TextMatrix(i, 1), "hh:mm:ss")
    tstStep.Post
    mdiMain.pBar.Value = i
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Updating Run Step " & i & " of " & flxImport.Rows - 1 & " (" & numErr & ") errors found"
    If Err.Number <> 0 Then
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Run Step (FAILED) " & Now & " " & flxImport.TextMatrix(i, 1) & " " & Err.Description
        numErr = numErr + 1
        stsBar.Panels(1).Picture = imgList_Sts.ListImages(3).Picture: stsBar.Panels(2).Text = "Updating Run Step " & i & " of " & flxImport.TextMatrix(i, 1) & " (" & numErr & ") errors found"
    Else
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Test Instance (PASSED) " & Now & " " & flxImport.TextMatrix(i, 1)
    End If
    Err.Clear
    On Error GoTo 0
Next
    tstRun.Field("RN_EXECUTION_DATE") = Format(flxImport.TextMatrix(1, 1), "dd/mmm/yyyy")
    tstRun.Field("RN_EXECUTION_TIME") = Format(flxImport.TextMatrix(1, 1), "hh:mm:ss")
    tstRun.Field("RN_DURATION") = DateDiff("s", flxImport.TextMatrix(1, 1), flxImport.TextMatrix(flxImport.Rows - 1, 1))
    tstRun.Field("RN_TESTER_NAME") = curUser
    tstRun.Field("RN_HOST") = RN_HOST
    tstRun.Post
    mdiMain.pBar.Value = mdiMain.pBar.Max
    FXGirl.EZPlay FXDataUploadCompleted
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = "Updated Run Step " & flxImport.Rows - 1 & " of " & flxImport.Rows - 1 & " (" & numErr & ") errors found"
    QCConnection.SendMail "user@companyemail.com", "", "[HPQC UPDATES] Updated Run Step  by " & curUser & " in " & curDomain & "-" & curProject, "Updated Run Step " & flxImport.Rows - 1 & " of " & flxImport.Rows - 1 & " (" & numErr & ") errors found" & "<br><br>" & "Source Data FileName: " & dlgOpenXML.filename, "", "HTML"
    QCConnection.SendMail curUser, "", "[HPQC UPDATES] Updated Run Step  by " & curUser & " in " & curDomain & "-" & curProject, "Updated Run Step " & flxImport.Rows - 1 & " of " & flxImport.Rows - 1 & " (" & numErr & ") errors found" & "<br><br>" & "Source Data FileName: " & dlgOpenXML.filename, "", "HTML"
    If numErr <> 0 Then
      Dim tmpFile As New clsFiles
      frmLogs.Caption = App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log"
      frmLogs.txtLogs.Text = tmpFile.ReadFromFile_FAILED(App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log")
      frmLogs.Show 1
    End If
End Sub
'########################### End Of Add Run Steps to Test Instance ###########################
