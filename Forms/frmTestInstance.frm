VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTestInstance 
   Caption         =   "Test Instance Module"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12675
   Icon            =   "frmTestInstance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   12675
   Tag             =   "Test Instance Module"
   WindowState     =   2  'Maximized
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
      Picture         =   "frmTestInstance.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Import step description and expected results from an excel file"
      Top             =   1800
      Width           =   2205
   End
   Begin VB.ListBox lstUpdateFields 
      Columns         =   4
      Height          =   960
      ItemData        =   "frmTestInstance.frx":1070
      Left            =   4560
      List            =   "frmTestInstance.frx":1072
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
         TabIndex        =   8
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
            Picture         =   "frmTestInstance.frx":1074
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestInstance.frx":1306
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestInstance.frx":1598
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
      Cols            =   28
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
            Picture         =   "frmTestInstance.frx":1826
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestInstance.frx":1F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestInstance.frx":264A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestInstance.frx":2D5C
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
            Picture         =   "frmTestInstance.frx":346E
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
            Picture         =   "frmTestInstance.frx":39BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestInstance.frx":3CA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestInstance.frx":41F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestInstance.frx":4743
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
Attribute VB_Name = "frmTestInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UpdateList()

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
Dim a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26, a27, last_a0, last_a1

    TimeSt = Format(Now, "mmm-dd-yyyy hhmmss") & "-"
    If chkCSV.Value = Checked Then
      AllF = InputBox("Enter file name", "File name", "[Test Instance] ")
    Else
      AllF = "[Test Instance] "
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
      objCommand.CommandText = "SELECT TC_TESTCYCL_ID, CF_ITEM_NAME, CY_CYCLE, CY_CYCLE_ID, CY_STATUS, CY_USER_02, CY_USER_03 ,CY_USER_05, CY_USER_06, TS_NAME, TC_PLAN_SCHEDULING_DATE, TC_USER_TEMPLATE_03, TC_USER_TEMPLATE_04, TC_USER_TEMPLATE_05, TC_USER_TEMPLATE_09, TC_TESTER_NAME, CY_USER_04, TC_USER_TEMPLATE_02, TC_STATUS, TS_USER_TEMPLATE_07, TC_USER_TEMPLATE_10, TC_USER_TEMPLATE_01, TC_USER_TEMPLATE_13, TC_USER_02, TC_USER_01," & _
                             "TC_ITERATIONS, TC_USER_26, TC_TEST_ORDER FROM TEST, TESTCYCL, CYCLE, CYCL_FOLD WHERE TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & _
                              strPath & " AND " & Trim(txtFilter.Text) & " ORDER BY CY_CYCLE, TC_TEST_ORDER "
    Else
      objCommand.CommandText = "SELECT TC_TESTCYCL_ID, CF_ITEM_NAME, CY_CYCLE, CY_CYCLE_ID, CY_STATUS, CY_USER_02, CY_USER_03 ,CY_USER_05, CY_USER_06, TS_NAME, TC_PLAN_SCHEDULING_DATE, TC_USER_TEMPLATE_03, TC_USER_TEMPLATE_04, TC_USER_TEMPLATE_05, TC_USER_TEMPLATE_09, TC_TESTER_NAME, CY_USER_04, TC_USER_TEMPLATE_02, TC_STATUS, TS_USER_TEMPLATE_07, TC_USER_TEMPLATE_10, TC_USER_TEMPLATE_01, TC_USER_TEMPLATE_13, TC_USER_02, TC_USER_01," & _
                             "TC_ITERATIONS, TC_USER_26, TC_TEST_ORDER FROM TEST, TESTCYCL, CYCLE, CYCL_FOLD WHERE TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & _
                              strPath & " ORDER BY CY_CYCLE, TC_TEST_ORDER "
    End If
    
    Debug.Print Me.Caption & "-" & objCommand.CommandText
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[SQL] " & Now & " " & objCommand.CommandText
    Set rs = objCommand.Execute                                                                                                                                                                                                                                                 'HERE!!!!!! <<<-------------
    If rs.RecordCount > 10000 And chkCSV.Value = Unchecked Then '***
        MsgBox "The records found exceeds 2500 records. It will be automatically generated as a CSV file.", vbOKOnly
        chkCSV.Value = Checked
        Exit Sub
        GenerateOutput
    End If '***
    AllScript = """" & "Test Instance ID" & """" & "," & """" & "Test Set Folder" & """" & "," & """" & "Test Set" & """" & "," & """" & "SV Status" & """" & "," & """" & "SV: Plan Scripting Start Date" & """" & "," & """" & "SV: Plan Scripting End Date" & """" & "," & """" & "SV: Plan Exec. Start Date" & """" & "," & """" & "SV: Plan Exec. End Date" & """" & "," & """" & "Test Name" & """" & "," & """" & "Planned Exec Start Date" & """" & "," & """" & "Planned Exec End Date" & """" & "," & """" & "Actual Exec Start Date" & """" & "," & """" & "Actual Exec End Date" & """" & "," & """" & "Data Scripter" & """" & "," & """" & "Responsible Tester" & """" & "," & """" & "Executed by" & """" & "," & """" & "QA Reviewer" & """" & "," & """" & "Test Lab Status" & """" & "," & """" & "Test Script Status" & """" & "," & """" & "Data Scripting Status" & """" & "," & """" & "QA Review Status" & """" & "," & """" & "Assigned Group" & """" & "," & """" & "Check for Dependency" & """" & "," & """" & _
                "Test Lab: Description" & """" & "," & """" & "Iterations" & """" & "," & """" & "Output Parameter" & """" & "," & """" & "Order" & """" & "," & """" & "Action" & """"
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
                  flxImport.TextMatrix(k, 0) = rs.FieldValue("TC_TESTCYCL_ID")
                  If rs.FieldValue("TC_TESTCYCL_ID") <> flxImport.TextMatrix(k - 1, 0) Then
                      flxImport.TextMatrix(k, 1) = GetTestSetFolderPath(rs.FieldValue("CY_CYCLE_ID")) 'rs.FieldValue("CF_ITEM_NAME")
                  Else
                      flxImport.TextMatrix(k, 1) = flxImport.TextMatrix(k - 1, 1)
                  End If
                  flxImport.TextMatrix(k, 2) = ReplaceAllEnter(rs.FieldValue("CY_CYCLE"))
                  flxImport.TextMatrix(k, 3) = rs.FieldValue("CY_STATUS")
                  flxImport.TextMatrix(k, 4) = rs.FieldValue("CY_USER_02")
                  flxImport.TextMatrix(k, 5) = rs.FieldValue("CY_USER_03")
                  flxImport.TextMatrix(k, 6) = rs.FieldValue("CY_USER_05")
                  flxImport.TextMatrix(k, 7) = rs.FieldValue("CY_USER_06")
                  flxImport.TextMatrix(k, 8) = rs.FieldValue("TS_NAME")
                  flxImport.TextMatrix(k, 9) = Format(rs.FieldValue("TC_PLAN_SCHEDULING_DATE"), "dd-mmm-yyyy")
                  flxImport.TextMatrix(k, 10) = Format(rs.FieldValue("TC_USER_TEMPLATE_03"), "dd-mmm-yyyy")
                  flxImport.TextMatrix(k, 11) = rs.FieldValue("TC_USER_TEMPLATE_04")
                  flxImport.TextMatrix(k, 12) = rs.FieldValue("TC_USER_TEMPLATE_05")
                  flxImport.TextMatrix(k, 13) = rs.FieldValue("TC_USER_TEMPLATE_09")
                  flxImport.TextMatrix(k, 14) = rs.FieldValue("TC_TESTER_NAME")
                  flxImport.TextMatrix(k, 15) = rs.FieldValue("CY_USER_04")
                  flxImport.TextMatrix(k, 16) = rs.FieldValue("TC_USER_TEMPLATE_02")
                  flxImport.TextMatrix(k, 17) = rs.FieldValue("TC_STATUS")
                  flxImport.TextMatrix(k, 18) = rs.FieldValue("TS_USER_TEMPLATE_07")
                  flxImport.TextMatrix(k, 19) = rs.FieldValue("TC_USER_TEMPLATE_10")
                  flxImport.TextMatrix(k, 20) = rs.FieldValue("TC_USER_TEMPLATE_01")
                  flxImport.TextMatrix(k, 21) = rs.FieldValue("TC_USER_TEMPLATE_13")
                  flxImport.TextMatrix(k, 22) = rs.FieldValue("TC_USER_02")
                  flxImport.TextMatrix(k, 23) = rs.FieldValue("TC_USER_01")
                  flxImport.TextMatrix(k, 24) = rs.FieldValue("TC_ITERATIONS")
                  flxImport.TextMatrix(k, 25) = ReverseCleanHTML(rs.FieldValue("TC_USER_26"))
                  flxImport.TextMatrix(k, 26) = rs.FieldValue("TC_TEST_ORDER")
                  flxImport.TextMatrix(k, 27) = ""
                Else
                  a0 = rs.FieldValue("TC_TESTCYCL_ID")
                  If a0 <> last_a0 Then
                    a1 = GetTestSetFolderPath(rs.FieldValue("CY_CYCLE_ID"))
                    last_a1 = a1
                    last_a0 = a0
                  Else
                    a1 = last_a1
                  End If
                  a2 = ReplaceAllEnter(rs.FieldValue("CY_CYCLE"))
                  a3 = rs.FieldValue("CY_STATUS")
                  a4 = rs.FieldValue("CY_USER_02")
                  a5 = rs.FieldValue("CY_USER_03")
                  a6 = rs.FieldValue("CY_USER_05")
                  a7 = rs.FieldValue("CY_USER_06")
                  a8 = rs.FieldValue("TS_NAME")
                  a9 = Format(rs.FieldValue("TC_PLAN_SCHEDULING_DATE"), "dd-mmm-yyyy")
                  a10 = Format(rs.FieldValue("TC_USER_TEMPLATE_03"), "dd-mmm-yyyy")
                  a11 = rs.FieldValue("TC_USER_TEMPLATE_04")
                  a12 = rs.FieldValue("TC_USER_TEMPLATE_05")
                  a13 = rs.FieldValue("TC_USER_TEMPLATE_09")
                  a14 = rs.FieldValue("TC_TESTER_NAME")
                  a15 = rs.FieldValue("CY_USER_04")
                  a16 = rs.FieldValue("TC_USER_TEMPLATE_02")
                  a17 = rs.FieldValue("TC_STATUS")
                  a18 = rs.FieldValue("TS_USER_TEMPLATE_07")
                  a19 = rs.FieldValue("TC_USER_TEMPLATE_10")
                  a20 = rs.FieldValue("TC_USER_TEMPLATE_01")
                  a21 = rs.FieldValue("TC_USER_TEMPLATE_13")
                  a22 = rs.FieldValue("TC_USER_02")
                  a23 = rs.FieldValue("TC_USER_01")
                  a24 = rs.FieldValue("TC_ITERATIONS")
                  a25 = ReverseCleanHTML(rs.FieldValue("TC_USER_26"))
                  a26 = rs.FieldValue("TC_TEST_ORDER")
                  a27 = ""
                  If Trim(AllScript) <> "" Then
                        AllScript = AllScript & vbCrLf & """" & a0 & """" & "," & """" & a1 & """" & "," & """" & a2 & """" & "," & """" & a3 & """" & "," & """" & a4 & """" & "," & """" & a5 & """" & "," & """" & a6 & """" & "," & """" & a7 & """" & "," & """" & a8 & """" & "," & """" & a9 & """" & "," & """" & a10 & """" & "," & """" & a11 & """" & "," & """" & a12 & """" & "," & """" & a13 & """" & "," & """" & a14 & """" & "," & """" & a15 & """" & "," & """" & a16 & """" & "," & """" & a17 & """" & "," & """" & a18 & """" & "," & """" & a19 & """" & "," & """" & a20 & """" & "," & """" & a21 & """" & "," & """" & a22 & """" & "," & """" & a23 & """" & "," & """" & a24 & """" & "," & """" & a25 & """" & "," & """" & a26 & """" & "," & """" & a27 & """"
                  Else
                        AllScript = AllScript & """" & a0 & """" & "," & """" & a1 & """" & "," & """" & a2 & """" & "," & """" & a3 & """" & "," & """" & a4 & """" & "," & """" & a5 & """" & "," & """" & a6 & """" & "," & """" & a7 & """" & "," & """" & a8 & """" & "," & """" & a9 & """" & "," & """" & a10 & """" & "," & """" & a11 & """" & "," & """" & a12 & """" & "," & """" & a13 & """" & "," & """" & a14 & """" & "," & """" & a15 & """" & "," & """" & a16 & """" & "," & """" & a17 & """" & "," & """" & a18 & """" & "," & """" & a19 & """" & "," & """" & a20 & """" & "," & """" & a21 & """" & "," & """" & a22 & """" & "," & """" & a23 & """" & "," & """" & a24 & """" & "," & """" & a25 & """" & "," & """" & a26 & """" & "," & """" & a27 & """"
                  End If
                End If '***
                If chkCSV.Value = Checked Then '***
                    If i Mod 500 = 0 Then
                        FileAppend App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv", AllScript
                        AllScript = ""
                    End If
                End If '***
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
stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = mdiMain.pBar.Max & " record(s) generated"
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
         If InStr(1, GetCommentText(.Range("A1")), "Test Instance") = 0 Then
            MsgBox "Import file is invalid. Please use only sheets generated by the SuperQC"
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
         For i = 1 To 28
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
        .Cols = 28
        .RowSel = lastrow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
        .ColSel = 28 - 1
    End With
    
     With flxImport
        .Clear
        .Redraw = False     'Dont draw until the end, so we avoid that flash
        .row = 0            'Paste from first cell
        .col = 0
        .Rows = lastrow
        .Cols = 28
        .RowSel = lastrow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
        .ColSel = 28 - 1
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
        If MsgBox("Are you sure you want to mass update " & flxImport.Rows - 1 & " record(s) of Test Instances?", vbYesNo) = vbYes Then
            Randomize: tmpR = CInt(Rnd(1000) * 10000)
            If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
                GetUpdateList
                Upload_Test_Instance
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
    QCTree.Nodes.Add , , "Root", "Root", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = 0 ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    
    lstUpdateFields.Clear
'    lstUpdateFields.AddItem "Test Instance ID"
'    lstUpdateFields.AddItem "Test Set Folder"
'    lstUpdateFields.AddItem "Test Set"
'    lstUpdateFields.AddItem "SV Status"
'    lstUpdateFields.AddItem "SV: Plan Scripting Start Date"
'    lstUpdateFields.AddItem "SV: Plan Scripting End Date"
'    lstUpdateFields.AddItem "SV: Plan Exec. Start Date"
'    lstUpdateFields.AddItem "SV: Plan Exec. End Date"
'    lstUpdateFields.AddItem "Test Name"
    lstUpdateFields.AddItem "Planned Exec Start Date"
    lstUpdateFields.AddItem "Planned Exec End Date"
    lstUpdateFields.AddItem "Actual Exec Start Date"
    lstUpdateFields.AddItem "Actual Exec End Date"
    lstUpdateFields.AddItem "Data Scripter"
    lstUpdateFields.AddItem "Responsible Tester"
'    lstUpdateFields.AddItem "Executed by"
    lstUpdateFields.AddItem "QA Reviewer"
    lstUpdateFields.AddItem "Test Lab Status"
'    lstUpdateFields.AddItem "Test Script Status"
    lstUpdateFields.AddItem "Data Scripting Status"
    lstUpdateFields.AddItem "QA Review Status"
   lstUpdateFields.AddItem "Assigned Group"
    lstUpdateFields.AddItem "Check for Dependency"
    lstUpdateFields.AddItem "Test Lab: Description"
    lstUpdateFields.AddItem "Iterations"
    lstUpdateFields.AddItem "Output Parameter"
    lstUpdateFields.AddItem "Order"
    lstUpdateFields.AddItem "Action"
     Me.Caption = Me.Tag
     txtFilter.Text = ""
     txtFilter.Locked = True
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.TextMatrix(0, 0) = "Test Instance ID"
flxImport.TextMatrix(0, 1) = "Test Set Folder"
flxImport.TextMatrix(0, 2) = "Test Set"
flxImport.TextMatrix(0, 3) = "SV Status"
flxImport.TextMatrix(0, 4) = "SV: Plan Scripting Start Date"
flxImport.TextMatrix(0, 5) = "SV: Plan Scripting End Date"
flxImport.TextMatrix(0, 6) = "SV: Plan Exec. Start Date"
flxImport.TextMatrix(0, 7) = "SV: Plan Exec. End Date"
flxImport.TextMatrix(0, 8) = "Test Name"
flxImport.TextMatrix(0, 9) = "Planned Exec Start Date"
flxImport.TextMatrix(0, 10) = "Planned Exec End Date"
flxImport.TextMatrix(0, 11) = "Actual Exec Start Date"
flxImport.TextMatrix(0, 12) = "Actual Exec End Date"
flxImport.TextMatrix(0, 13) = "Data Scripter"
flxImport.TextMatrix(0, 14) = "Responsible Tester"
flxImport.TextMatrix(0, 15) = "Executed by"
flxImport.TextMatrix(0, 16) = "QA Reviewer"
flxImport.TextMatrix(0, 17) = "Test Lab Status"
flxImport.TextMatrix(0, 18) = "Test Script Status"
flxImport.TextMatrix(0, 19) = "Data Scripting Status"
flxImport.TextMatrix(0, 20) = "QA Review Status"
flxImport.TextMatrix(0, 21) = "Assigned Group"
flxImport.TextMatrix(0, 22) = "Check for Dependency"
flxImport.TextMatrix(0, 23) = "Test Lab: Description"
flxImport.TextMatrix(0, 24) = "Iterations"
flxImport.TextMatrix(0, 25) = "Output Parameter"
flxImport.TextMatrix(0, 26) = "Order"
flxImport.TextMatrix(0, 27) = "Action"
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
  curTab = "TEST_INSTANCE01"
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
    xlObject.Sheets(curTab).Range("A:AB").Select
        
    xlObject.Sheets(curTab).Range("A:AB").Borders(xlDiagonalDown).LineStyle = xlNone
    xlObject.Sheets(curTab).Range("A:AB").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlObject.Sheets(curTab).Range("A:AB").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:AB").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:AB").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:AB").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:AB").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:AB").Borders(xlInsideHorizontal)
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
    xlObject.Sheets(curTab).Range("A:AB").Select
    xlObject.Sheets(curTab).Range("A:AB").EntireColumn.AutoFit
    xlObject.Sheets(curTab).Range("A1").Select
    
    xlObject.Sheets(curTab).Range("A1").AddComment
    xlObject.Sheets(curTab).Range("A1").Comment.Visible = False
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & mdiMain.Caption & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""
    
    xlObject.Sheets(curTab).Range(GetEditableFields).Interior.ColorIndex = 35
    xlObject.Sheets(curTab).Protection.AllowEditRanges.Add Title:="Range1", Range:=xlObject.Sheets(curTab).Range(GetEditableFields)
    xlObject.Sheets(curTab).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
  xlObject.Workbooks(1).SaveAs "TEST_INSTANCE01-" & CleanTheString(QCTree.SelectedItem.Text) & "-" & Format(Now, "mmddyyyy HHMM AMPM")
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
   GetEditableFields = "IV:IV"
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
        If Left(UCase(Trim(UpdateList(i))), 8) = "[DELETE]" Then
            Set tsTestF = QCConnection.TSTestFactory
            tmp = Trim(Replace(UpdateList(i), "[DELETE]", ""))
            tsTestF.RemoveItem tmp
            If Err.Number <> 0 Then
                FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Deleting Test Instance (FAILED) " & Now & " " & objCommand.CommandText & " " & Err.Description
                numErr = numErr + 1
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(3).Picture: stsBar.Panels(2).Text = "Deleting Test Instance " & i & " of " & UBound(UpdateList) + 1 & " (" & numErr & ") errors found"
            Else
                FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Deleting Test Instance (PASSED) " & Now & " " & objCommand.CommandText
            End If
            Err.Clear
            On Error GoTo 0
            Set objCommand = Nothing
            mdiMain.pBar.Value = i + 1
        ElseIf Left(UCase(Trim(UpdateList(i))), 10) = "[RESTRING]" Then
            Set tsTestF = QCConnection.TSTestFactory
            tmp = Trim(Replace(UpdateList(i), "[RESTRING]", ""))
            Set tsTest = tsTestF.Item(tmp) ' <<----
            tmpTestID = tsTest.Field("TC_TEST_ID")
            tmpTestSetID = tsTest.Field("TC_CYCLE_ID")
            tmpOrder = tsTest.Field("TC_TEST_ORDER")
            tmpStartDate = tsTest.Field("TC_PLAN_SCHEDULING_DATE")
            tmpEndDate = tsTest.Field("TC_USER_TEMPLATE_03")
            tmpScripter = tsTest.Field("TC_TESTER_NAME")
            tmpQA = tsTest.Field("TC_USER_TEMPLATE_02")
            tmpGroup = tsTest.Field("TC_USER_TEMPLATE_13")
            tmpDataScript = tsTest.Field("TC_USER_TEMPLATE_10")
            tsTestF.RemoveItem tmp
                        
            Set TestSetFact = QCConnection.TestSetFactory
            Set mytestset = TestSetFact.Item(tmpTestSetID)
            Set tfact = QCConnection.TestFactory
            Set mytest = tfact.Item(tmpTestID)
            Set tsttestsetFact = mytestset.TSTestFactory
            Set mytsttestset = tsttestsetFact.AddItem(mytest.ID)
            mytsttestset.Field("TC_PLAN_SCHEDULING_DATE") = tmpStartDate
            mytsttestset.Field("TC_USER_TEMPLATE_03") = tmpEndDate
            mytsttestset.Field("TC_TESTER_NAME") = tmpScripter
            mytsttestset.Field("TC_USER_TEMPLATE_02") = tmpQA
            mytsttestset.Field("TC_USER_TEMPLATE_13") = tmpGroup
            mytsttestset.Field("TC_TEST_ORDER") = tmpOrder
            mytsttestset.Field("TC_USER_TEMPLATE_10") = tmpDataScript
            mytsttestset.Post
            
            Set tfact = Nothing
            Set mytest = Nothing
            Set mytsttestset = Nothing
            If Err.Number <> 0 Then
                FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Deleting Test Instance (FAILED) " & Now & " " & objCommand.CommandText & " " & Err.Description
                numErr = numErr + 1
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(3).Picture: stsBar.Panels(2).Text = "Deleting Test Instance " & i & " of " & UBound(UpdateList) + 1 & " (" & numErr & ") errors found"
            Else
                FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Deleting Test Instance (PASSED) " & Now & " " & objCommand.CommandText
            End If
            Err.Clear
            On Error GoTo 0
            Set objCommand = Nothing
            mdiMain.pBar.Value = i + 1
        Else
            Set objCommand = QCConnection.Command
            objCommand.CommandText = UpdateList(i)
            'InputBox "", "", UpdateList(i)
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
    If flxImport.TextMatrix(i, 27) = "" Then
        X = X + 1
        ReDim Preserve UpdateList(X)
        For j = 0 To lstUpdateFields.ListCount - 1
            If lstUpdateFields.Selected(j) = True Then
                curField = GetFieldNameinDB(lstUpdateFields.List(j))
                If curField = "TC_PLAN_SCHEDULING_DATE" Or curField = "TC_USER_TEMPLATE_03" Then
                    tmpColVal = tmpColVal & curField & " = '" & Format(flxImport.TextMatrix(i, GetFieldNumberinDB_byName(curField)), "dd-mmm-yyyy") & "', "
                Else
                    tmpColVal = tmpColVal & curField & " = '" & flxImport.TextMatrix(i, GetFieldNumberinDB_byName(curField)) & "', "
                End If
            End If
        Next
        tmpColVal = Left(tmpColVal, Len(tmpColVal) - 2)
        UpdateList(X) = "UPDATE TESTCYCL SET " & tmpColVal & " WHERE TC_TESTCYCL_ID = " & flxImport.TextMatrix(i, 0)
        tmpColVal = ""
    ElseIf UCase(Trim(flxImport.TextMatrix(i, 27))) = "DELETE" Then
        X = X + 1
        ReDim Preserve UpdateList(X)
        UpdateList(X) = "[DELETE]" & flxImport.TextMatrix(i, 0)
    ElseIf UCase(Trim(flxImport.TextMatrix(i, 27))) = "RESTRING" Then
        X = X + 1
        ReDim Preserve UpdateList(X)
        UpdateList(X) = "[RESTRING]" & flxImport.TextMatrix(i, 0)
    End If
Next
End Function
'========<><><><><><><> DATES SHOULD BE ENCLOSED WITH # # OR ???

Private Function GetFieldNameinDB(X As String)
Select Case X
    Case "Test Instance ID"
        GetFieldNameinDB = "TC_TESTCYCL_ID"
    Case "Test Set Folder"
        GetFieldNameinDB = "CF_ITEM_NAME"
    Case "Test Set"
        GetFieldNameinDB = "CY_CYCLE"
    Case "SV Status"
        GetFieldNameinDB = "CY_STATUS"
    Case "SV: Plan Scripting Start Date"
        GetFieldNameinDB = "CY_USER_02"
    Case "SV: Plan Scripting End Date"
        GetFieldNameinDB = "CY_USER_03"
    Case "SV: Plan Exec. Start Date"
        GetFieldNameinDB = "CY_USER_05"
    Case "SV: Plan Exec. End Date"
        GetFieldNameinDB = "CY_USER_06"
    Case "Test Name"
        GetFieldNameinDB = "TS_NAME"
    Case "Planned Exec Start Date"
        GetFieldNameinDB = "TC_PLAN_SCHEDULING_DATE"
    Case "Planned Exec End Date"
        GetFieldNameinDB = "TC_USER_TEMPLATE_03"
    Case "Actual Exec Start Date"
        GetFieldNameinDB = "TC_USER_TEMPLATE_04"
    Case "Actual Exec End Date"
        GetFieldNameinDB = "TC_USER_TEMPLATE_05"
    Case "Data Scripter"
        GetFieldNameinDB = "TC_USER_TEMPLATE_09"
    Case "Responsible Tester"
        GetFieldNameinDB = "TC_TESTER_NAME"
    Case "Executed by"
        GetFieldNameinDB = "CY_USER_04"
    Case "QA Reviewer"
        GetFieldNameinDB = "TC_USER_TEMPLATE_02"
    Case "Test Lab Status"
        GetFieldNameinDB = "TC_STATUS"
    Case "Test Script Status"
        GetFieldNameinDB = "TS_USER_TEMPLATE_07"
    Case "Data Scripting Status"
        GetFieldNameinDB = "TC_USER_TEMPLATE_10"
    Case "QA Review Status"
        GetFieldNameinDB = "TC_USER_TEMPLATE_01"
    Case "Assigned Group"
        GetFieldNameinDB = "TC_USER_TEMPLATE_13"
    Case "Check for Dependency"
        GetFieldNameinDB = "TC_USER_02"
    Case "Test Lab: Description"
        GetFieldNameinDB = "TC_USER_01"
    Case "Iterations"
        GetFieldNameinDB = "TC_ITERATIONS"
    Case "Output Parameter"
        GetFieldNameinDB = "TC_USER_26"
    Case "Order"
        GetFieldNameinDB = "TC_TEST_ORDER"
    Case "Action"
        GetFieldNameinDB = "Action"
End Select
End Function

Private Function GetFieldNumberinDB(X As Integer)
Select Case X
    Case 0
        GetFieldNumberinDB = "TC_TESTCYCL_ID"
    Case 1
        GetFieldNumberinDB = "CF_ITEM_NAME"
    Case 2
        GetFieldNumberinDB = "CY_CYCLE"
    Case 3
        GetFieldNumberinDB = "CY_STATUS"
    Case 4
        GetFieldNumberinDB = "CY_USER_02"
    Case 5
        GetFieldNumberinDB = "CY_USER_03"
    Case 6
        GetFieldNumberinDB = "CY_USER_05"
    Case 7
        GetFieldNumberinDB = "CY_USER_06"
    Case 8
        GetFieldNumberinDB = "TS_NAME"
    Case 9
        GetFieldNumberinDB = "TC_PLAN_SCHEDULING_DATE"
    Case 10
        GetFieldNumberinDB = "TC_USER_TEMPLATE_03"
    Case 11
        GetFieldNumberinDB = "TC_USER_TEMPLATE_04"
    Case 12
        GetFieldNumberinDB = "TC_USER_TEMPLATE_05"
    Case 13
        GetFieldNumberinDB = "TC_USER_TEMPLATE_09"
    Case 14
        GetFieldNumberinDB = "TC_TESTER_NAME"
    Case 15
        GetFieldNumberinDB = "CY_USER_04"
    Case 16
        GetFieldNumberinDB = "TC_USER_TEMPLATE_02"
    Case 17
        GetFieldNumberinDB = "TC_STATUS"
    Case 18
        GetFieldNumberinDB = "TS_USER_TEMPLATE_07"
    Case 19
        GetFieldNumberinDB = "TC_USER_TEMPLATE_10"
    Case 20
        GetFieldNumberinDB = "TC_USER_TEMPLATE_01"
    Case 21
        GetFieldNumberinDB = "TC_USER_TEMPLATE_13"
    Case 22
        GetFieldNumberinDB = "TC_USER_02"
    Case 23
        GetFieldNumberinDB = "TC_USER_01"
    Case 24
        GetFieldNumberinDB = "TC_ITERATIONS"
    Case 25
        GetFieldNumberinDB = "TC_USER_26"
    Case 26
        GetFieldNumberinDB = "TC_TEST_ORDER"
    Case 27
        GetFieldNumberinDB = "Action"
End Select
End Function

Private Function GetFieldNumberinDB_byName(X As String)
Select Case X
    Case "Test Instance ID"
        GetFieldNumberinDB_byName = 0
    Case "Test Set Folder"
        GetFieldNumberinDB_byName = 1
    Case "Test Set"
        GetFieldNumberinDB_byName = 2
    Case "SV Status"
        GetFieldNumberinDB_byName = 3
    Case "SV: Plan Scripting Start Date"
        GetFieldNumberinDB_byName = 4
    Case "SV: Plan Scripting End Date"
        GetFieldNumberinDB_byName = 5
    Case "SV: Plan Exec. Start Date"
        GetFieldNumberinDB_byName = 6
    Case "SV: Plan Exec. End Date"
        GetFieldNumberinDB_byName = 7
    Case "Test Name"
        GetFieldNumberinDB_byName = 8
    Case "Planned Exec Start Date"
        GetFieldNumberinDB_byName = 9
    Case "Planned Exec End Date"
        GetFieldNumberinDB_byName = 10
    Case "Actual Exec Start Date"
        GetFieldNumberinDB_byName = 11
    Case "Actual Exec End Date"
        GetFieldNumberinDB_byName = 12
    Case "Data Scripter"
        GetFieldNumberinDB_byName = 13
    Case "Responsible Tester"
        GetFieldNumberinDB_byName = 14
    Case "Executed by"
        GetFieldNumberinDB_byName = 15
    Case "QA Reviewer"
        GetFieldNumberinDB_byName = 16
    Case "Test Lab Status"
        GetFieldNumberinDB_byName = 17
    Case "Test Script Status"
        GetFieldNumberinDB_byName = 18
    Case "Data Scripting Status"
        GetFieldNumberinDB_byName = 19
    Case "QA Review Status"
        GetFieldNumberinDB_byName = 20
    Case "Assigned Group"
        GetFieldNumberinDB_byName = 21
    Case "Check for Dependency"
        GetFieldNumberinDB_byName = 22
    Case "Test Lab: Description"
        GetFieldNumberinDB_byName = 23
    Case "Iterations"
        GetFieldNumberinDB_byName = 24
    Case "Output Parameter"
        GetFieldNumberinDB_byName = 25
    Case "Order"
        GetFieldNumberinDB_byName = 26
    Case "Action"
        GetFieldNumberinDB_byName = 27
        
    Case "TC_TESTCYCL_ID"
        GetFieldNumberinDB_byName = 0
    Case "CF_ITEM_NAME"
        GetFieldNumberinDB_byName = 1
    Case "CY_CYCLE"
        GetFieldNumberinDB_byName = 2
    Case "CY_STATUS"
        GetFieldNumberinDB_byName = 3
    Case "CY_USER_02"
        GetFieldNumberinDB_byName = 4
    Case "CY_USER_03"
        GetFieldNumberinDB_byName = 5
    Case "CY_USER_05"
        GetFieldNumberinDB_byName = 6
    Case "CY_USER_06"
        GetFieldNumberinDB_byName = 7
    Case "TS_NAME"
        GetFieldNumberinDB_byName = 8
    Case "TC_PLAN_SCHEDULING_DATE"
        GetFieldNumberinDB_byName = 9
    Case "TC_USER_TEMPLATE_03"
        GetFieldNumberinDB_byName = 10
    Case "TC_USER_TEMPLATE_04"
        GetFieldNumberinDB_byName = 11
    Case "TC_USER_TEMPLATE_05"
        GetFieldNumberinDB_byName = 12
    Case "TC_USER_TEMPLATE_09"
        GetFieldNumberinDB_byName = 13
    Case "TC_TESTER_NAME"
        GetFieldNumberinDB_byName = 14
    Case "CY_USER_04"
        GetFieldNumberinDB_byName = 15
    Case "TC_USER_TEMPLATE_02"
        GetFieldNumberinDB_byName = 16
    Case "TC_STATUS"
        GetFieldNumberinDB_byName = 17
    Case "TS_USER_TEMPLATE_07"
        GetFieldNumberinDB_byName = 18
    Case "TC_USER_TEMPLATE_10"
        GetFieldNumberinDB_byName = 19
    Case "TC_USER_TEMPLATE_01"
        GetFieldNumberinDB_byName = 20
    Case "TC_USER_TEMPLATE_13"
        GetFieldNumberinDB_byName = 21
    Case "TC_USER_02"
        GetFieldNumberinDB_byName = 22
    Case "TC_USER_01"
        GetFieldNumberinDB_byName = 23
    Case "TC_ITERATIONS"
        GetFieldNumberinDB_byName = 24
    Case "TC_USER_26"
        GetFieldNumberinDB_byName = 25
    Case "TC_TEST_ORDER"
        GetFieldNumberinDB_byName = 26
    Case "Action"
        GetFieldNumberinDB_byName = 27
End Select
End Function

Private Function IncorrectHeaderDetails() As Boolean
    If flxImport.TextMatrix(0, 0) <> "Test Instance ID" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 1) <> "Test Set Folder" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 2) <> "Test Set" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 3) <> "SV Status" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 4) <> "SV: Plan Scripting Start Date" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 5) <> "SV: Plan Scripting End Date" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 6) <> "SV: Plan Exec. Start Date" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 7) <> "SV: Plan Exec. End Date" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 8) <> "Test Name" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 9) <> "Planned Exec Start Date" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 10) <> "Planned Exec End Date" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 11) <> "Actual Exec Start Date" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 12) <> "Actual Exec End Date" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 13) <> "Data Scripter" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 14) <> "Responsible Tester" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 15) <> "Executed by" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 16) <> "QA Reviewer" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 17) <> "Test Lab Status" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 18) <> "Test Script Status" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 19) <> "Data Scripting Status" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 20) <> "QA Review Status" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 21) <> "Assigned Group" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 22) <> "Check for Dependency" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 23) <> "Test Lab: Description" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 24) <> "Iterations" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 25) <> "Output Parameter" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 26) <> "Order" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 27) <> "Action" Then IncorrectHeaderDetails = True
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
