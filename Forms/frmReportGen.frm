VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReportGen 
   Caption         =   "SuperQC Quicky Quick Extreme Reporter"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13230
   Icon            =   "frmReportGen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   13230
   Tag             =   "SuperQC Quicky Quick Extreme Reporter"
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFilter 
      Height          =   375
      Left            =   9000
      TabIndex        =   15
      Top             =   900
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   60
      TabIndex        =   7
      Top             =   6360
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   5953
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabMaxWidth     =   3528
      TabCaption(0)   =   "SQL"
      TabPicture(0)   =   "frmReportGen.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtSQL"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Order"
      TabPicture(1)   =   "frmReportGen.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstExtract"
      Tab(1).Control(1)=   "UpDown"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Result"
      TabPicture(2)   =   "frmReportGen.frx":0902
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "flxImport"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtSQL 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   420
         Width           =   12675
      End
      Begin VB.ListBox lstExtract 
         Height          =   2595
         Left            =   -74880
         TabIndex        =   10
         Top             =   420
         Width           =   8775
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   795
         Left            =   -66060
         TabIndex        =   8
         Top             =   420
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1402
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid flxImport 
         Height          =   1995
         Left            =   120
         TabIndex        =   9
         Top             =   420
         Width           =   12705
         _ExtentX        =   22410
         _ExtentY        =   3519
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
   End
   Begin VB.CommandButton cmdModule 
      Height          =   435
      Left            =   4440
      Picture         =   "frmReportGen.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.ListBox lstFields 
      Height          =   4335
      Left            =   9000
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1620
      Width           =   3975
   End
   Begin VB.ListBox lstTables 
      Height          =   4740
      Left            =   5160
      TabIndex        =   4
      Top             =   1380
      Width           =   3735
   End
   Begin VB.ComboBox cmbModule 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmReportGen.frx":0FB7
      Left            =   60
      List            =   "frmReportGen.frx":0FB9
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   4395
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13230
      _ExtentX        =   23336
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
            Key             =   "cmdSave"
            Object.ToolTipText     =   "Save to Report List"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin VB.CheckBox chkCSV 
         Caption         =   "Download to CSV"
         Height          =   315
         Left            =   2040
         TabIndex        =   18
         Top             =   120
         Width           =   2655
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   30
         Left            =   1080
         TabIndex        =   2
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
            Picture         =   "frmReportGen.frx":0FBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportGen.frx":124D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportGen.frx":14DF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView QCTree 
      Height          =   4815
      Left            =   60
      TabIndex        =   0
      Top             =   1380
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   8493
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
            Picture         =   "frmReportGen.frx":176D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportGen.frx":1E7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportGen.frx":2591
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportGen.frx":2CA3
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
      TabIndex        =   17
      Top             =   9855
      Width           =   13230
      _ExtentX        =   23336
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmReportGen.frx":33B5
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   22110
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
            Picture         =   "frmReportGen.frx":3906
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportGen.frx":3BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportGen.frx":4139
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportGen.frx":468A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Filter"
      Height          =   255
      Left            =   9000
      TabIndex        =   16
      Top             =   660
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Select Fields to Report"
      Height          =   195
      Left            =   9000
      TabIndex        =   14
      Top             =   1380
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Tables"
      Height          =   195
      Left            =   5160
      TabIndex        =   13
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Select Source Folder"
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   1140
      Width           =   1935
   End
End
Attribute VB_Name = "frmReportGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim Requirements() As HPQC_FIELDS
Dim BusinessComponent() As HPQC_FIELDS
Dim TestPlan() As HPQC_FIELDS
Dim TestSet() As HPQC_FIELDS
Dim TestInstance() As HPQC_FIELDS
Dim TestRun() As HPQC_FIELDS
Dim Step() As HPQC_FIELDS
Dim Defects() As HPQC_FIELDS
Dim Extract_Fields() As HPQC_FIELDS
Dim RunSteps() As HPQC_FIELDS
Dim AllScripts
Dim IsAutoCheck As Boolean
Dim CheckedItems_() As String

Private Sub Populate_Fields()
Dim i
ReDim Requirements(73)
ReDim BusinessComponent(50)
ReDim TestPlan(37)
ReDim TestSet(33)
ReDim TestInstance(40)
ReDim TestRun(26)
ReDim Step(7)
ReDim Defects(66)
ReDim RunSteps(21)
ReDim Extract_Fields(1)

For i = LBound(Requirements) To UBound(Requirements)
    Requirements(i).Table_Name = "REQUIREMENT"
Next
For i = LBound(BusinessComponent) To UBound(BusinessComponent)
    BusinessComponent(i).Table_Name = "COMPONENT"
Next
For i = LBound(TestPlan) To UBound(TestPlan)
    TestPlan(i).Table_Name = "TEST PLAN"
Next
For i = LBound(TestSet) To UBound(TestSet)
    TestSet(i).Table_Name = "TEST SET"
Next
For i = LBound(TestInstance) To UBound(TestInstance)
    TestInstance(i).Table_Name = "TEST INSTANCE"
Next
For i = LBound(TestRun) To UBound(TestRun)
    TestRun(i).Table_Name = "RUN"
Next
For i = LBound(Step) To UBound(Step)
    Step(i).Table_Name = "STEP"
Next
For i = LBound(Defects) To UBound(Defects)
    Defects(i).Table_Name = "DEFECT"
Next
For i = LBound(RunSteps) To UBound(RunSteps)
    RunSteps(i).Table_Name = "RUN STEPS"
Next

Requirements(1).Field_Name = "Additional Info": Requirements(1).Technical_Name = "RQ_USER_TEMPLATE_02"
Requirements(2).Field_Name = "Attachment": Requirements(2).Technical_Name = "RQ_ATTACHMENT"
Requirements(3).Field_Name = "Author": Requirements(3).Technical_Name = "RQ_REQ_AUTHOR"
Requirements(4).Field_Name = "Change Request (CR) Comment": Requirements(4).Technical_Name = "RQ_DEV_COMMENTS"
Requirements(5).Field_Name = "Creation Date": Requirements(5).Technical_Name = "RQ_REQ_DATE"
Requirements(6).Field_Name = "Creation Time": Requirements(6).Technical_Name = "RQ_REQ_TIME"
Requirements(7).Field_Name = "Direct Cover Status": Requirements(7).Technical_Name = "RQ_REQ_STATUS"
Requirements(8).Field_Name = "Is Folder": Requirements(8).Technical_Name = "RQ_IS_FOLDER"
Requirements(9).Field_Name = "Is Template": Requirements(9).Technical_Name = "RQ_ISTEMPLATE"
Requirements(10).Field_Name = "ITG Assigned To": Requirements(10).Technical_Name = "RQ_REQUEST_ASSIGN_TO"
Requirements(11).Field_Name = "ITG Request Id": Requirements(11).Technical_Name = "RQ_REQUEST_ID"
Requirements(12).Field_Name = "ITG Request Note": Requirements(12).Technical_Name = "RQ_REQUEST_NOTE"
Requirements(13).Field_Name = "ITG Request Status": Requirements(13).Technical_Name = "RQ_REQUEST_STATUS"
Requirements(14).Field_Name = "ITG Request Type": Requirements(14).Technical_Name = "RQ_REQUEST_TYPE"
Requirements(15).Field_Name = "ITG Server URL": Requirements(15).Technical_Name = "RQ_REQUEST_SERVER"
Requirements(16).Field_Name = "ITG Synchronization Data": Requirements(16).Technical_Name = "RQ_REQUEST_UPDATES"
Requirements(17).Field_Name = "Modified": Requirements(17).Technical_Name = "RQ_VTS"
Requirements(18).Field_Name = "Name": Requirements(18).Technical_Name = "RQ_REQ_NAME"
Requirements(19).Field_Name = "Number of sons": Requirements(19).Technical_Name = "RQ_NO_OF_SONS"
Requirements(20).Field_Name = "Old Type (obsolete)": Requirements(20).Technical_Name = "RQ_REQ_TYPE"
Requirements(21).Field_Name = "Peer Reviewer": Requirements(21).Technical_Name = "RQ_USER_TEMPLATE_01"
Requirements(22).Field_Name = "Priority": Requirements(22).Technical_Name = "RQ_REQ_PRIORITY"
Requirements(23).Field_Name = "Product": Requirements(23).Technical_Name = "RQ_REQ_PRODUCT"
Requirements(24).Field_Name = "RBQM Analysis result data": Requirements(24).Technical_Name = "RQ_RBT_ANALYSIS_RESULT_DATA"
Requirements(25).Field_Name = "RBQM Analysis setup data": Requirements(25).Technical_Name = "RQ_RBT_ANALYSIS_SETUP_DATA"
Requirements(26).Field_Name = "RBQM Assessment data": Requirements(26).Technical_Name = "RQ_RBT_ASSESSMENT_DATA"
Requirements(27).Field_Name = "RBQM business impact": Requirements(27).Technical_Name = "RQ_RBT_BSNS_IMPACT"
Requirements(28).Field_Name = "RBQM custom business impact": Requirements(28).Technical_Name = "RQ_RBT_CUSTOM_BSNS_IMPACT"
Requirements(29).Field_Name = "RBQM custom failure probability": Requirements(29).Technical_Name = "RQ_RBT_CUSTOM_FAIL_PROB"
Requirements(30).Field_Name = "RBQM custom Functional Complexity": Requirements(30).Technical_Name = "RQ_RBT_CUSTOM_FUNC_CMPLX"
Requirements(31).Field_Name = "RBQM custom Risk": Requirements(31).Technical_Name = "RQ_RBT_CUSTOM_RISK"
Requirements(32).Field_Name = "RBQM custom testing hours": Requirements(32).Technical_Name = "RQ_RBT_CUSTOM_TESTING_HOURS"
Requirements(33).Field_Name = "RBQM custom testing level": Requirements(33).Technical_Name = "RQ_RBT_CUSTOM_TESTING_LEVEL"
Requirements(34).Field_Name = "RBQM Date of last Analysis": Requirements(34).Technical_Name = "RQ_RBT_LAST_ANALYSIS_DATE"
Requirements(35).Field_Name = "RBQM effective business impact": Requirements(35).Technical_Name = "RQ_RBT_EFFECTIVE_BSNS_IMPACT"
Requirements(36).Field_Name = "RBQM effective failure probability": Requirements(36).Technical_Name = "RQ_RBT_EFFECTIVE_FAIL_PROB"
Requirements(37).Field_Name = "RBQM effective Functional Complexity": Requirements(37).Technical_Name = "RQ_RBT_EFFECTIVE_FUNC_CMPLX"
Requirements(38).Field_Name = "RBQM effective Risk": Requirements(38).Technical_Name = "RQ_RBT_EFFECTIVE_RISK"
Requirements(39).Field_Name = "RBQM estimated RnD effort": Requirements(39).Technical_Name = "RQ_RBT_RND_ESTIM_EFFORT_HOURS"
Requirements(40).Field_Name = "RBQM Exclude from analysis": Requirements(40).Technical_Name = "RQ_RBT_IGNORE_IN_ANALYSIS"
Requirements(41).Field_Name = "RBQM failure probability": Requirements(41).Technical_Name = "RQ_RBT_FAIL_PROB"
Requirements(42).Field_Name = "RBQM Functional Complexity": Requirements(42).Technical_Name = "RQ_RBT_FUNC_CMPLX"
Requirements(43).Field_Name = "RBQM ID of parent analysis req": Requirements(43).Technical_Name = "RQ_RBT_ANALYSIS_PARENT_REQ_ID"
Requirements(44).Field_Name = "RBQM Risk": Requirements(44).Technical_Name = "RQ_RBT_RISK"
Requirements(45).Field_Name = "RBQM testing hours": Requirements(45).Technical_Name = "RQ_RBT_TESTING_HOURS"
Requirements(46).Field_Name = "RBQM testing level": Requirements(46).Technical_Name = "RQ_RBT_TESTING_LEVEL"
Requirements(47).Field_Name = "RBQM use custom business impact": Requirements(47).Technical_Name = "RQ_RBT_USE_CUSTOM_BSNS_IMPACT"
Requirements(48).Field_Name = "RBQM use custom failure probability": Requirements(48).Technical_Name = "RQ_RBT_USE_CUSTOM_FAIL_PROB"
Requirements(49).Field_Name = "RBQM use custom Functional Complexity": Requirements(49).Technical_Name = "RQ_RBT_USE_CUSTOM_FUNC_CMPLX"
Requirements(50).Field_Name = "RBQM use custom results": Requirements(50).Technical_Name = "RQ_RBT_USE_CUSTOM_TL_AND_TE"
Requirements(51).Field_Name = "RBQM use custom Risk": Requirements(51).Technical_Name = "RQ_RBT_USE_CUSTOM_RISK"
Requirements(52).Field_Name = "Req Father ID": Requirements(52).Technical_Name = "RQ_FATHER_ID"
Requirements(53).Field_Name = "Req ID": Requirements(53).Technical_Name = "RQ_REQ_ID"
Requirements(54).Field_Name = "Req Order ID": Requirements(54).Technical_Name = "RQ_ORDER_ID"
Requirements(55).Field_Name = "Req Parent": Requirements(55).Technical_Name = "RQ_FATHER_NAME"
Requirements(56).Field_Name = "Req Path": Requirements(56).Technical_Name = "RQ_REQ_PATH"
Requirements(57).Field_Name = "Req Folder Path": Requirements(57).Technical_Name = "RequirementFolderPath"
Requirements(58).Field_Name = "Requirement Type": Requirements(58).Technical_Name = "RQ_TYPE_ID"
Requirements(59).Field_Name = "Reviewed": Requirements(59).Technical_Name = "RQ_REQ_REVIEWED"
Requirements(60).Field_Name = "Rich Text": Requirements(60).Technical_Name = "RQ_HAS_RICH_CONTENT"
Requirements(61).Field_Name = "Target Cycle": Requirements(61).Technical_Name = "RQ_TARGET_RCYC"
Requirements(62).Field_Name = "Target Release": Requirements(62).Technical_Name = "RQ_TARGET_REL"
Requirements(63).Field_Name = "Version Check In Comments": Requirements(63).Technical_Name = "RQ_VC_CHECKIN_COMMENTS"
Requirements(64).Field_Name = "Version Check In Date": Requirements(64).Technical_Name = "RQ_VC_CHECKIN_DATE"
Requirements(65).Field_Name = "Version Check In Time": Requirements(65).Technical_Name = "RQ_VC_CHECKIN_TIME"
Requirements(66).Field_Name = "Version Check Out Comments": Requirements(66).Technical_Name = "RQ_VC_CHECKOUT_COMMENTS"
Requirements(67).Field_Name = "Version Check Out Date": Requirements(67).Technical_Name = "RQ_VC_CHECKOUT_DATE"
Requirements(68).Field_Name = "Version Check Out Time": Requirements(68).Technical_Name = "RQ_VC_CHECKOUT_TIME"
Requirements(69).Field_Name = "Version Checked In By": Requirements(69).Technical_Name = "RQ_VC_CHECKIN_USER_NAME"
Requirements(70).Field_Name = "Version Checked Out By": Requirements(70).Technical_Name = "RQ_VC_CHECKOUT_USER_NAME"
Requirements(71).Field_Name = "Version Number": Requirements(71).Technical_Name = "RQ_VC_VERSION_NUMBER"
Requirements(72).Field_Name = "Version Status": Requirements(72).Technical_Name = "RQ_VC_STATUS"
Requirements(73).Field_Name = "WRIEF Info": Requirements(73).Technical_Name = "RQ_REQ_COMMENT"

BusinessComponent(1).Field_Name = "Actual Scripting End": BusinessComponent(1).Technical_Name = "CO_USER_TEMPLATE_07"
BusinessComponent(2).Field_Name = "Actual Scripting Start": BusinessComponent(2).Technical_Name = "CO_USER_TEMPLATE_06"
BusinessComponent(3).Field_Name = "Allow iterations": BusinessComponent(3).Technical_Name = "CO_IS_ITERATABLE"
BusinessComponent(4).Field_Name = "Application  Related ID": BusinessComponent(4).Technical_Name = "CO_BPTA_LOCATION_ID_IN_APP"
BusinessComponent(5).Field_Name = "Automation engine": BusinessComponent(5).Technical_Name = "CO_SCRIPT_TYPE"
BusinessComponent(6).Field_Name = "Change Status": BusinessComponent(6).Technical_Name = "CO_BPTA_CHANGE_DETECTED"
BusinessComponent(7).Field_Name = "Comments": BusinessComponent(7).Technical_Name = "CO_DEV_COMMENTS"
BusinessComponent(8).Field_Name = "Application Type": BusinessComponent(8).Technical_Name = "CO_BPTA_COMPONENT_TYPE"
BusinessComponent(9).Field_Name = "Component Folder": BusinessComponent(9).Technical_Name = "BusinessComponentFolderPath"
BusinessComponent(10).Field_Name = "Component Folder ID": BusinessComponent(10).Technical_Name = "CO_FOLDER_ID"
BusinessComponent(11).Field_Name = "Component ID": BusinessComponent(11).Technical_Name = "CO_ID"
BusinessComponent(12).Field_Name = "Component name": BusinessComponent(12).Technical_Name = "CO_NAME"
BusinessComponent(13).Field_Name = "Component Status": BusinessComponent(13).Technical_Name = "CO_STATUS"
BusinessComponent(14).Field_Name = "Components Steps": BusinessComponent(14).Technical_Name = "CO_DATA"
BusinessComponent(15).Field_Name = "Created by": BusinessComponent(15).Technical_Name = "CO_CREATED_BY"
BusinessComponent(16).Field_Name = "Creation date": BusinessComponent(16).Technical_Name = "CO_CREATION_DATE"
BusinessComponent(17).Field_Name = "Default Picture ID": BusinessComponent(17).Technical_Name = "CO_DEFAULT_PICTURE_ID"
BusinessComponent(18).Field_Name = "Deleted on": BusinessComponent(18).Technical_Name = "CO_DELETION_DATE"
BusinessComponent(19).Field_Name = "Description": BusinessComponent(19).Technical_Name = "CO_DESC"
BusinessComponent(20).Field_Name = "Execution Method": BusinessComponent(20).Technical_Name = "CO_USER_TEMPLATE_08"
BusinessComponent(21).Field_Name = "Has Picture": BusinessComponent(21).Technical_Name = "CO_HAS_PICTURE"
BusinessComponent(22).Field_Name = "Is Obsolete": BusinessComponent(22).Technical_Name = "CO_IS_OBSOLETE"
BusinessComponent(23).Field_Name = "Last Detected Change": BusinessComponent(23).Technical_Name = "CO_BPTA_CHANGE_TIMESTAMP"
BusinessComponent(24).Field_Name = "Last Update": BusinessComponent(24).Technical_Name = "CO_BPTA_LAST_UPDATE_TIMESTAMP"
BusinessComponent(25).Field_Name = "Linked Area ID": BusinessComponent(25).Technical_Name = "CO_APP_AREA_ID"
BusinessComponent(26).Field_Name = "Linked flow test ID": BusinessComponent(26).Technical_Name = "CO_BPTA_FLOW_TEST_ID"
BusinessComponent(27).Field_Name = "Original Location": BusinessComponent(27).Technical_Name = "CO_DELETED_FROM_PATH"
BusinessComponent(28).Field_Name = "Peer Reviewer": BusinessComponent(28).Technical_Name = "CO_USER_TEMPLATE_02"
BusinessComponent(29).Field_Name = "Physical path": BusinessComponent(29).Technical_Name = "CO_PHYSICAL_PATH"
BusinessComponent(30).Field_Name = "Planned Scripting Date": BusinessComponent(30).Technical_Name = "CO_USER_TEMPLATE_05"
BusinessComponent(31).Field_Name = "Planned Scripting Date": BusinessComponent(31).Technical_Name = "CO_USER_TEMPLATE_04"
BusinessComponent(32).Field_Name = "QA Reviewed Date": BusinessComponent(32).Technical_Name = "CO_USER_TEMPLATE_09"
BusinessComponent(33).Field_Name = "QA Reviewer": BusinessComponent(33).Technical_Name = "CO_USER_TEMPLATE_03"
BusinessComponent(34).Field_Name = "Scripter": BusinessComponent(34).Technical_Name = "CO_RESPONSIBLE"
BusinessComponent(35).Field_Name = "Object Repository ID": BusinessComponent(35).Technical_Name = "CO_BPTA_SOR_ID"
BusinessComponent(36).Field_Name = "Source ID": BusinessComponent(36).Technical_Name = "CO_SRC_ID"
BusinessComponent(37).Field_Name = "Status": BusinessComponent(37).Technical_Name = "CO_USER_TEMPLATE_01"
BusinessComponent(38).Field_Name = "Step Data": BusinessComponent(38).Technical_Name = "CO_STEPS_DATA"
BusinessComponent(39).Field_Name = "Version": BusinessComponent(39).Technical_Name = "CO_VERSION"
BusinessComponent(40).Field_Name = "Check In Comments": BusinessComponent(40).Technical_Name = "CO_VC_CHECKIN_COMMENTS"
BusinessComponent(41).Field_Name = "Check In Date": BusinessComponent(41).Technical_Name = "CO_VC_CHECKIN_DATE"
BusinessComponent(42).Field_Name = "Check In Time": BusinessComponent(42).Technical_Name = "CO_VC_CHECKIN_TIME"
BusinessComponent(43).Field_Name = "Check Out Comments": BusinessComponent(43).Technical_Name = "CO_VC_CHECKOUT_COMMENTS"
BusinessComponent(44).Field_Name = "Check Out Date": BusinessComponent(44).Technical_Name = "CO_VC_CHECKOUT_DATE"
BusinessComponent(45).Field_Name = "Check Out Time": BusinessComponent(45).Technical_Name = "CO_VC_CHECKOUT_TIME"
BusinessComponent(46).Field_Name = "Checked In By": BusinessComponent(46).Technical_Name = "CO_VC_CHECKIN_USER_NAME"
BusinessComponent(47).Field_Name = "Checked Out By": BusinessComponent(47).Technical_Name = "CO_VC_CHECKOUT_USER_NAME"
BusinessComponent(48).Field_Name = "Version Number": BusinessComponent(48).Technical_Name = "CO_VC_VERSION_NUMBER"
BusinessComponent(49).Field_Name = "Version Stamp": BusinessComponent(49).Technical_Name = "CO_VER_STAMP"
BusinessComponent(50).Field_Name = "Version Status": BusinessComponent(50).Technical_Name = "CO_VC_STATUS"

TestPlan(1).Field_Name = "Actual Scripting End": TestPlan(1).Technical_Name = "TS_USER_TEMPLATE_01"
TestPlan(2).Field_Name = "Actual Scripting Start": TestPlan(2).Technical_Name = "TS_USER_TEMPLATE_06"
TestPlan(3).Field_Name = "Attachment": TestPlan(3).Technical_Name = "TS_ATTACHMENT"
TestPlan(4).Field_Name = "Audit ID for start": TestPlan(4).Technical_Name = "TS_VC_START_AUDIT_ACTION_ID"
TestPlan(5).Field_Name = "Audit ID that is a end": TestPlan(5).Technical_Name = "TS_VC_END_AUDIT_ACTION_ID"
TestPlan(6).Field_Name = "Base Test ID": TestPlan(6).Technical_Name = "TS_BASE_TEST_ID"
TestPlan(7).Field_Name = "Change Status": TestPlan(7).Technical_Name = "TS_BPTA_CHANGE_DETECTED"
TestPlan(8).Field_Name = "Comments": TestPlan(8).Technical_Name = "TS_DEV_COMMENTS"
TestPlan(9).Field_Name = "Creation Date": TestPlan(9).Technical_Name = "TS_CREATION_DATE"
TestPlan(10).Field_Name = "Description": TestPlan(10).Technical_Name = "TS_DESCRIPTION"
TestPlan(11).Field_Name = "Estimated DevTime": TestPlan(11).Technical_Name = "TS_ESTIMATE_DEVTIME"
TestPlan(12).Field_Name = "Execution Status": TestPlan(12).Technical_Name = "TS_EXEC_STATUS"
TestPlan(13).Field_Name = "Mock Phase": TestPlan(13).Technical_Name = "TS_USER_01"
TestPlan(14).Field_Name = "Modified": TestPlan(14).Technical_Name = "TS_VTS"
TestPlan(15).Field_Name = "Path": TestPlan(15).Technical_Name = "TS_PATH"
TestPlan(16).Field_Name = "Peer Reviewer": TestPlan(16).Technical_Name = "TS_USER_TEMPLATE_02"
TestPlan(17).Field_Name = "Planned Scripting End": TestPlan(17).Technical_Name = "TS_USER_TEMPLATE_04"
TestPlan(18).Field_Name = "Planned Scripting Start": TestPlan(18).Technical_Name = "TS_USER_TEMPLATE_05"
TestPlan(19).Field_Name = "QA Reviewer": TestPlan(19).Technical_Name = "TS_USER_TEMPLATE_03"
TestPlan(20).Field_Name = "Scripter": TestPlan(20).Technical_Name = "TS_RESPONSIBLE"
TestPlan(21).Field_Name = "Status": TestPlan(21).Technical_Name = "TS_STATUS"
TestPlan(22).Field_Name = "Step Param": TestPlan(22).Technical_Name = "TS_STEP_PARAM"
TestPlan(23).Field_Name = "Steps": TestPlan(23).Technical_Name = "TS_STEPS"
TestPlan(24).Field_Name = "Subject": TestPlan(24).Technical_Name = "TS_SUBJECT"
TestPlan(25).Field_Name = "Template": TestPlan(25).Technical_Name = "TS_TEMPLATE"
TestPlan(26).Field_Name = "Test Folder": TestPlan(26).Technical_Name = "TestFolderPath"
TestPlan(27).Field_Name = "Test ID": TestPlan(27).Technical_Name = "TS_TEST_ID"
TestPlan(28).Field_Name = "Test Name": TestPlan(28).Technical_Name = "TS_NAME"
TestPlan(29).Field_Name = "Test Runtime Data": TestPlan(29).Technical_Name = "TS_RUNTIME_DATA"
TestPlan(30).Field_Name = "Test Script Status": TestPlan(30).Technical_Name = "TS_USER_TEMPLATE_07"
TestPlan(31).Field_Name = "Type": TestPlan(31).Technical_Name = "TS_TYPE"
TestPlan(32).Field_Name = "Version Comments": TestPlan(32).Technical_Name = "TS_VC_COMMENTS"
TestPlan(33).Field_Name = "Version Date": TestPlan(33).Technical_Name = "TS_VC_DATE"
TestPlan(34).Field_Name = "Version Number": TestPlan(34).Technical_Name = "TS_VC_VERSION_NUMBER"
TestPlan(35).Field_Name = "Version Owner": TestPlan(35).Technical_Name = "TS_VC_USER_NAME"
TestPlan(36).Field_Name = "Version Status": TestPlan(36).Technical_Name = "TS_VC_STATUS"
TestPlan(37).Field_Name = "Version Time": TestPlan(37).Technical_Name = "TS_VC_TIME"

TestSet(1).Field_Name = "Attachment": TestSet(1).Technical_Name = "CY_ATTACHMENT"
TestSet(2).Field_Name = "Baseline": TestSet(2).Technical_Name = "CY_PINNED_BASELINE"
TestSet(3).Field_Name = "CIT Assigned Group": TestSet(3).Technical_Name = "CY_USER_10"
TestSet(4).Field_Name = "Close Date": TestSet(4).Technical_Name = "CY_CLOSE_DATE"
TestSet(5).Field_Name = "Comments": TestSet(5).Technical_Name = "CY_USER_09"
TestSet(6).Field_Name = "Conditions": TestSet(6).Technical_Name = "CY_DESCRIPTION"
TestSet(7).Field_Name = "Configuration": TestSet(7).Technical_Name = "CY_OS_CONFIG"
TestSet(8).Field_Name = "CR #": TestSet(8).Technical_Name = "CY_USER_13"
TestSet(9).Field_Name = "Criticality": TestSet(9).Technical_Name = "CY_USER_11"
TestSet(10).Field_Name = "Dependency": TestSet(10).Technical_Name = "CY_USER_07"
TestSet(11).Field_Name = "Description": TestSet(11).Technical_Name = "CY_COMMENT"
TestSet(12).Field_Name = "Executed by": TestSet(12).Technical_Name = "CY_USER_04"
TestSet(13).Field_Name = "Execution Status": TestSet(13).Technical_Name = "CY_STATUS"
TestSet(14).Field_Name = "ITG Request Id": TestSet(14).Technical_Name = "CY_REQUEST_ID"
TestSet(15).Field_Name = "Mail Settings": TestSet(15).Technical_Name = "CY_MAIL_SETTINGS"
TestSet(16).Field_Name = "Modified": TestSet(16).Technical_Name = "CY_VTS"
TestSet(17).Field_Name = "Open Date": TestSet(17).Technical_Name = "CY_OPEN_DATE"
TestSet(18).Field_Name = "Output": TestSet(18).Technical_Name = "CY_USER_08"
TestSet(19).Field_Name = "Pending CR": TestSet(19).Technical_Name = "CY_USER_12"
TestSet(20).Field_Name = "Planned Execution End": TestSet(20).Technical_Name = "CY_USER_06"
TestSet(21).Field_Name = "Planned Execution Start": TestSet(21).Technical_Name = "CY_USER_05"
TestSet(22).Field_Name = "Planned Scripting End": TestSet(22).Technical_Name = "CY_USER_03"
TestSet(23).Field_Name = "Planned Scripting Start": TestSet(23).Technical_Name = "CY_USER_02"
TestSet(24).Field_Name = "Scripting Status": TestSet(24).Technical_Name = "CY_USER_01"
TestSet(25).Field_Name = "Target Cycle": TestSet(25).Technical_Name = "CY_ASSIGN_RCYC"
TestSet(26).Field_Name = "Test Set ID": TestSet(26).Technical_Name = "CY_CYCLE_ID"
TestSet(27).Field_Name = "Test Set": TestSet(27).Technical_Name = "CY_CYCLE"
TestSet(28).Field_Name = "Test Set Folder ID": TestSet(28).Technical_Name = "CY_FOLDER_ID"
TestSet(29).Field_Name = "Test Set Folder": TestSet(29).Technical_Name = "TestSetFolderPath"
TestSet(30).Field_Name = "Event Handling": TestSet(30).Technical_Name = "CY_EXEC_EVENT_HANDLE"
TestSet(31).Field_Name = "IT Reference Key": TestSet(31).Technical_Name = "CY_USER_14"
TestSet(32).Field_Name = "UAT Key": TestSet(32).Technical_Name = "CY_USER_16"
TestSet(33).Field_Name = "Execution Method": TestSet(33).Technical_Name = "CY_USER_15"

TestInstance(1).Field_Name = "Actual Exec End": TestInstance(1).Technical_Name = "TC_USER_TEMPLATE_05"
TestInstance(2).Field_Name = "Actual Exec Start": TestInstance(2).Technical_Name = "TC_USER_TEMPLATE_04"
TestInstance(3).Field_Name = "Assigned Group": TestInstance(3).Technical_Name = "TC_USER_TEMPLATE_13"
TestInstance(4).Field_Name = "Attachment": TestInstance(4).Technical_Name = "TC_ATTACHMENT"
TestInstance(5).Field_Name = "Baseline": TestInstance(5).Technical_Name = "TC_PINNED_BASELINE"
TestInstance(6).Field_Name = "Check for Dependency": TestInstance(6).Technical_Name = "TC_USER_02"
TestInstance(7).Field_Name = "Comments": TestInstance(7).Technical_Name = "TC_USER_27"
TestInstance(8).Field_Name = "Component Dependency": TestInstance(8).Technical_Name = "TC_USER_25"
TestInstance(9).Field_Name = "Configuration": TestInstance(9).Technical_Name = "TC_OS_CONFIG"
TestInstance(10).Field_Name = "Data Object": TestInstance(10).Technical_Name = "TC_USER_TEMPLATE_06"
TestInstance(11).Field_Name = "Data Scripter": TestInstance(11).Technical_Name = "TC_USER_TEMPLATE_09"
TestInstance(12).Field_Name = "Data Scripting Status": TestInstance(12).Technical_Name = "TC_USER_TEMPLATE_10"
TestInstance(13).Field_Name = "Data Validation Status": TestInstance(13).Technical_Name = "TC_USER_TEMPLATE_11"
TestInstance(14).Field_Name = "eparams": TestInstance(14).Technical_Name = "TC_EPARAMS"
TestInstance(15).Field_Name = "Execution Date": TestInstance(15).Technical_Name = "TC_EXEC_DATE"
TestInstance(16).Field_Name = "Iterations": TestInstance(16).Technical_Name = "TC_ITERATIONS"
TestInstance(17).Field_Name = "Modified": TestInstance(17).Technical_Name = "TC_VTS"
TestInstance(18).Field_Name = " Test Script Status": TestInstance(18).Technical_Name = "TC_USER_TEMPLATE_12"
TestInstance(19).Field_Name = "Output Parameter": TestInstance(19).Technical_Name = "TC_USER_26"
TestInstance(20).Field_Name = "Planned Exec End": TestInstance(20).Technical_Name = "TC_USER_TEMPLATE_03"
TestInstance(21).Field_Name = "Planned Exec Start": TestInstance(21).Technical_Name = "TC_PLAN_SCHEDULING_DATE"
TestInstance(22).Field_Name = "Planned Exec Time": TestInstance(22).Technical_Name = "TC_PLAN_SCHEDULING_TIME"
TestInstance(23).Field_Name = "Planned Host Name": TestInstance(23).Technical_Name = "TC_HOST_NAME"
TestInstance(24).Field_Name = "Planned Validation End": TestInstance(24).Technical_Name = "TC_USER_TEMPLATE_08"
TestInstance(25).Field_Name = "Planned Validation Start": TestInstance(25).Technical_Name = "TC_USER_TEMPLATE_07"
TestInstance(26).Field_Name = "QA Review Status": TestInstance(26).Technical_Name = "TC_USER_TEMPLATE_01"
TestInstance(27).Field_Name = "QA Reviewer": TestInstance(27).Technical_Name = "TC_USER_TEMPLATE_02"
TestInstance(28).Field_Name = "Responsible Tester": TestInstance(28).Technical_Name = "TC_TESTER_NAME"
TestInstance(29).Field_Name = "Status": TestInstance(29).Technical_Name = "TC_STATUS"
TestInstance(30).Field_Name = "Target Cycle": TestInstance(30).Technical_Name = "TC_ASSIGN_RCYC"
TestInstance(31).Field_Name = "Test": TestInstance(31).Technical_Name = "TC_TEST_ID"
TestInstance(32).Field_Name = "Event Handling": TestInstance(32).Technical_Name = "TC_EXEC_EVENT_HANDLE"
TestInstance(33).Field_Name = "Test Instance": TestInstance(33).Technical_Name = "TC_TEST_INSTANCE"
TestInstance(34).Field_Name = "Test Instance ID": TestInstance(34).Technical_Name = "TC_TESTCYCL_ID"
TestInstance(35).Field_Name = "Test Lab: Description": TestInstance(35).Technical_Name = "TC_USER_01"
TestInstance(36).Field_Name = "Test Order": TestInstance(36).Technical_Name = "TC_TEST_ORDER"
TestInstance(37).Field_Name = "Test Set": TestInstance(37).Technical_Name = "TC_CYCLE"
TestInstance(38).Field_Name = "Tester": TestInstance(38).Technical_Name = "TC_ACTUAL_TESTER"
TestInstance(39).Field_Name = "Test Set ID": TestInstance(39).Technical_Name = "TC_CYCLE_ID"
TestInstance(40).Field_Name = "Time": TestInstance(40).Technical_Name = "TC_EXEC_TIME"

TestRun(1).Field_Name = "Attachment": TestRun(1).Technical_Name = "RN_ATTACHMENT"
TestRun(2).Field_Name = "Test Set": TestRun(2).Technical_Name = "RN_CYCLE"
TestRun(3).Field_Name = "Cycle ID": TestRun(3).Technical_Name = "RN_CYCLE_ID"
TestRun(4).Field_Name = "Duration": TestRun(4).Technical_Name = "RN_DURATION"
TestRun(5).Field_Name = "Exec Date": TestRun(5).Technical_Name = "RN_EXECUTION_DATE"
TestRun(6).Field_Name = "Exec Time": TestRun(6).Technical_Name = "RN_EXECUTION_TIME"
TestRun(7).Field_Name = "Host": TestRun(7).Technical_Name = "RN_HOST"
TestRun(8).Field_Name = "Operating System": TestRun(8).Technical_Name = "RN_OS_NAME"
TestRun(9).Field_Name = "Run#": TestRun(9).Technical_Name = "RN_RUN_ID"
TestRun(10).Field_Name = "Status": TestRun(10).Technical_Name = "RN_STATUS"
TestRun(11).Field_Name = "Test ID": TestRun(11).Technical_Name = "RN_TEST_ID"
TestRun(12).Field_Name = "Test Instance ID": TestRun(12).Technical_Name = "RN_TESTCYCL_ID"
TestRun(13).Field_Name = "Tester": TestRun(13).Technical_Name = "RN_TESTER_NAME"
TestRun(14).Field_Name = "Run VC User": TestRun(14).Technical_Name = "RN_VC_LOKEDBY"
TestRun(15).Field_Name = "Baseline": TestRun(15).Technical_Name = "RN_PINNED_BASELINE"
TestRun(16).Field_Name = "Change Status": TestRun(16).Technical_Name = "RN_BPTA_CHANGE_DETECTED"
TestRun(17).Field_Name = "Detection Mode": TestRun(17).Technical_Name = "RN_BPTA_CHANGE_AWARENESS"
TestRun(18).Field_Name = "Run VC Version": TestRun(18).Technical_Name = "RN_VC_VERSION_NUMBER"
TestRun(19).Field_Name = "OS Service Pack": TestRun(19).Technical_Name = "RN_OS_SP"
TestRun(20).Field_Name = "Target Cycle": TestRun(20).Technical_Name = "RN_ASSIGN_RCYC"
TestRun(21).Field_Name = "Path": TestRun(21).Technical_Name = "RN_PATH"
TestRun(22).Field_Name = "Test Instance": TestRun(22).Technical_Name = "RN_TEST_INSTANCE"
TestRun(23).Field_Name = "Run VC Status": TestRun(23).Technical_Name = "RN_VC_STATUS"
TestRun(24).Field_Name = "Run Name": TestRun(24).Technical_Name = "RN_RUN_NAME"
TestRun(25).Field_Name = "OS Build Number": TestRun(25).Technical_Name = "RN_OS_BUILD"
TestRun(26).Field_Name = "Configuration": TestRun(26).Technical_Name = "RN_OS_CONFIG"

Step(1).Field_Name = "Checked Out By": Step(1).Technical_Name = "CS_VC_CHECKOUT_USER_NAME"
Step(2).Field_Name = "Component ID": Step(2).Technical_Name = "CS_COMPONENT_ID"
Step(3).Field_Name = "Description": Step(3).Technical_Name = "CS_DESCRIPTION"
Step(4).Field_Name = "Expected Result": Step(4).Technical_Name = "CS_EXPECTED"
Step(5).Field_Name = "Step Name": Step(5).Technical_Name = "CS_STEP_NAME"
Step(6).Field_Name = "Step Order": Step(6).Technical_Name = "CS_STEP_ORDER"
Step(7).Field_Name = "Step ID": Step(7).Technical_Name = "CS_STEP_ID"

Defects(1).Field_Name = "Attachment": Defects(1).Technical_Name = "BG_ATTACHMENT"
Defects(2).Field_Name = "Version Stamp": Defects(2).Technical_Name = "BG_BUG_VER_STAMP"
Defects(3).Field_Name = "Closed in Version": Defects(3).Technical_Name = "BG_CLOSING_VERSION"
Defects(4).Field_Name = "Cycle ID": Defects(4).Technical_Name = "BG_CYCLE_ID"
Defects(5).Field_Name = "Detected By": Defects(5).Technical_Name = "BG_DETECTED_BY"
Defects(6).Field_Name = "Detected in Version": Defects(6).Technical_Name = "BG_DETECTION_VERSION"
Defects(7).Field_Name = "Comments": Defects(7).Technical_Name = "BG_DEV_COMMENTS"
Defects(8).Field_Name = "Estimated Fix Time": Defects(8).Technical_Name = "BG_ESTIMATED_FIX_TIME"
Defects(9).Field_Name = "Extended Reference": Defects(9).Technical_Name = "BG_EXTENDED_REFERENCE"
Defects(10).Field_Name = "Planned Closing Version": Defects(10).Technical_Name = "BG_PLANNED_CLOSING_VER"
Defects(11).Field_Name = "Priority": Defects(11).Technical_Name = "BG_PRIORITY"
Defects(12).Field_Name = "Project": Defects(12).Technical_Name = "BG_PROJECT"
Defects(13).Field_Name = "ITG Request Id": Defects(13).Technical_Name = "BG_REQUEST_ID"
Defects(14).Field_Name = "ITG Request Note": Defects(14).Technical_Name = "BG_REQUEST_NOTE"
Defects(15).Field_Name = "ITG Server URL": Defects(15).Technical_Name = "BG_REQUEST_SERVER"
Defects(16).Field_Name = "ITG Request Type": Defects(16).Technical_Name = "BG_REQUEST_TYPE"
Defects(17).Field_Name = "Run Reference": Defects(17).Technical_Name = "BG_RUN_REFERENCE"
Defects(18).Field_Name = "Severity": Defects(18).Technical_Name = "BG_SEVERITY"
Defects(19).Field_Name = "Status": Defects(19).Technical_Name = "BG_STATUS"
Defects(20).Field_Name = "Step Reference": Defects(20).Technical_Name = "BG_STEP_REFERENCE"
Defects(21).Field_Name = "Summary": Defects(21).Technical_Name = "BG_SUMMARY"
Defects(22).Field_Name = "Test Reference": Defects(22).Technical_Name = "BG_TEST_REFERENCE"
Defects(23).Field_Name = "Normal Correction": Defects(23).Technical_Name = "BG_USER_01"
Defects(24).Field_Name = "Data Validation": Defects(24).Technical_Name = "BG_USER_02"
Defects(25).Field_Name = "WRIEF Defect Root Cause": Defects(25).Technical_Name = "BG_USER_05"
Defects(26).Field_Name = "BAU Project": Defects(26).Technical_Name = "BG_USER_06"
Defects(27).Field_Name = "Temporary/Permanent Fix": Defects(27).Technical_Name = "BG_USER_07"
Defects(28).Field_Name = "Fixed By": Defects(28).Technical_Name = "BG_USER_08"
Defects(29).Field_Name = "Business Work Around": Defects(29).Technical_Name = "BG_USER_11"
Defects(30).Field_Name = "Root Cause (Conversion)": Defects(30).Technical_Name = "BG_USER_13"
Defects(31).Field_Name = "Detected in Release": Defects(31).Technical_Name = "BG_DETECTED_IN_REL"
Defects(32).Field_Name = "Test Stage": Defects(32).Technical_Name = "BG_DETECTED_IN_RCYC"
Defects(33).Field_Name = "Target Cycle": Defects(33).Technical_Name = "BG_TARGET_RCYC"
Defects(34).Field_Name = "Transaction Code": Defects(34).Technical_Name = "BG_USER_TEMPLATE_01"
Defects(35).Field_Name = "Test Environment": Defects(35).Technical_Name = "BG_USER_TEMPLATE_02"
Defects(36).Field_Name = "Category": Defects(36).Technical_Name = "BG_USER_TEMPLATE_03"
Defects(37).Field_Name = "Assigned Team": Defects(37).Technical_Name = "BG_USER_TEMPLATE_04"
Defects(38).Field_Name = "CR #": Defects(38).Technical_Name = "BG_USER_TEMPLATE_06"
Defects(39).Field_Name = "Test Phase": Defects(39).Technical_Name = "BG_USER_TEMPLATE_08"
Defects(40).Field_Name = "Detected Group": Defects(40).Technical_Name = "BG_USER_03"
Defects(41).Field_Name = "To Mail": Defects(41).Technical_Name = "BG_TO_MAIL"
Defects(42).Field_Name = "Line Items Impacted": Defects(42).Technical_Name = "BG_USER_12"
Defects(43).Field_Name = "Actual Fix Date": Defects(43).Technical_Name = "BG_CLOSING_DATE"
Defects(44).Field_Name = "SAP Msg #": Defects(44).Technical_Name = "BG_USER_TEMPLATE_10"
Defects(45).Field_Name = "WRIEF #": Defects(45).Technical_Name = "BG_USER_TEMPLATE_13"
Defects(46).Field_Name = "WRIEF Scope": Defects(46).Technical_Name = "BG_USER_TEMPLATE_16"
Defects(47).Field_Name = "PR Root Causes (R2b)": Defects(47).Technical_Name = "BG_USER_10"
Defects(48).Field_Name = "Has Change": Defects(48).Technical_Name = "BG_HAS_CHANGE"
Defects(49).Field_Name = "Modified": Defects(49).Technical_Name = "BG_VTS"
Defects(50).Field_Name = "Assigned To": Defects(50).Technical_Name = "BG_RESPONSIBLE"
Defects(51).Field_Name = "Data Object": Defects(51).Technical_Name = "BG_USER_TEMPLATE_11"
Defects(52).Field_Name = "Process Area": Defects(52).Technical_Name = "BG_USER_TEMPLATE_05"
Defects(53).Field_Name = "Defect ID": Defects(53).Technical_Name = "BG_BUG_ID"
Defects(54).Field_Name = "Transport ID #": Defects(54).Technical_Name = "BG_USER_TEMPLATE_07"
Defects(55).Field_Name = "Description": Defects(55).Technical_Name = "BG_DESCRIPTION"
Defects(56).Field_Name = "Estimated Fix Date": Defects(56).Technical_Name = "BG_USER_TEMPLATE_09"
Defects(57).Field_Name = "Actual Fix Time": Defects(57).Technical_Name = "BG_ACTUAL_FIX_TIME"
Defects(58).Field_Name = "CR Status": Defects(58).Technical_Name = "BG_USER_TEMPLATE_14"
Defects(59).Field_Name = "TestSet Reference": Defects(59).Technical_Name = "BG_CYCLE_REFERENCE"
Defects(60).Field_Name = "Target Release": Defects(60).Technical_Name = "BG_TARGET_REL"
Defects(61).Field_Name = "Resolution": Defects(61).Technical_Name = "BG_USER_TEMPLATE_12"
Defects(62).Field_Name = "Subject": Defects(62).Technical_Name = "BG_SUBJECT"
Defects(63).Field_Name = "Detected on Date": Defects(63).Technical_Name = "BG_DETECTION_DATE"
Defects(64).Field_Name = "WRIEF Defect Type": Defects(64).Technical_Name = "BG_USER_04"
Defects(65).Field_Name = "Reproducible": Defects(65).Technical_Name = "BG_REPRODUCIBLE"
Defects(66).Field_Name = "Portal Root Cause": Defects(66).Technical_Name = "BG_USER_09"

RunSteps(1).Field_Name = "Actual": RunSteps(1).Technical_Name = "ST_ACTUAL"
RunSteps(2).Field_Name = "Attachment": RunSteps(2).Technical_Name = "ST_ATTACHMENT"
RunSteps(3).Field_Name = "Component Step Data": RunSteps(3).Technical_Name = "ST_COMPONENT_DATA"
RunSteps(4).Field_Name = "Condition": RunSteps(4).Technical_Name = "ST_BPTA_CONDITION"
RunSteps(5).Field_Name = "Description": RunSteps(5).Technical_Name = "ST_DESCRIPTION"
RunSteps(6).Field_Name = "DesignStep ID": RunSteps(6).Technical_Name = "ST_DESSTEP_ID"
RunSteps(7).Field_Name = "Exec Date": RunSteps(7).Technical_Name = "ST_EXECUTION_DATE"
RunSteps(8).Field_Name = "Exec Time": RunSteps(8).Technical_Name = "ST_EXECUTION_TIME"
RunSteps(9).Field_Name = "Expected": RunSteps(9).Technical_Name = "ST_EXPECTED"
RunSteps(10).Field_Name = "Extended Reference": RunSteps(10).Technical_Name = "ST_EXTENDED_REFERENCE"
RunSteps(11).Field_Name = "Level": RunSteps(11).Technical_Name = "ST_OBJ_ID"
RunSteps(12).Field_Name = "Level": RunSteps(12).Technical_Name = "ST_LEVEL"
RunSteps(13).Field_Name = "Line_no": RunSteps(13).Technical_Name = "ST_LINE_NO"
RunSteps(14).Field_Name = "Path": RunSteps(14).Technical_Name = "ST_PATH"
RunSteps(15).Field_Name = "Run ID": RunSteps(15).Technical_Name = "ST_RUN_ID"
RunSteps(16).Field_Name = "Source Test": RunSteps(16).Technical_Name = "ST_TEST_ID"
RunSteps(17).Field_Name = "Status": RunSteps(17).Technical_Name = "ST_STATUS"
RunSteps(18).Field_Name = "Step ID": RunSteps(18).Technical_Name = "ST_ID"
RunSteps(19).Field_Name = "Step Name": RunSteps(19).Technical_Name = "ST_STEP_NAME"
RunSteps(20).Field_Name = "Step Order": RunSteps(20).Technical_Name = "ST_STEP_ORDER"
RunSteps(21).Field_Name = "Step Parent ID": RunSteps(21).Technical_Name = "ST_PARENT_ID"


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

Private Sub flxImport_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 67 And Shift = vbCtrlMask Then
    Clipboard.Clear
    Clipboard.SetText flxImport.Clip
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
SSTab1.Top = 6360
SSTab1.Left = 60
txtSQL.Top = 420
txtSQL.Left = 120
lstExtract.Top = 420
lstExtract.Left = 120
flxImport.Top = 420
flxImport.Left = 120
SSTab1.width = Me.width - SSTab1.Left - 200
SSTab1.height = stsBar.Top - SSTab1.Top - 200
txtSQL.width = SSTab1.width - txtSQL.Left - 200
txtSQL.height = SSTab1.height - txtSQL.Top - 200
lstExtract.height = SSTab1.height - lstExtract.Top - 200
flxImport.width = SSTab1.width - flxImport.Left - 200
flxImport.height = SSTab1.height - flxImport.Top - 200
End Sub

Sub AddField(TableName As String, FieldName As String, TechName As String)
Dim i
    If Extract_Fields(1).Table_Name = "" Then
        Extract_Fields(UBound(Extract_Fields)).Table_Name = TableName
        Extract_Fields(UBound(Extract_Fields)).Field_Name = FieldName
        Extract_Fields(UBound(Extract_Fields)).Technical_Name = TechName
        If InStr(1, TechName, "_") = 0 Then Extract_Fields(UBound(Extract_Fields)).IsSpecial = True
    Else
        For i = LBound(Extract_Fields) To UBound(Extract_Fields)
            If Extract_Fields(i).Table_Name = TableName And Extract_Fields(i).Field_Name = FieldName And Extract_Fields(i).Technical_Name = TechName Then
                Exit Sub
            End If
        Next
        ReDim Preserve Extract_Fields(UBound(Extract_Fields) + 1)
        Extract_Fields(UBound(Extract_Fields)).Table_Name = TableName
        Extract_Fields(UBound(Extract_Fields)).Field_Name = FieldName
        Extract_Fields(UBound(Extract_Fields)).Technical_Name = TechName
        If InStr(1, TechName, "_") = 0 Then Extract_Fields(UBound(Extract_Fields)).IsSpecial = True
    End If
End Sub

Sub RemoveField(TableName As String, FieldName As String, TechName As String)
    Dim i, stringFunct As New clsStrings
    On Error Resume Next
    If UBound(Extract_Fields) = 1 And TableName = Extract_Fields(UBound(Extract_Fields)).Table_Name And FieldName = Extract_Fields(UBound(Extract_Fields)).Field_Name And TechName = Extract_Fields(UBound(Extract_Fields)).Technical_Name Then
        ReDim Extract_Fields(1): Exit Sub
    End If
    For i = LBound(Extract_Fields) To UBound(Extract_Fields)
        If TableName = Extract_Fields(i).Table_Name And FieldName = Extract_Fields(i).Field_Name And TechName = Extract_Fields(i).Technical_Name Then
            Array_RemoveItem_1 CInt(i)
        End If
    Next
End Sub

Sub AddFilter(TableName As String, FieldName As String, TechName As String, Filter As String)
    Dim i, stringFunct As New clsStrings
    On Error Resume Next
    For i = LBound(Extract_Fields) To UBound(Extract_Fields)
        If TableName = Extract_Fields(i).Table_Name And FieldName = Extract_Fields(i).Field_Name And TechName = Extract_Fields(i).Technical_Name Then
            Extract_Fields(i).Filter_Value = Trim(txtFilter.Text)
        End If
    Next
End Sub

Sub ShowFilter(TableName As String, FieldName As String, TechName As String)
    Dim i, stringFunct As New clsStrings
    On Error Resume Next
    txtFilter.Text = ""
    For i = LBound(Extract_Fields) To UBound(Extract_Fields)
        If TableName = Extract_Fields(i).Table_Name And FieldName = Extract_Fields(i).Field_Name And TechName = Extract_Fields(i).Technical_Name Then
            txtFilter.Text = Extract_Fields(i).Filter_Value
        End If
    Next
End Sub

Sub Populate_Order()
Dim i
lstExtract.Clear
For i = LBound(Extract_Fields) To UBound(Extract_Fields)
    lstExtract.AddItem Extract_Fields(i).Table_Name & " - " & Extract_Fields(i).Field_Name & " (" & Extract_Fields(i).Filter_Value & ")"
Next
End Sub

Private Sub lstFields_Click()
Select Case lstTables.List(lstTables.ListIndex)
    Case "REQUIREMENT"
        ShowFilter Requirements(lstFields.ListIndex + 1).Table_Name, Requirements(lstFields.ListIndex + 1).Field_Name, Requirements(lstFields.ListIndex + 1).Technical_Name
    Case "COMPONENT"
        ShowFilter BusinessComponent(lstFields.ListIndex + 1).Table_Name, BusinessComponent(lstFields.ListIndex + 1).Field_Name, BusinessComponent(lstFields.ListIndex + 1).Technical_Name
    Case "TEST PLAN"
        ShowFilter TestPlan(lstFields.ListIndex + 1).Table_Name, TestPlan(lstFields.ListIndex + 1).Field_Name, TestPlan(lstFields.ListIndex + 1).Technical_Name
    Case "TEST SET"
        ShowFilter TestSet(lstFields.ListIndex + 1).Table_Name, TestSet(lstFields.ListIndex + 1).Field_Name, TestSet(lstFields.ListIndex + 1).Technical_Name
    Case "TEST INSTANCE"
        ShowFilter TestInstance(lstFields.ListIndex + 1).Table_Name, TestInstance(lstFields.ListIndex + 1).Field_Name, TestInstance(lstFields.ListIndex + 1).Technical_Name
    Case "RUN"
        ShowFilter TestRun(lstFields.ListIndex + 1).Table_Name, TestRun(lstFields.ListIndex + 1).Field_Name, TestRun(lstFields.ListIndex + 1).Technical_Name
    Case "STEP"
        ShowFilter Step(lstFields.ListIndex + 1).Table_Name, Step(lstFields.ListIndex + 1).Field_Name, Step(lstFields.ListIndex + 1).Technical_Name
    Case "DEFECT"
        ShowFilter Defects(lstFields.ListIndex + 1).Table_Name, Defects(lstFields.ListIndex + 1).Field_Name, Defects(lstFields.ListIndex + 1).Technical_Name
    Case "RUN STEPS"
        ShowFilter RunSteps(lstFields.ListIndex + 1).Table_Name, RunSteps(lstFields.ListIndex + 1).Field_Name, RunSteps(lstFields.ListIndex + 1).Technical_Name
End Select
End Sub

Private Sub lstFields_ItemCheck(Item As Integer)
If IsAutoCheck = True Then Exit Sub
If cmbModule.Enabled = True Then Exit Sub
If lstFields.Selected(Item) = True Then
    Select Case lstTables.List(lstTables.ListIndex)
    Case "REQUIREMENT"
        AddField Requirements(Item + 1).Table_Name, Requirements(Item + 1).Field_Name, Requirements(Item + 1).Technical_Name
    Case "COMPONENT"
        AddField BusinessComponent(Item + 1).Table_Name, BusinessComponent(Item + 1).Field_Name, BusinessComponent(Item + 1).Technical_Name
    Case "TEST PLAN"
        AddField TestPlan(Item + 1).Table_Name, TestPlan(Item + 1).Field_Name, TestPlan(Item + 1).Technical_Name
    Case "TEST SET"
        AddField TestSet(Item + 1).Table_Name, TestSet(Item + 1).Field_Name, TestSet(Item + 1).Technical_Name
    Case "TEST INSTANCE"
        AddField TestInstance(Item + 1).Table_Name, TestInstance(Item + 1).Field_Name, TestInstance(Item + 1).Technical_Name
    Case "RUN"
        AddField TestRun(Item + 1).Table_Name, TestRun(Item + 1).Field_Name, TestRun(Item + 1).Technical_Name
    Case "STEP"
        AddField Step(Item + 1).Table_Name, Step(Item + 1).Field_Name, Step(Item + 1).Technical_Name
    Case "DEFECT"
        AddField Defects(Item + 1).Table_Name, Defects(Item + 1).Field_Name, Defects(Item + 1).Technical_Name
    Case "RUN STEPS"
        AddField RunSteps(Item + 1).Table_Name, RunSteps(Item + 1).Field_Name, RunSteps(Item + 1).Technical_Name
    End Select
Else
    Select Case lstTables.List(lstTables.ListIndex)
    Case "REQUIREMENT"
        RemoveField Requirements(Item + 1).Table_Name, Requirements(Item + 1).Field_Name, Requirements(Item + 1).Technical_Name
    Case "COMPONENT"
        RemoveField BusinessComponent(Item + 1).Table_Name, BusinessComponent(Item + 1).Field_Name, BusinessComponent(Item + 1).Technical_Name
    Case "TEST PLAN"
        RemoveField TestPlan(Item + 1).Table_Name, TestPlan(Item + 1).Field_Name, TestPlan(Item + 1).Technical_Name
    Case "TEST SET"
        RemoveField TestSet(Item + 1).Table_Name, TestSet(Item + 1).Field_Name, TestSet(Item + 1).Technical_Name
    Case "TEST INSTANCE"
        RemoveField TestInstance(Item + 1).Table_Name, TestInstance(Item + 1).Field_Name, TestInstance(Item + 1).Technical_Name
    Case "RUN"
        RemoveField TestRun(Item + 1).Table_Name, TestRun(Item + 1).Field_Name, TestRun(Item + 1).Technical_Name
    Case "STEP"
        RemoveField Step(Item + 1).Table_Name, Step(Item + 1).Field_Name, Step(Item + 1).Technical_Name
    Case "DEFECT"
        RemoveField Defects(Item + 1).Table_Name, Defects(Item + 1).Field_Name, Defects(Item + 1).Technical_Name
    Case "RUN STEPS"
        RemoveField RunSteps(Item + 1).Table_Name, RunSteps(Item + 1).Field_Name, RunSteps(Item + 1).Technical_Name
    End Select
End If

    Select Case lstTables.List(lstTables.ListIndex)
    Case "REQUIREMENT"
        ShowFilter Requirements(Item + 1).Table_Name, Requirements(Item + 1).Field_Name, Requirements(Item + 1).Technical_Name
    Case "COMPONENT"
        ShowFilter BusinessComponent(Item + 1).Table_Name, BusinessComponent(Item + 1).Field_Name, BusinessComponent(Item + 1).Technical_Name
    Case "TEST PLAN"
        ShowFilter TestPlan(Item + 1).Table_Name, TestPlan(Item + 1).Field_Name, TestPlan(Item + 1).Technical_Name
    Case "TEST SET"
        ShowFilter TestSet(Item + 1).Table_Name, TestSet(Item + 1).Field_Name, TestSet(Item + 1).Technical_Name
    Case "TEST INSTANCE"
        ShowFilter TestInstance(Item + 1).Table_Name, TestInstance(Item + 1).Field_Name, TestInstance(Item + 1).Technical_Name
    Case "RUN"
        ShowFilter TestRun(Item + 1).Table_Name, TestRun(Item + 1).Field_Name, TestRun(Item + 1).Technical_Name
    Case "STEP"
        ShowFilter Step(Item + 1).Table_Name, Step(Item + 1).Field_Name, Step(Item + 1).Technical_Name
    Case "DEFECT"
        ShowFilter Defects(Item + 1).Table_Name, Defects(Item + 1).Field_Name, Defects(Item + 1).Technical_Name
    Case "RUN STEP"
        ShowFilter RunSteps(Item + 1).Table_Name, RunSteps(Item + 1).Field_Name, RunSteps(Item + 1).Technical_Name
    End Select
Populate_Order
End Sub

Private Sub lstTables_Click()
Dim i
If cmbModule.Enabled = True Then Exit Sub
Select Case lstTables.List(lstTables.ListIndex)
Case "REQUIREMENT"
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(Requirements) To UBound(Requirements)
        lstFields.AddItem Requirements(i).Field_Name
    Next
Case "COMPONENT"
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(BusinessComponent) To UBound(BusinessComponent)
        lstFields.AddItem BusinessComponent(i).Field_Name
    Next
Case "TEST PLAN"
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestPlan) To UBound(TestPlan)
        lstFields.AddItem TestPlan(i).Field_Name
    Next
Case "TEST SET"
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestSet) To UBound(TestSet)
        lstFields.AddItem TestSet(i).Field_Name
    Next
Case "TEST INSTANCE"
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestInstance) To UBound(TestInstance)
        lstFields.AddItem TestInstance(i).Field_Name
    Next
Case "RUN"
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestRun) To UBound(TestRun)
        lstFields.AddItem TestRun(i).Field_Name
    Next
Case "STEP"
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(Step) To UBound(Step)
        lstFields.AddItem Step(i).Field_Name
    Next
Case "DEFECT"
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(Defects) To UBound(Defects)
        lstFields.AddItem Defects(i).Field_Name
    Next
Case "RUN STEPS"
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(RunSteps) To UBound(RunSteps)
        lstFields.AddItem RunSteps(i).Field_Name
    Next
End Select
IsAutoCheck = True
Check_UnCheck_List
IsAutoCheck = False
End Sub

Private Sub Check_UnCheck_List()
Dim i, j
For i = LBound(Extract_Fields) To UBound(Extract_Fields)
    If Extract_Fields(i).Table_Name = lstTables.List(lstTables.ListIndex) Then
        For j = 0 To lstFields.ListCount - 1
            If lstFields.List(j) = Extract_Fields(i).Field_Name Then
                lstFields.Selected(j) = True
            End If
        Next
    End If
Next
End Sub

Private Sub cmdModule_Click()
QCTree.Nodes.Clear
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
ReDim Extract_Fields(1)
If cmbModule.Text = "REQUIREMENT" Then
    QCTree.Nodes.Add , , "Root", "Requirements", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT RQ_REQ_ID, RQ_REQ_NAME FROM REQ WHERE RQ_FATHER_ID = '0' AND RQ_TYPE_ID = '1' ORDER BY RQ_REQ_ID"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("RQ_REQ_ID")), rs.FieldValue("RQ_REQ_NAME"), 1
        rs.Next
    Next
    lstTables.Clear
    lstTables.AddItem "TEST PLAN"
    lstTables.AddItem "REQUIREMENT"
    lstTables.ListIndex = 0
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestPlan) To UBound(TestPlan)
        lstFields.AddItem BusinessComponent(i).Field_Name
    Next
    QCTree.Nodes.Item(2).Selected = True
    cmbModule.Enabled = False
ElseIf cmbModule.Text = "COMPONENT" Then
    QCTree.Nodes.Add , , "Root", "Components", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT FC_ID, FC_NAME FROM COMPONENT_FOLDER WHERE FC_FATHER_ID = 1 ORDER BY FC_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("FC_ID")), rs.FieldValue("FC_NAME"), 1
        rs.Next
    Next
    lstTables.Clear
    lstTables.AddItem "COMPONENT"
    lstTables.AddItem "STEP"
    lstTables.ListIndex = 0
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(BusinessComponent) To UBound(BusinessComponent)
        lstFields.AddItem BusinessComponent(i).Field_Name
    Next
    QCTree.Nodes.Item(2).Selected = True
    cmbModule.Enabled = False
ElseIf cmbModule.Text = "TEST PLAN" Then
    QCTree.Nodes.Add , , "Root", "Subject", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT AL_ITEM_ID, AL_DESCRIPTION FROM ALL_LISTS WHERE AL_FATHER_ID = 2 ORDER BY AL_DESCRIPTION"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("AL_ITEM_ID")), rs.FieldValue("AL_DESCRIPTION"), 1
        rs.Next
    Next
    lstTables.Clear
    lstTables.AddItem "TEST PLAN"
    lstTables.AddItem "COMPONENT"
    'lstTables.AddItem "STEP"
    lstTables.ListIndex = 0
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestPlan) To UBound(TestPlan)
        lstFields.AddItem TestPlan(i).Field_Name
    Next
    QCTree.Nodes.Item(2).Selected = True
    cmbModule.Enabled = False
ElseIf cmbModule.Text = "TEST SET" Then
    QCTree.Nodes.Add , , "Root", "Root", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = 0 ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    lstTables.Clear
    lstTables.AddItem "TEST SET"
    lstTables.ListIndex = 0
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestSet) To UBound(TestSet)
        lstFields.AddItem TestSet(i).Field_Name
    Next
    QCTree.Nodes.Item(2).Selected = True
    cmbModule.Enabled = False
ElseIf cmbModule.Text = "TEST INSTANCE" Then
    QCTree.Nodes.Add , , "Root", "Root", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = 0 ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    lstTables.Clear
    lstTables.AddItem "TEST SET"
    lstTables.AddItem "TEST PLAN"
    lstTables.AddItem "COMPONENT"
    lstTables.AddItem "TEST INSTANCE"
    lstTables.ListIndex = 0
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestSet) To UBound(TestSet)
        lstFields.AddItem TestSet(i).Field_Name
    Next
    QCTree.Nodes.Item(2).Selected = True
    cmbModule.Enabled = False
ElseIf cmbModule.Text = "RUN" Then
    QCTree.Nodes.Add , , "Root", "Root", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = 0 ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    lstTables.Clear
    lstTables.AddItem "TEST SET"
    lstTables.AddItem "TEST PLAN"
    lstTables.AddItem "COMPONENT"
    lstTables.AddItem "TEST INSTANCE"
    lstTables.AddItem "RUN"
    lstTables.ListIndex = 0
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestSet) To UBound(TestSet)
        lstFields.AddItem TestSet(i).Field_Name
    Next
    QCTree.Nodes.Item(2).Selected = True
    cmbModule.Enabled = False
ElseIf cmbModule.Text = "STEP" Then
    QCTree.Nodes.Add , , "Root", "Root", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = 0 ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    lstTables.Clear
    lstTables.AddItem "TEST SET"
    lstTables.AddItem "TEST PLAN"
    lstTables.AddItem "COMPONENT"
    lstTables.AddItem "TEST INSTANCE"
    lstTables.AddItem "STEP"
    lstTables.ListIndex = 0
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestSet) To UBound(TestSet)
        lstFields.AddItem TestSet(i).Field_Name
    Next
    QCTree.Nodes.Item(2).Selected = True
    cmbModule.Enabled = False
ElseIf cmbModule.Text = "DEFECT (TEST INSTANCE)" Then
    QCTree.Nodes.Add , , "Root", "Root", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = 0 ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    lstTables.Clear
    lstTables.AddItem "TEST SET"
    lstTables.AddItem "TEST PLAN"
    lstTables.AddItem "TEST INSTANCE"
    lstTables.AddItem "DEFECT"
    lstTables.ListIndex = 0
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestSet) To UBound(TestSet)
        lstFields.AddItem Defects(i).Field_Name
    Next
    QCTree.Nodes.Item(2).Selected = True
    cmbModule.Enabled = False
ElseIf cmbModule.Text = "DEFECT (RUN)" Then
    QCTree.Nodes.Add , , "Root", "Root", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = 0 ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    lstTables.Clear
    lstTables.AddItem "TEST SET"
    lstTables.AddItem "TEST PLAN"
    lstTables.AddItem "TEST INSTANCE"
    lstTables.AddItem "RUN"
    lstTables.AddItem "DEFECT"
    lstTables.ListIndex = 0
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestSet) To UBound(TestSet)
        lstFields.AddItem TestSet(i).Field_Name
    Next
    QCTree.Nodes.Item(2).Selected = True
    cmbModule.Enabled = False
ElseIf cmbModule.Text = "DEFECT (STEP)" Then
    QCTree.Nodes.Add , , "Root", "Root", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = 0 ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    lstTables.Clear
    lstTables.AddItem "TEST SET"
    lstTables.AddItem "TEST PLAN"
    lstTables.AddItem "TEST INSTANCE"
    lstTables.AddItem "RUN"
    lstTables.AddItem "DEFECT"
    lstTables.ListIndex = 0
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestSet) To UBound(TestSet)
        lstFields.AddItem TestSet(i).Field_Name
    Next
    QCTree.Nodes.Item(2).Selected = True
    cmbModule.Enabled = False
ElseIf cmbModule.Text = "RUN STEPS" Then
    QCTree.Nodes.Add , , "Root", "Root", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = 0 ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    lstTables.Clear
    lstTables.AddItem "TEST SET"
    lstTables.AddItem "TEST PLAN"
    lstTables.AddItem "COMPONENT"
    lstTables.AddItem "TEST INSTANCE"
    lstTables.AddItem "RUN"
    lstTables.AddItem "RUN STEPS"
    lstTables.ListIndex = 0
    lstFields.Clear
    txtFilter.Text = ""
    For i = LBound(TestSet) To UBound(TestSet)
        lstFields.AddItem TestSet(i).Field_Name
    Next
    QCTree.Nodes.Item(2).Selected = True
    cmbModule.Enabled = False
End If
QCTree.Nodes(1).Expanded = True
End Sub

Private Sub Form_Load()
ClearForm
Populate_Fields
End Sub

Private Sub QCTree_DblClick()
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
Dim nodx As Node
If cmbModule.Enabled = True Then Exit Sub
If cmbModule.Text = "REQUIREMENT" Then
    If QCTree.SelectedItem.Children <> 0 Then Exit Sub
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT RQ_REQ_ID, RQ_REQ_NAME FROM REQ WHERE RQ_FATHER_ID = '" & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & "' ORDER BY RQ_REQ_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("F" & rs.FieldValue("RQ_REQ_ID")), rs.FieldValue("RQ_REQ_NAME"), 1
        rs.Next
    Next
ElseIf cmbModule.Text = "COMPONENT" Then
    If QCTree.SelectedItem.Children <> 0 Then Exit Sub
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT FC_ID, FC_NAME FROM COMPONENT_FOLDER WHERE FC_FATHER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY FC_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("F" & rs.FieldValue("FC_ID")), rs.FieldValue("FC_NAME"), 1
        rs.Next
    Next
    If Left(QCTree.SelectedItem.Key, 1) = "F" Then
        Set objCommand = QCConnection.Command
        objCommand.CommandText = "SELECT CO_ID, CO_NAME FROM COMPONENT, COMPONENT_FOLDER WHERE CO_FOLDER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " AND CO_FOLDER_ID = FC_ID ORDER BY CO_NAME"
        Set rs = objCommand.Execute
        For i = 1 To rs.RecordCount
            QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("C" & rs.FieldValue("CO_ID")), rs.FieldValue("CO_NAME"), 3
            rs.Next
        Next
    End If
ElseIf cmbModule.Text = "TEST PLAN" Then
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
ElseIf cmbModule.Text = "TEST INSTANCE" Or cmbModule.Text = "TEST SET" Or cmbModule.Text = "RUN" Or cmbModule.Text = "STEP" Then
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
ElseIf InStr(1, cmbModule.Text, "DEFECT") <> 0 Then
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
ElseIf InStr(1, cmbModule.Text, "RUN STEPS") <> 0 Then
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
End If
End Sub

Private Sub QCTree_NodeCheck(ByVal Node As MSComctlLib.Node)
Node.Selected = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpRep
Select Case Button.Key
Case "cmdRefresh"
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
    ClearForm
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
    AllScripts = ""
Case "cmdGenerate"
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
    On Error GoTo OutputErr
    AllScripts = ""
    GenerateOutput
    If SSTab1.Tab = 2 Then
        stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing report..."
        GenerateOutputToTable
    End If
    Exit Sub
OutputErr:
    MsgBox "Data was been truncated because of an error." & vbCrLf & Err.Description
Case "cmdOutput"
    If flxImport.Rows <= 1 Then
        MsgBox "Nothing to output", vbInformation
    Else
            QCConnection.SendMail "user@companyemail.com", "", "[HPQC UPDATES] Extreme Report Generated by " & curUser & " in " & curDomain & "-" & curProject, "<b>Info:</b> " & flxImport.Rows - 1 & " record(s) loaded successfully" & "<br>" & "<b>SQL Code:</b> " & txtSQL.Text, "", "HTML"
            QCConnection.SendMail curUser, "", "[HPQC UPDATES] Extreme Report Generated by " & curUser & " in " & curDomain & "-" & curProject, "<b>Info:</b> " & flxImport.Rows - 1 & " record(s) loaded successfully" & "<br>" & "<b>SQL Code:</b> " & txtSQL.Text, "", "HTML"
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
            OutputTable ColumnLetter(flxImport.Cols)
    End If
Case "cmdSave" 'HERE!!!
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
    On Error GoTo OutputErr1
    AllScripts = ""
    GenerateOutput
    tmpRep = InputBox("Enter Report Name", "Save New Report", "Extreme Report " & Format(Now, "mm dd yyyy hhmmss"))
    If Trim(tmpRep) <> "" Then
        SaveReport tmpRep
        stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = "Report saved"
    End If
Exit Sub
OutputErr1:
    MsgBox "Data was been truncated because of an error." & vbCrLf & Err.Description
End Select
End Sub

Private Sub SaveReport(ReportName)
Dim tmpRun, tmpTime, tmpTarget, tmpDays, tmpSQL, tmpPostP, tmpLogs, tmpFatherID, tmpName, tmpKey, i
Dim tmpReportFileName, stringFunct As New clsStrings, FileFunct As New clsFiles
        tmpName = stringFunct.ProperCase(CStr(ReportName))
        tmpSQL = Replace(Trim(stringFunct.RemoveAllEnter(txtSQL.Text)), ":", ";")
        FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", CStr(tmpKey), "�NAME=" & tmpName & "��SQL=" & tmpSQL & "��POSP=" & tmpPostP & "��RUN=" & tmpRun & "��TIME=" & tmpTime & "��TARGET=" & tmpTarget & "��DAYS=" & tmpDays & "��PARENTID=" & "F" & "�"
End Sub

Sub ClearForm()
With Me
    .cmbModule.Clear
    .cmbModule.AddItem "REQUIREMENT"
    .cmbModule.AddItem "COMPONENT"
    .cmbModule.AddItem "TEST PLAN"
    .cmbModule.AddItem "TEST SET"
    .cmbModule.AddItem "TEST INSTANCE"
    .cmbModule.AddItem "RUN"
    .cmbModule.AddItem "RUN STEPS"
    .cmbModule.AddItem "STEP"
    .cmbModule.AddItem "DEFECT (TEST INSTANCE)"
    .cmbModule.AddItem "DEFECT (RUN)"
    .cmbModule.AddItem "DEFECT (STEP)"
    .QCTree.Nodes.Clear
    .lstTables.Clear
    .lstFields.Clear
    .txtSQL.Text = ""
    .flxImport.Clear
    .flxImport.Rows = 2
    .flxImport.Cols = 2
    .cmbModule.Enabled = True
    .SSTab1.Tab = 0
    ReDim Extract_Fields(1)
    cmbModule.ListIndex = 0
    txtFilter.Text = ""
    lstExtract.Clear
    mdiMain.pBar.Value = 0
    mdiMain.pBar.Max = 100
End With
 Me.Caption = Me.Tag
End Sub

Private Sub OutputTable(ColLetter As String)
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim i, Protections
Dim curTab
Dim w

FileWrite App.path & "\SQC DAT" & "\" & "REPORT01" & ".xls", (CStr(AllScripts))

Set xlObject = New Excel.Application

On Error Resume Next
For Each w In xlObject.Workbooks
   w.Close savechanges:=False
Next w
On Error GoTo 0

Set xlWB = xlObject.Workbooks.Open(App.path & "\SQC DAT" & "\" & "REPORT01" & ".xls")
    xlObject.Sheets.Add
    xlObject.Sheets(1).Range("A1").Value = "Report Name: " & QCTree.SelectedItem.FullPath
    xlObject.Sheets(1).Range("A2").Value = "Code: " & txtSQL.Text
    xlObject.Sheets(1).Range("A3").Value = "Report Date: " & Format(Now, "mmm/dd/yyyy hh:mm")
    xlObject.Sheets(1).Range("A4").Value = "Environment:"
    xlObject.Sheets(1).Range("B4").Value = curDomain & "-" & curProject
    xlObject.Sheets(1).Columns("A:A").EntireColumn.AutoFit
    With xlObject.Sheets(1).Columns("A:A").Font
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
  xlObject.Sheets(1).Name = "INSTRUCTIONS"
  'xlObject.Visible = True
  'curTab = "Report01"
  'xlObject.Sheets(1).Name = curTab
'On Error Resume Next
    xlObject.Sheets(2).Select
    xlObject.Sheets(2).Range("A:" & ColLetter).Select

    xlObject.Sheets(2).Range("A:" & ColLetter).Borders(xlDiagonalDown).LineStyle = xlNone
    xlObject.Sheets(2).Range("A:" & ColLetter).Borders(xlDiagonalUp).LineStyle = xlNone
    With xlObject.Sheets(2).Range("A:" & ColLetter).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(2).Range("A:" & ColLetter).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(2).Range("A:" & ColLetter).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(2).Range("A:" & ColLetter).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(2).Range("A:" & ColLetter).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(2).Range("A:" & ColLetter).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    xlObject.Sheets(2).Rows("1:1").Select
    With xlObject.Sheets(2).Rows("1:1").Interior
        .ColorIndex = 6
        .Pattern = xlSolid
    End With
    xlObject.Sheets(2).Rows("1:1").Font.Bold = True
    xlObject.Sheets(2).Range("A:" & ColLetter).Select
    xlObject.Sheets(2).Range("A:" & ColLetter).EntireColumn.AutoFit
    xlObject.Sheets(2).Range("A1").Select

    xlObject.Sheets(2).Range("A1").AddComment
    xlObject.Sheets(2).Range("A1").Comment.Visible = False
    xlObject.Sheets(2).Range("A1").Comment.Text Text:="" & "[" & mdiMain.Caption & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""
    xlObject.Sheets(2).Protection.AllowEditRanges.Add Title:="Range1", Range:=xlObject.Sheets(2).Range("A:" & ColLetter)
  
  xlObject.Sheets(2).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  xlObject.Workbooks(1).SaveAs CleanTheString(QCTree.SelectedItem.Text & "-" & Format(Now, "mm-dd-yyyy HH-MM AMPM"))
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate
  FXGirl.EZPlay FXExportToExcel
  Set xlWB = Nothing
  Set xlObject = Nothing
  FXGirl.EZPlay FXExportToExcel
  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export to MS Excel completed.": Exit Sub:
OutErr:     MsgBox Err.Description, vbCritical: xlObject.Visible = True: xlObject.ActiveWindow.Activate: Set xlWB = Nothing: Set xlObject = Nothing
On Error GoTo 0
End Sub

Sub GenerateOutput()
Dim tmpSQL, tmpID, strPath, j, tmpFilter
tmpFilter = GetAllFilters
Select Case cmbModule.List(cmbModule.ListIndex)
Case "REQUIREMENT"
    If QCTree.SelectedItem.Index <> 1 Then
            ReDim CheckedItems_(1): strPath = ""
            GetAllCheckedItems_ QCTree.Nodes(1)
            For j = LBound(CheckedItems_) To UBound(CheckedItems_) - 1
                If Left(CheckedItems_(j), 1) = "F" Then
                    strPath = strPath & "RQ_REQ_PATH LIKE '" & GetFromTable(Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1), "RQ_REQ_ID", "RQ_REQ_PATH", "REQ") & "%' OR "
                ElseIf Left(CheckedItems_(j), 1) = "C" Then
                    'strPath = strPath & "CO_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                End If
            Next
            If Trim(strPath) <> "" Then
                strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
            Else
                MsgBox "Please select and check source(s) in the HPQC folder tree"
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                Exit Sub
            End If
      If Trim(tmpFilter) <> "" Then
        If IsTableUsed("TEST PLAN") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM REQ, TEST, REQ_COVER WHERE  RC_ENTITY_TYPE = 'TEST' AND RC_ENTITY_ID = TS_TEST_ID AND RC_REQ_ID = RQ_REQ_ID AND " & strPath & " AND " & tmpFilter
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM REQ WHERE " & strPath & " AND " & tmpFilter
        End If
      Else
        If IsTableUsed("TEST PLAN") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM REQ, TEST, REQ_COVER WHERE  RC_ENTITY_TYPE = 'TEST' AND RC_ENTITY_ID = TS_TEST_ID AND RC_REQ_ID = RQ_REQ_ID AND " & strPath
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM REQ WHERE " & strPath
        End If
      End If
    Else
      If Trim(tmpFilter) <> "" Then
        If IsTableUsed("TEST PLAN") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM REQ, TEST, REQ_COVER WHERE  RC_ENTITY_TYPE = 'TEST' AND RC_ENTITY_ID = TS_TEST_ID AND RC_REQ_ID = RQ_REQ_ID WHERE " & tmpFilter
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM REQ WHERE " & tmpFilter
        End If
      Else
        If IsTableUsed("TEST PLAN") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM REQ, TEST, REQ_COVER WHERE  RC_ENTITY_TYPE = 'TEST' AND RC_ENTITY_ID = TS_TEST_ID AND RC_REQ_ID = RQ_REQ_ID"
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM REQ"
        End If
      End If
    End If
    Debug.Print tmpSQL
Case "COMPONENT"
    If QCTree.SelectedItem.Index <> 1 Then
      'tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
            ReDim CheckedItems_(1): strPath = ""
            GetAllCheckedItems_ QCTree.Nodes(1)
            For j = LBound(CheckedItems_) To UBound(CheckedItems_) - 1
                If Left(CheckedItems_(j), 1) = "F" Then
                    strPath = strPath & "FC_PATH LIKE '" & GetFromTable(Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1), "FC_ID", "FC_PATH", "COMPONENT_FOLDER") & "%' OR "
                ElseIf Left(CheckedItems_(j), 1) = "C" Then
                    strPath = strPath & "CO_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                End If
            Next
            If Trim(strPath) <> "" Then
                strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
            Else
                MsgBox "Please select and check source(s) in the HPQC folder tree"
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                Exit Sub
            End If
      'tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "FC_ID", "FC_PATH", "COMPONENT_FOLDER") & "%"
      If Trim(tmpFilter) <> "" Then
        If IsTableUsed("STEP") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, COMPONENT_FOLDER, COMPONENT_STEP WHERE FC_ID = CO_FOLDER_ID AND CO_ID = CS_COMPONENT_ID AND " & strPath & " AND " & tmpFilter & " ORDER BY CO_FOLDER_ID, CO_ID, CS_STEP_ORDER"
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, COMPONENT_FOLDER WHERE FC_ID = CO_FOLDER_ID AND " & strPath & " AND " & tmpFilter
        End If
      Else
        If IsTableUsed("STEP") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, COMPONENT_FOLDER, COMPONENT_STEP WHERE FC_ID = CO_FOLDER_ID AND CO_ID = CS_COMPONENT_ID AND " & strPath & " ORDER BY CO_FOLDER_ID, CO_ID, CS_STEP_ORDER"
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, COMPONENT_FOLDER WHERE FC_ID = CO_FOLDER_ID AND " & strPath
        End If
      End If
    Else
      If Trim(tmpFilter) <> "" Then
        If IsTableUsed("STEP") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, COMPONENT_STEP WHERE CO_ID = CS_COMPONENT_ID AND " & tmpFilter & " ORDER BY CO_FOLDER_ID, CO_ID, CS_STEP_ORDER"
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT WHERE " & tmpFilter
        End If
      Else
        If IsTableUsed("STEP") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, COMPONENT_STEP"
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT"
        End If
      End If
    End If
    Debug.Print tmpSQL
Case "TEST PLAN"
    If QCTree.SelectedItem.Index <> 1 Then
      'tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
            ReDim CheckedItems_(1): strPath = ""
            GetAllCheckedItems_ QCTree.Nodes(1)
            For j = LBound(CheckedItems_) To UBound(CheckedItems_) - 1
                If Left(CheckedItems_(j), 1) = "F" Then
                    strPath = strPath & "AL_ABSOLUTE_PATH LIKE '" & GetFromTable(Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1), "AL_ITEM_ID", "AL_ABSOLUTE_PATH", "ALL_LISTS") & "%'" & " OR "
                ElseIf Left(CheckedItems_(j), 1) = "C" Then
                    strPath = strPath & "TS_TEST_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                End If
            Next
            If Trim(strPath) <> "" Then
                strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
            Else
                MsgBox "Please select and check source(s) in the HPQC folder tree"
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                Exit Sub
            End If
      'tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "AL_ITEM_ID", "AL_ABSOLUTE_PATH", "ALL_LISTS") & "%"
      If Trim(tmpFilter) <> "" Then
        If IsTableUsed("COMPONENT") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM ALL_LISTS RIGHT JOIN TEST on AL_ITEM_ID = TS_SUBJECT LEFT JOIN BPTEST_TO_COMPONENTS on TS_TEST_ID = BC_BPT_ID LEFT JOIN COMPONENT on BC_CO_ID = CO_ID WHERE " & strPath & " AND " & tmpFilter
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM ALL_LISTS RIGHT JOIN TEST on AL_ITEM_ID = TS_SUBJECT WHERE " & strPath & " AND " & tmpFilter
        End If
      Else
        If IsTableUsed("COMPONENT") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM ALL_LISTS RIGHT JOIN TEST on AL_ITEM_ID = TS_SUBJECT LEFT JOIN BPTEST_TO_COMPONENTS on TS_TEST_ID = BC_BPT_ID LEFT JOIN COMPONENT on BC_CO_ID = CO_ID WHERE " & strPath
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM ALL_LISTS RIGHT JOIN TEST on AL_ITEM_ID = TS_SUBJECT WHERE " & strPath
        End If
      End If
    Else
      If Trim(tmpFilter) <> "" Then
        If IsTableUsed("COMPONENT") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM ALL_LISTS RIGHT JOIN TEST on AL_ITEM_ID = TS_SUBJECT LEFT JOIN BPTEST_TO_COMPONENTS on TS_TEST_ID = BC_BPT_ID LEFT JOIN COMPONENT on BC_CO_ID = CO_ID WHERE " & tmpFilter
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM ALL_LISTS RIGHT JOIN TEST on AL_ITEM_ID = TS_SUBJECT WHERE " & tmpFilter
        End If
      Else
        If IsTableUsed("COMPONENT") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM ALL_LISTS RIGHT JOIN TEST on AL_ITEM_ID = TS_SUBJECT LEFT JOIN BPTEST_TO_COMPONENTS on TS_TEST_ID = BC_BPT_ID LEFT JOIN COMPONENT on BC_CO_ID = CO_ID"
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM ALL_LISTS RIGHT JOIN TEST on AL_ITEM_ID = TS_SUBJECT"
        End If
      End If
    End If
Case "TEST SET"
    If QCTree.SelectedItem.Index <> 1 Then
      'tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
            ReDim CheckedItems_(1): strPath = ""
            GetAllCheckedItems_ QCTree.Nodes(1)
            For j = LBound(CheckedItems_) To UBound(CheckedItems_) - 1
                If Left(CheckedItems_(j), 1) = "F" Then
                    strPath = strPath & "CF_ITEM_PATH LIKE '" & GetFromTable(Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%'" & " OR "
                ElseIf Left(CheckedItems_(j), 1) = "C" Then
                    strPath = strPath & "CY_CYCLE_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                End If
            Next
            If Trim(strPath) <> "" Then
                strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
            Else
                MsgBox "Please select and check source(s) in the HPQC folder tree"
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                Exit Sub
            End If
      'tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%"
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM CYCLE, CYCL_FOLD WHERE CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " AND " & tmpFilter
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM CYCLE, CYCL_FOLD WHERE CY_FOLDER_ID = CF_ITEM_ID AND " & strPath
      End If
    Else
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM CYCLE, CYCL_FOLD WHERE " & tmpFilter
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM CYCLE, CYCL_FOLD"
      End If
    End If
Case "TEST INSTANCE"
    If QCTree.SelectedItem.Index <> 1 Then
      'tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
            ReDim CheckedItems_(1): strPath = ""
            GetAllCheckedItems_ QCTree.Nodes(1)
            For j = LBound(CheckedItems_) To UBound(CheckedItems_) - 1
                If Left(CheckedItems_(j), 1) = "F" Then
                    strPath = strPath & "CF_ITEM_PATH LIKE '" & GetFromTable(Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%'" & " OR "
                ElseIf Left(CheckedItems_(j), 1) = "T" Then
                    strPath = strPath & "TC_TESTCYCL_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                Else
                    strPath = strPath & "TC_CYCLE_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                End If
            Next
            If Trim(strPath) <> "" Then
                strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
            Else
                MsgBox "Please select and check source(s) in the HPQC folder tree"
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                Exit Sub
            End If
      'tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%"
      If Trim(tmpFilter) <> "" Then
        If IsTableUsed("COMPONENT") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, TEST, CYCLE, TESTCYCL, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " AND " & tmpFilter
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM TEST, CYCLE, TESTCYCL, CYCL_FOLD WHERE TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " AND " & tmpFilter
        End If
      Else
        If IsTableUsed("COMPONENT") = True Then
            tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, TEST, CYCLE, TESTCYCL, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath
        Else
            tmpSQL = "SELECT " & GetAllFields & " FROM TEST, CYCLE, TESTCYCL, CYCL_FOLD WHERE TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath
        End If
      End If
    Else
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, TEST, CYCLE, TESTCYCL, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & tmpFilter
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, TEST, CYCLE, TESTCYCL, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID"
      End If
    End If
Case "RUN"
    If QCTree.SelectedItem.Index <> 1 Then
      'tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
        ReDim CheckedItems_(1): strPath = ""
            GetAllCheckedItems_ QCTree.Nodes(1)
            For j = LBound(CheckedItems_) To UBound(CheckedItems_) - 1
                If Left(CheckedItems_(j), 1) = "F" Then
                    strPath = strPath & "CF_ITEM_PATH LIKE '" & GetFromTable(Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%'" & " OR "
                ElseIf Left(CheckedItems_(j), 1) = "T" Then
                    strPath = strPath & "TC_TESTCYCL_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                Else
                    strPath = strPath & "TC_CYCLE_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                End If
            Next
            If Trim(strPath) <> "" Then
                strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
            Else
                MsgBox "Please select and check source(s) in the HPQC folder tree"
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                Exit Sub
            End If
      'tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%"
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, TEST, CYCLE, TESTCYCL, RUN, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE RN_TEST_ID = TS_TEST_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID AND RN_CYCLE_ID = CY_CYCLE_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " AND " & tmpFilter
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, TEST, CYCLE, TESTCYCL, RUN, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE RN_TEST_ID = TS_TEST_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID AND RN_CYCLE_ID = CY_CYCLE_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath
      End If
    Else
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, TEST, CYCLE, TESTCYCL, RUN, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE RN_TEST_ID = TS_TEST_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID AND RN_CYCLE_ID = CY_CYCLE_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & tmpFilter
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, TEST, CYCLE, TESTCYCL, RUN, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE RN_TEST_ID = TS_TEST_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID AND RN_CYCLE_ID = CY_CYCLE_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID"
      End If
    End If
Case "RUN STEPS"
    If QCTree.SelectedItem.Index <> 1 Then
      'tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
        ReDim CheckedItems_(1): strPath = ""
            GetAllCheckedItems_ QCTree.Nodes(1)
            For j = LBound(CheckedItems_) To UBound(CheckedItems_) - 1
                If Left(CheckedItems_(j), 1) = "F" Then
                    strPath = strPath & "CF_ITEM_PATH LIKE '" & GetFromTable(Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%'" & " OR "
                ElseIf Left(CheckedItems_(j), 1) = "T" Then
                    strPath = strPath & "TC_TESTCYCL_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                Else
                    strPath = strPath & "TC_CYCLE_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                End If
            Next
            If Trim(strPath) <> "" Then
                strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
            Else
                MsgBox "Please select and check source(s) in the HPQC folder tree"
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                Exit Sub
            End If
      'tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%"
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, TEST, CYCLE, TESTCYCL, RUN, STEP, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE RN_TEST_ID = TS_TEST_ID AND RN_RUN_ID = ST_RUN_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID AND RN_CYCLE_ID = CY_CYCLE_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " AND " & tmpFilter
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, TEST, CYCLE, TESTCYCL, RUN, STEP, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE RN_TEST_ID = TS_TEST_ID AND RN_RUN_ID = ST_RUN_ID AND  RN_TESTCYCL_ID = TC_TESTCYCL_ID AND RN_CYCLE_ID = CY_CYCLE_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath
      End If
    Else
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, TEST, CYCLE, TESTCYCL, RUN, STEP, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE RN_TEST_ID = TS_TEST_ID AND RN_RUN_ID = ST_RUN_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID AND RN_CYCLE_ID = CY_CYCLE_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & tmpFilter
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM COMPONENT, TEST, CYCLE, TESTCYCL, RUN, STEP, CYCL_FOLD, BPTEST_TO_COMPONENTS WHERE RN_TEST_ID = TS_TEST_ID AND RN_RUN_ID = ST_RUN_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID AND RN_CYCLE_ID = CY_CYCLE_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = BC_CO_ID AND TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID"
      End If
    End If
Case "STEP"
    If QCTree.SelectedItem.Index <> 1 Then
      'tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
            ReDim CheckedItems_(1): strPath = ""
            GetAllCheckedItems_ QCTree.Nodes(1)
            For j = LBound(CheckedItems_) To UBound(CheckedItems_) - 1
                If Left(CheckedItems_(j), 1) = "F" Then
                    strPath = strPath & "CF_ITEM_PATH LIKE '" & GetFromTable(Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%'" & " OR "
                ElseIf Left(CheckedItems_(j), 1) = "T" Then
                    strPath = strPath & "TC_TESTCYCL_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                Else
                    strPath = strPath & "TC_CYCLE_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                End If
            Next
            If Trim(strPath) <> "" Then
                strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
            Else
                MsgBox "Please select and check source(s) in the HPQC folder tree"
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                Exit Sub
            End If
      'tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%"
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, CYCLE, CYCL_FOLD, COMPONENT, BPTEST_TO_COMPONENTS, COMPONENT_STEP WHERE TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND CO_ID = BC_CO_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = CS_COMPONENT_ID AND " & strPath & " AND " & tmpFilter & " ORDER BY CF_ITEM_ID, CY_CYCLE, TC_TEST_ORDER, BC_ORDER, CS_STEP_ORDER"
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, CYCLE, CYCL_FOLD, COMPONENT, BPTEST_TO_COMPONENTS, COMPONENT_STEP WHERE TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND CO_ID = BC_CO_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = CS_COMPONENT_ID AND " & strPath & " ORDER BY CF_ITEM_ID, CY_CYCLE, TC_TEST_ORDER, BC_ORDER, CS_STEP_ORDER"
      End If
    Else
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, CYCLE, CYCL_FOLD, COMPONENT, BPTEST_TO_COMPONENTS, COMPONENT_STEP WHERE TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND CO_ID = BC_CO_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = CS_COMPONENT_ID AND " & tmpFilter & " ORDER BY CF_ITEM_ID, CY_CYCLE, TC_TEST_ORDER, BC_ORDER, CS_STEP_ORDER"
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, CYCLE, CYCL_FOLD, COMPONENT, BPTEST_TO_COMPONENTS, COMPONENT_STEP WHERE TC_TEST_ID = TS_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND CO_ID = BC_CO_ID AND BC_BPT_ID = TS_TEST_ID AND CO_ID = CS_COMPONENT_ID ORDER BY CF_ITEM_ID, CY_CYCLE, TC_TEST_ORDER, BC_ORDER, CS_STEP_ORDER"
      End If
    End If
Case "DEFECT (TEST INSTANCE)"
    If QCTree.SelectedItem.Index <> 1 Then
      'tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
            ReDim CheckedItems_(1): strPath = ""
            GetAllCheckedItems_ QCTree.Nodes(1)
            For j = LBound(CheckedItems_) To UBound(CheckedItems_) - 1
                If Left(CheckedItems_(j), 1) = "F" Then
                    strPath = strPath & "CF_ITEM_PATH LIKE '" & GetFromTable(Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%'" & " OR "
                ElseIf Left(CheckedItems_(j), 1) = "T" Then
                    strPath = strPath & "TC_TESTCYCL_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                Else
                    strPath = strPath & "TC_CYCLE_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                End If
            Next
            If Trim(strPath) <> "" Then
                strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
            Else
                MsgBox "Please select and check source(s) in the HPQC folder tree"
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                Exit Sub
            End If
      'tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%"
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, LINK, BUG, CYCLE, CYCL_FOLD WHERE TS_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " AND BUG.BG_BUG_ID = LINK.LN_BUG_ID AND ((LN_ENTITY_TYPE = 'TESTCYCL' AND TESTCYCL.TC_TESTCYCL_ID = LINK.LN_ENTITY_ID) OR (LN_ENTITY_TYPE = 'TEST' AND TC_TEST_ID = LN_ENTITY_ID)) AND " & tmpFilter
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, LINK, BUG, CYCLE, CYCL_FOLD WHERE TS_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " AND BUG.BG_BUG_ID = LINK.LN_BUG_ID AND ((LN_ENTITY_TYPE = 'TESTCYCL' AND TESTCYCL.TC_TESTCYCL_ID = LINK.LN_ENTITY_ID) OR (LN_ENTITY_TYPE = 'TEST' AND TC_TEST_ID = LN_ENTITY_ID))"
      End If
    Else
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, LINK, BUG, CYCLE, CYCL_FOLD WHERE TS_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND BUG.BG_BUG_ID = LINK.LN_BUG_ID AND ((LN_ENTITY_TYPE = 'TESTCYCL' AND TESTCYCL.TC_TESTCYCL_ID = LINK.LN_ENTITY_ID) OR (LN_ENTITY_TYPE = 'TEST' AND TC_TEST_ID = LN_ENTITY_ID))"
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, LINK, BUG, CYCLE, CYCL_FOLD WHERE TS_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND BUG.BG_BUG_ID = LINK.LN_BUG_ID AND ((LN_ENTITY_TYPE = 'TESTCYCL' AND TESTCYCL.TC_TESTCYCL_ID = LINK.LN_ENTITY_ID) OR (LN_ENTITY_TYPE = 'TEST' AND TC_TEST_ID = LN_ENTITY_ID))"
      End If
    End If
Case "DEFECT (RUN)"
    If QCTree.SelectedItem.Index <> 1 Then
      'tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
            ReDim CheckedItems_(1): strPath = ""
            GetAllCheckedItems_ QCTree.Nodes(1)
            For j = LBound(CheckedItems_) To UBound(CheckedItems_) - 1
                If Left(CheckedItems_(j), 1) = "F" Then
                    strPath = strPath & "CF_ITEM_PATH LIKE '" & GetFromTable(Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%'" & " OR "
                ElseIf Left(CheckedItems_(j), 1) = "T" Then
                    strPath = strPath & "TC_TESTCYCL_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                Else
                    strPath = strPath & "TC_CYCLE_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                End If
            Next
            If Trim(strPath) <> "" Then
                strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
            Else
                MsgBox "Please select and check source(s) in the HPQC folder tree"
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                Exit Sub
            End If
      'tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%"
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, LINK, BUG, RUN, CYCLE, CYCL_FOLD WHERE TEST.TS_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " AND BUG.BG_BUG_ID = LINK.LN_BUG_ID AND LN_ENTITY_TYPE = 'RUN' AND RN_RUN_ID = LN_ENTITY_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID AND " & tmpFilter
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, LINK, BUG, RUN, CYCLE, CYCL_FOLD WHERE TEST.TS_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " AND BUG.BG_BUG_ID = LINK.LN_BUG_ID AND LN_ENTITY_TYPE = 'RUN' AND RN_RUN_ID = LN_ENTITY_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID"
      End If
    Else
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, LINK, BUG, RUN, CYCLE, CYCL_FOLD WHERE TEST.TS_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND BG_BUG_ID = LINK.LN_BUG_ID AND LN_ENTITY_TYPE = 'RUN' AND RN_RUN_ID = LN_ENTITY_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID AND " & tmpFilter
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, LINK, BUG, RUN, CYCLE, CYCL_FOLD WHERE TEST.TS_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND BG_BUG_ID = LINK.LN_BUG_ID AND LN_ENTITY_TYPE = 'RUN' AND RN_RUN_ID = LN_ENTITY_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID"
      End If
    End If
Case "DEFECT (STEP)"
    If QCTree.SelectedItem.Index <> 1 Then
      'tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
            ReDim CheckedItems_(1): strPath = ""
            GetAllCheckedItems_ QCTree.Nodes(1)
            For j = LBound(CheckedItems_) To UBound(CheckedItems_) - 1
                If Left(CheckedItems_(j), 1) = "F" Then
                    strPath = strPath & "CF_ITEM_PATH LIKE '" & GetFromTable(Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%'" & " OR "
                ElseIf Left(CheckedItems_(j), 1) = "T" Then
                    strPath = strPath & "TC_TESTCYCL_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                Else
                    strPath = strPath & "TC_CYCLE_ID = " & Right(CheckedItems_(j), Len(CheckedItems_(j)) - 1) & " OR "
                End If
            Next
            If Trim(strPath) <> "" Then
                strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
            Else
                MsgBox "Please select and check source(s) in the HPQC folder tree"
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                Exit Sub
            End If
      'tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%"
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, LINK, BUG, RUN, STEP, CYCLE, CYCL_FOLD WHERE TEST.TS_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " AND BUG.BG_BUG_ID = LINK.LN_BUG_ID AND LN_ENTITY_TYPE = 'STEP' AND ST_ID = LN_ENTITY_ID AND ST_RUN_ID = RN_RUN_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID AND " & tmpFilter
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, LINK, BUG, RUN, STEP, CYCLE, CYCL_FOLD WHERE TEST.TS_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " AND BUG.BG_BUG_ID = LINK.LN_BUG_ID AND LN_ENTITY_TYPE = 'STEP' AND ST_ID = LN_ENTITY_ID AND ST_RUN_ID = RN_RUN_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID"
      End If
    Else
      If Trim(tmpFilter) <> "" Then
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, LINK, BUG, RUN, STEP, CYCLE, CYCL_FOLD WHERE TEST.TS_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND BG_BUG_ID = LN_BUG_ID AND LN_ENTITY_TYPE = 'STEP' AND ST_ID = LN_ENTITY_ID AND ST_RUN_ID = RN_RUN_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID AND " & tmpFilter
      Else
        tmpSQL = "SELECT " & GetAllFields & " FROM TEST, TESTCYCL, LINK, BUG, RUN, STEP, CYCLE, CYCL_FOLD WHERE TEST.TS_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND BG_BUG_ID = LN_BUG_ID AND LN_ENTITY_TYPE = 'STEP' AND ST_ID = LN_ENTITY_ID AND ST_RUN_ID = RN_RUN_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID"
      End If
    End If
End Select
txtSQL.Text = tmpSQL
Debug.Print tmpSQL
stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "SQL generated"
End Sub

Private Sub GenerateOutputToTable()
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i
Dim strPath
Dim iterd
Dim NewVal
Dim iter
Dim TimeSt
Dim AllF
Dim Decompile
Dim P
Dim k
Dim z
Dim curColumnName
Dim tmpF
Dim stringFunct As New clsStrings
Dim ApiFunct As New clsAPIfunctions
Dim LastFolderID As String
    AllScripts = ""
    strPath = txtSQL.Text
    AllF = "[XTRM REPORT] _" & cmbModule.Text & "-" & Format(Now, "mmm-dd-yyyy hhmmss")
    Set objCommand = QCConnection.Command

    objCommand.CommandText = Replace(strPath, ":", ";")
    Debug.Print objCommand.CommandText
    Set rs = objCommand.Execute                                                                                                                                                                                                                                                 'HERE!!!!!! <<<-------------
    If rs.RecordCount > 10000 And chkCSV.Value = Unchecked Then
        MsgBox "The records found exceeds 2500 records. It will be automatically generated as a CSV file.", vbOKOnly
        chkCSV.Value = Checked
        Exit Sub
        GenerateOutput
    End If
    ClearTable
    flxImport.Cols = rs.ColCount
    If chkCSV.Value = Checked Then
      For i = 0 To rs.ColCount - 1
          flxImport.TextMatrix(0, i) = Replace(rs.ColName(i), ";", ":")
          AllScripts = AllScripts & """" & Replace(rs.ColName(i), ";", ":") & """" & ","
          rs.Next
      Next
      AllScripts = Left(AllScripts, Len(AllScripts) - 1)
    Else
      For i = 0 To rs.ColCount - 1
          flxImport.TextMatrix(0, i) = Replace(rs.ColName(i), ";", ":")
          AllScripts = AllScripts & Replace(rs.ColName(i), ";", ":") & vbTab
          rs.Next
      Next
      AllScripts = AllScripts & vbCrLf
    End If
    rs.First
    If chkCSV.Value = Unchecked Then
        flxImport.Rows = rs.RecordCount + 1
    End If
    k = 0
    mdiMain.pBar.Max = rs.RecordCount + 3
    For i = 1 To rs.RecordCount
                If chkCSV.Value = Unchecked Then
                  k = k + 1
                  flxImport.Rows = k + 1
                End If
                For P = 0 To rs.ColCount - 1
                    curColumnName = Replace(flxImport.TextMatrix(0, P), ":", ";")
                    Select Case UCase(Trim(curColumnName))
                        Case UCase("RequirementFolderPath")
                                If rs.FieldValue("RQ_REQ_ID") <> "" Then
                                    If LastFolderID <> rs.FieldValue("RQ_REQ_ID") Then
                                        tmpF = GetRequirementFolderPath(rs.FieldValue("RQ_REQ_ID"))
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    Else
                                        If chkCSV.Value = Unchecked Then
                                          tmpF = flxImport.TextMatrix(k - 1, P)
                                        Else
                                          tmpF = tmpF
                                        End If
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    End If
                                    LastFolderID = rs.FieldValue("RQ_REQ_ID")
                                Else
                                    If LastFolderID <> rs.FieldValue("Requirement ID") Then
                                        tmpF = GetRequirementFolderPath(rs.FieldValue("Requirement ID"))
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    Else
                                        If chkCSV.Value = Unchecked Then
                                          tmpF = flxImport.TextMatrix(k - 1, P)
                                        Else
                                          tmpF = tmpF
                                        End If
                                       If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    End If
                                    LastFolderID = rs.FieldValue("Requirement ID")
                                End If
                        Case UCase("TestSetFolderPath")
                                If rs.FieldValue("CY_CYCLE_ID") <> "" Then
                                    If LastFolderID <> rs.FieldValue("CY_CYCLE_ID") Then
                                        tmpF = GetTestSetFolderPath(rs.FieldValue("CY_CYCLE_ID"))
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    Else
                                        If chkCSV.Value = Unchecked Then
                                          tmpF = flxImport.TextMatrix(k - 1, P)
                                        Else
                                          tmpF = tmpF
                                        End If
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    End If
                                    LastFolderID = rs.FieldValue("CY_CYCLE_ID")
                                Else
                                    If LastFolderID <> rs.FieldValue("Test Set ID") Then
                                        tmpF = GetTestSetFolderPath(rs.FieldValue("Test Set ID"))
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    Else
                                        If chkCSV.Value = Unchecked Then
                                          tmpF = flxImport.TextMatrix(k - 1, P)
                                        Else
                                          tmpF = tmpF
                                        End If
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    End If
                                    LastFolderID = rs.FieldValue("Test Set ID")
                                End If
                        Case UCase("BusinessComponentFolderPath")
                                If rs.FieldValue("CO_ID") <> "" Then
                                    If LastFolderID <> rs.FieldValue("CO_ID") Then
                                        tmpF = GetBusinessComponentFolderPath(rs.FieldValue("CO_ID"))
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    Else
                                        If chkCSV.Value = Unchecked Then
                                          tmpF = flxImport.TextMatrix(k - 1, P)
                                        Else
                                          tmpF = tmpF
                                        End If
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    End If
                                    LastFolderID = rs.FieldValue("CO_ID")
                                Else
                                    If LastFolderID <> rs.FieldValue("Component ID") Then
                                        tmpF = GetBusinessComponentFolderPath(rs.FieldValue("Component ID"))
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    Else
                                        If chkCSV.Value = Unchecked Then
                                          tmpF = flxImport.TextMatrix(k - 1, P)
                                        Else
                                          tmpF = tmpF
                                        End If
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    End If
                                    LastFolderID = rs.FieldValue("Component ID")
                                End If
                        Case UCase("TestFolderPath")
                                If rs.FieldValue("TS_SUBJECT") <> "" Then
                                    If LastFolderID <> rs.FieldValue("TS_SUBJECT") Then
                                        tmpF = GetTestFolderPath(rs.FieldValue("TS_SUBJECT"))
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    Else
                                        If chkCSV.Value = Unchecked Then
                                          tmpF = flxImport.TextMatrix(k - 1, P)
                                        Else
                                          tmpF = tmpF
                                        End If
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    End If
                                    LastFolderID = rs.FieldValue("TS_SUBJECT")
                                Else
                                    If LastFolderID <> rs.FieldValue("Subject") Then
                                        tmpF = GetTestFolderPath(rs.FieldValue("Subject"))
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    Else
                                        If chkCSV.Value = Unchecked Then
                                          tmpF = flxImport.TextMatrix(k - 1, P)
                                        Else
                                          tmpF = tmpF
                                        End If
                                        If chkCSV.Value = Checked Then
                                          If P = 0 Then
                                            AllScripts = AllScripts & vbCrLf & """" & (tmpF) & """" & ","
                                          Else
                                            AllScripts = AllScripts & """" & (tmpF) & """" & ","
                                          End If
                                        Else
                                          AllScripts = AllScripts & (tmpF) & vbTab
                                          flxImport.TextMatrix(k, P) = tmpF
                                        End If
                                    End If
                                    LastFolderID = rs.FieldValue("Subject")
                                End If
                        Case Else
                            If chkCSV.Value = Unchecked Then flxImport.TextMatrix(k, P) = ReverseCleanHTML(rs.FieldValue(curColumnName))
                            If chkCSV.Value = Checked Then
                              If stringFunct.StrIn(rs.FieldValue(curColumnName), vbCrLf) = True Or stringFunct.StrIn(rs.FieldValue(curColumnName), Chr(10) + Chr(13)) = True Or stringFunct.StrIn(rs.FieldValue(curColumnName), Chr(10)) = True Or stringFunct.StrIn(rs.FieldValue(curColumnName), Chr(13)) = True Or stringFunct.StrIn(rs.FieldValue(curColumnName), vbTab) = True Then
                                      If P = 0 And AllScripts <> "" Then
                                        AllScripts = AllScripts & vbCrLf & """" & Replace(ReverseCleanHTML(rs.FieldValue(curColumnName)), vbTab, " ") & """" & ","
                                      Else
                                        AllScripts = AllScripts & """" & Replace(ReverseCleanHTML(rs.FieldValue(curColumnName)), vbTab, " ") & """" & ","
                                      End If
                              Else
                                      If P = 0 And AllScripts <> "" Then
                                        AllScripts = AllScripts & vbCrLf & """" & ReverseCleanHTML(rs.FieldValue(curColumnName)) & """" & ","
                                      Else
                                        AllScripts = AllScripts & """" & ReverseCleanHTML(rs.FieldValue(curColumnName)) & """" & ","
                                      End If
                              End If
                            Else
                              If stringFunct.StrIn(rs.FieldValue(curColumnName), vbCrLf) = True Or stringFunct.StrIn(rs.FieldValue(curColumnName), Chr(10) + Chr(13)) = True Or stringFunct.StrIn(rs.FieldValue(curColumnName), Chr(10)) = True Or stringFunct.StrIn(rs.FieldValue(curColumnName), Chr(13)) = True Or stringFunct.StrIn(rs.FieldValue(curColumnName), vbTab) = True Then
                                      AllScripts = AllScripts & Replace(ReverseCleanHTML(rs.FieldValue(curColumnName)), vbTab, " ") & vbTab
                              Else
                                      AllScripts = AllScripts & ReverseCleanHTML(rs.FieldValue(curColumnName)) & vbTab
                              End If
                            End If
                        End Select
                Next
                If chkCSV.Value = Unchecked Then AllScripts = AllScripts & vbCrLf
                rs.Next
                If i Mod 500 = 0 Then
                    Me.Refresh
                    stsBar.Refresh
                    ApiFunct.Pause 0.1
                    If chkCSV.Value = Checked Then
                      FileAppend App.path & "\SQC Logs" & "\" & AllF & ".csv", AllScripts
                      AllScripts = ""
                    End If
                End If
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Processing " & i & " out of " & rs.RecordCount
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
    FXGirl.EZPlay FXSQCExtractCompleted
    If chkCSV.Value = Checked Then '***
        FileAppend App.path & "\SQC Logs" & "\" & AllF & ".csv", AllScripts: If MsgBox("Successfully exported to " & App.path & "\SQC Logs" & "\" & AllF & ".csv" & vbCrLf & "Do you want to launch the extracted file?", vbYesNo) = vbYes Then Shell "explorer.exe " & App.path & "\SQC Logs" & "\", vbNormalFocus
        AllScripts = vbCrLf & " ,"
        AllScripts = AllScripts & """" & objCommand.CommandText & """"
        FileAppend App.path & "\SQC Logs" & "\" & AllF & ".csv", AllScripts
        AllScripts = ""
        QCConnection.SendMail "user@companyemail.com", "", "[HPQC UPDATES] Extreme Report Generated by " & curUser & " in " & curDomain & "-" & curProject, "<b>Info:</b> " & rs.RecordCount & " record(s) loaded successfully" & "<br>" & "<b>SQL Code:</b> " & txtSQL.Text, "", "HTML"
        QCConnection.SendMail curUser, "", "[HPQC UPDATES] Extreme Report Generated by " & curUser & " in " & curDomain & "-" & curProject, "<b>Info:</b> " & rs.RecordCount & " record(s) loaded successfully" & "<br>" & "<b>SQL Code:</b> " & txtSQL.Text, "", "HTML"
    End If '***
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = rs.RecordCount & " record(s) loaded successfully"
End Sub

Private Sub ClearTable()
flxImport.Clear
End Sub

Function IsTableUsed(TableName As String) As Boolean
Dim i
For i = LBound(Extract_Fields) To UBound(Extract_Fields)
    If UCase(Trim(Extract_Fields(i).Table_Name)) = UCase(Trim(TableName)) Then
       IsTableUsed = True
       Exit Function
    End If
Next
End Function

Function GetAllFields()
Dim i
Dim tmp
For i = LBound(Extract_Fields) To UBound(Extract_Fields)
    If Extract_Fields(i).IsSpecial = False Then
       tmp = tmp & " " & Extract_Fields(i).Technical_Name & " AS """ & Extract_Fields(i).Field_Name & """" & ","
    Else
       tmp = tmp & " " & "''" & " AS """ & Extract_Fields(i).Technical_Name & """" & ","
    End If
Next
tmp = Left(tmp, Len(tmp) - 1)
GetAllFields = Trim(tmp)
End Function

Function GetAllFilters()
Dim i
Dim tmp
On Error Resume Next
For i = LBound(Extract_Fields) To UBound(Extract_Fields)
    If Trim(Extract_Fields(i).Filter_Value) <> "" Then
        If InStr(1, Extract_Fields(i).Filter_Value, "%") <> 0 Then
            If Trim(UCase(Extract_Fields(i).Filter_Value)) = "%BLANK%" Then
                tmp = tmp & " " & Extract_Fields(i).Technical_Name & " = '' AND "
            Else
                tmp = tmp & " " & Extract_Fields(i).Technical_Name & " LIKE '" & Extract_Fields(i).Filter_Value & "' AND "
            End If
        Else
            tmp = tmp & " " & Extract_Fields(i).Technical_Name & " = '" & Extract_Fields(i).Filter_Value & "' AND "
        End If
    End If
Next
tmp = Left(tmp, Len(tmp) - 5)
GetAllFilters = Trim(tmp)
End Function

' @Function Array_RemoveItem
' -----------------------------
'@Author Roland Ross Hadi
'@Description Removes array element
'@Comments
' ItemArray - String to be processed
' ItemElement - Index
Private Sub Array_RemoveItem_1(ItemElement As Integer)
Dim lCtr As Long
Dim lTop As Long
Dim lBottom As Long


lTop = UBound(Extract_Fields)
lBottom = LBound(Extract_Fields)

If ItemElement < lBottom Or ItemElement > lTop Then
    Err.Raise 9, , "Subscript out of Range"
    Exit Sub
End If

For lCtr = ItemElement To lTop - 1
    Extract_Fields(lCtr) = Extract_Fields(lCtr + 1)
Next
On Error GoTo ErrorHandler:

ReDim Preserve Extract_Fields(lBottom To lTop - 1)

Exit Sub
ErrorHandler:
  'An error will occur if array is fixed
    Err.Raise Err.Number, , _
       "You must pass a resizable array to this function"
End Sub
' Function RemoveAllEnter
' -----------------------------

Private Sub txtFilter_Change()
If lstFields.ListIndex <> -1 Then
    Select Case lstTables.List(lstTables.ListIndex)
        Case "REQUIREMENT"
            AddFilter Requirements(lstFields.ListIndex + 1).Table_Name, Requirements(lstFields.ListIndex + 1).Field_Name, Requirements(lstFields.ListIndex + 1).Technical_Name, Trim(txtFilter.Text)
        Case "COMPONENT"
            AddFilter BusinessComponent(lstFields.ListIndex + 1).Table_Name, BusinessComponent(lstFields.ListIndex + 1).Field_Name, BusinessComponent(lstFields.ListIndex + 1).Technical_Name, Trim(txtFilter.Text)
        Case "TEST PLAN"
            AddFilter TestPlan(lstFields.ListIndex + 1).Table_Name, TestPlan(lstFields.ListIndex + 1).Field_Name, TestPlan(lstFields.ListIndex + 1).Technical_Name, Trim(txtFilter.Text)
        Case "TEST SET"
            AddFilter TestSet(lstFields.ListIndex + 1).Table_Name, TestSet(lstFields.ListIndex + 1).Field_Name, TestSet(lstFields.ListIndex + 1).Technical_Name, Trim(txtFilter.Text)
        Case "TEST INSTANCE"
            AddFilter TestInstance(lstFields.ListIndex + 1).Table_Name, TestInstance(lstFields.ListIndex + 1).Field_Name, TestInstance(lstFields.ListIndex + 1).Technical_Name, Trim(txtFilter.Text)
        Case "RUN"
            AddFilter TestRun(lstFields.ListIndex + 1).Table_Name, TestRun(lstFields.ListIndex + 1).Field_Name, TestRun(lstFields.ListIndex + 1).Technical_Name, Trim(txtFilter.Text)
        Case "STEP"
            AddFilter Step(lstFields.ListIndex + 1).Table_Name, Step(lstFields.ListIndex + 1).Field_Name, Step(lstFields.ListIndex + 1).Technical_Name, Trim(txtFilter.Text)
        Case "DEFECT"
            AddFilter Defects(lstFields.ListIndex + 1).Table_Name, Defects(lstFields.ListIndex + 1).Field_Name, Defects(lstFields.ListIndex + 1).Technical_Name, Trim(txtFilter.Text)
        Case "RUN STEPS"
            AddFilter RunSteps(lstFields.ListIndex + 1).Table_Name, RunSteps(lstFields.ListIndex + 1).Field_Name, RunSteps(lstFields.ListIndex + 1).Technical_Name, Trim(txtFilter.Text)
    End Select
    Populate_Order
End If
End Sub

Private Sub txtSQL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 67 And Shift = vbCtrlMask Then
    Clipboard.Clear
    Clipboard.SetText txtSQL.Text
End If
End Sub

Private Sub UpDown_UpClick()
Dim tmpFields As HPQC_FIELDS, tmpList As String
If lstExtract.ListIndex <= 0 Then Exit Sub
tmpList = lstExtract.List(lstExtract.ListIndex)
lstExtract.List(lstExtract.ListIndex) = lstExtract.List(lstExtract.ListIndex - 1)
lstExtract.List(lstExtract.ListIndex - 1) = tmpList
tmpFields = Extract_Fields(lstExtract.ListIndex + 1)
Extract_Fields(lstExtract.ListIndex + 1) = Extract_Fields(lstExtract.ListIndex - 1 + 1)
Extract_Fields(lstExtract.ListIndex - 1 + 1) = tmpFields
lstExtract.ListIndex = lstExtract.ListIndex - 1
End Sub

Private Sub UpDown_DownClick()
Dim tmpFields As HPQC_FIELDS, tmpList As String
If lstExtract.ListIndex = lstExtract.ListCount - 1 Or lstExtract.ListIndex = -1 Then Exit Sub
tmpList = lstExtract.List(lstExtract.ListIndex)
lstExtract.List(lstExtract.ListIndex) = lstExtract.List(lstExtract.ListIndex + 1)
lstExtract.List(lstExtract.ListIndex + 1) = tmpList
tmpFields = Extract_Fields(lstExtract.ListIndex + 1)
Extract_Fields(lstExtract.ListIndex + 1) = Extract_Fields(lstExtract.ListIndex + 1 + 1)
Extract_Fields(lstExtract.ListIndex + 1 + 1) = tmpFields
lstExtract.ListIndex = lstExtract.ListIndex + 1
End Sub

Private Sub GetAllCheckedItems_(objNode As Node)
    Dim objSiblingNode As Node
    Set objSiblingNode = objNode
Do
     If objSiblingNode.Checked = True Then
        CheckedItems_(UBound(CheckedItems_)) = objSiblingNode.Key
        ReDim Preserve CheckedItems_(UBound(CheckedItems_) + 1)
     End If
     If Not objSiblingNode.Child Is Nothing Then
         Call GetAllCheckedItems_(objSiblingNode.Child)
     End If
     Set objSiblingNode = objSiblingNode.Next
Loop While Not objSiblingNode Is Nothing
End Sub
