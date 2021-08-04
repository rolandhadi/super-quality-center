VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReports 
   Caption         =   "Report Generator"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12840
   Icon            =   "frmReports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   12840
   Tag             =   "Report Generator"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1380
      MaskColor       =   &H000000FF&
      Picture         =   "frmReports.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3960
      Width           =   1395
   End
   Begin VB.CommandButton cmdReport 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4140
      MaskColor       =   &H000000FF&
      Picture         =   "frmReports.frx":0E1D
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3960
      Width           =   1395
   End
   Begin VB.CommandButton cmdNewFolder 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2760
      MaskColor       =   &H000000FF&
      Picture         =   "frmReports.frx":161B
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3960
      Width           =   1395
   End
   Begin TabDlg.SSTab TabSQL 
      Height          =   3615
      Left            =   5580
      TabIndex        =   3
      Top             =   660
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Scheduling"
      TabPicture(0)   =   "frmReports.frx":1F77
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmeRun"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmeSched"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtReportName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "SQL Code"
      TabPicture(1)   =   "frmReports.frx":1F93
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtSQL"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Post Process"
      TabPicture(2)   =   "frmReports.frx":1FAF
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtPostProcess"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Logs"
      TabPicture(3)   =   "frmReports.frx":1FCB
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtLogs"
      Tab(3).ControlCount=   1
      Begin VB.TextBox txtPostProcess 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3195
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   360
         Width           =   2835
      End
      Begin VB.TextBox txtLogs 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3195
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   360
         Width           =   2835
      End
      Begin VB.TextBox txtReportName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3060
         TabIndex        =   27
         Top             =   1140
         Width           =   3975
      End
      Begin VB.Frame frmeSched 
         Caption         =   "Scheduling"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   180
         TabIndex        =   8
         Top             =   1500
         Width           =   6915
         Begin VB.ComboBox lstMinute 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   300
            Width           =   615
         End
         Begin VB.ComboBox lstHour 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   300
            Width           =   615
         End
         Begin VB.CommandButton cmdPath 
            Height          =   435
            Left            =   6300
            Picture         =   "frmReports.frx":1FE7
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtSavePath 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   780
            Width           =   4575
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Sunday"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   5340
            TabIndex        =   15
            Top             =   1200
            Width           =   1035
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Saturday"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   3960
            TabIndex        =   14
            Top             =   1560
            Width           =   1035
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Friday"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   3960
            TabIndex        =   13
            Top             =   1200
            Width           =   1035
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Thursday"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   2220
            TabIndex        =   12
            Top             =   1560
            Width           =   1155
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Wednesday"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2220
            TabIndex        =   11
            Top             =   1200
            Width           =   1275
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Tuesday"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   660
            TabIndex        =   10
            Top             =   1560
            Width           =   1035
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Monday"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   660
            TabIndex        =   9
            Top             =   1200
            Width           =   1035
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2340
            TabIndex        =   20
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Time:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Target Location:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame frmeRun 
         Caption         =   "Run Scheduled Execution"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   180
         TabIndex        =   5
         Top             =   480
         Width           =   2175
         Begin VB.OptionButton optNo 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   180
            TabIndex        =   7
            Top             =   540
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optYes 
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   180
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
      End
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
         Height          =   3195
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   360
         Width           =   2835
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2460
         TabIndex        =   26
         Top             =   1200
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12840
      _ExtentX        =   22648
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
            Object.ToolTipText     =   "Save"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin VB.CheckBox chkCSV 
         Caption         =   "Download to CSV"
         Height          =   315
         Left            =   2040
         TabIndex        =   31
         Top             =   120
         Width           =   2655
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11640
      Top             =   5880
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
            Picture         =   "frmReports.frx":2680
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReports.frx":2912
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReports.frx":2BA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   12240
      Top             =   5880
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
            Picture         =   "frmReports.frx":2E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReports.frx":3544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReports.frx":3C56
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReports.frx":4368
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
   Begin MSComctlLib.TreeView QCTree 
      Height          =   3255
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5741
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      SingleSel       =   -1  'True
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
      Height          =   1995
      Left            =   120
      TabIndex        =   1
      Top             =   4380
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   3519
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowUserResizing=   3
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   6435
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmReports.frx":4A7A
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   21423
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
            Picture         =   "frmReports.frx":4FCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReports.frx":52AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReports.frx":57FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReports.frx":5D4F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AllScripts
Dim FileFunct As New clsFiles
Dim stringFunct As New clsStrings
Dim ApiFunct As New clsAPIfunctions

Private Sub ReportChanged()
Dim i
With Me
    .optNo.ForeColor = vbRed
    .optYes.ForeColor = vbRed
    .txtReportName.ForeColor = vbRed
    .lstHour.ForeColor = vbRed
    .lstMinute.ForeColor = vbRed
    For i = 0 To 6
        chkDay(i).ForeColor = vbRed
    Next
    .txtSavePath.ForeColor = vbRed
    .txtSQL.ForeColor = vbRed
    .txtPostProcess.ForeColor = vbRed
End With
End Sub

Private Sub ReportUpdated()
Dim i
With Me
    .optNo.ForeColor = vbBlack
    .optYes.ForeColor = vbBlack
    .txtReportName.ForeColor = vbBlack
    .lstHour.ForeColor = vbBlack
    .lstMinute.ForeColor = vbBlack
    For i = 0 To 6
        chkDay(i).ForeColor = vbBlack
    Next
    .txtSavePath.ForeColor = vbBlack
    .txtSQL.ForeColor = vbBlack
    .txtPostProcess.ForeColor = vbBlack
End With
End Sub

Private Sub GenerateOutput()
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
    AllF = "[SQC REPORT] _" & txtReportName.Text & "-" & Format(Now, "mmm-dd-yyyy hhmmss")
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
    End If '***
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = rs.RecordCount & " record(s) loaded successfully"
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

Private Sub chkDay_Click(Index As Integer)
ReportChanged
End Sub

Private Sub cmdDelete_Click()
Dim tmpKey, tmpName, i
On Error Resume Next
If MsgBox("Are you sure you want to delete this item?", vbYesNo) = vbYes Then
    tmpKey = QCTree.SelectedItem.Key
    If tmpKey = "" Then
         MsgBox "Please select a parent folder"
    ElseIf stringFunct.StrIn(tmpKey, "R") = True Then
         QCTree.Nodes.Remove tmpKey
         FileFunct.DeleteFromFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "~" & CStr(tmpKey) & "~"
         FileFunct.FileDelete App.path & "\SQC Logs" & "\" & CStr(tmpKey) & ".txt"
            txtReportName.Text = ""
            optNo.Value = True
            lstHour.ListIndex = 11
            lstMinute.ListIndex = 0
            txtSavePath.Text = App.path & "\SQC Reports" & "\"
            For i = 0 To 6
                chkDay(i).Value = Unchecked
            Next
            txtSQL.Text = ""
            txtLogs.Text = ""
            txtPostProcess.Text = ""
            TabSQL.Enabled = False
            ClearTable
            AllScripts = ""
            
            QCTree.Refresh
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Deleted successfully"
    ElseIf stringFunct.StrIn(tmpKey, "F") = True Then
         If QCTree.Nodes(tmpKey).Children <> 0 Then
            MsgBox "Please remove all reports before removing the folder"
         Else
            QCTree.Nodes.Remove tmpKey
            FileFunct.DeleteFromFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "|" & CStr(tmpKey) & "|"
            
                txtReportName.Text = ""
                optNo.Value = True
                lstHour.ListIndex = 11
                lstMinute.ListIndex = 0
                txtSavePath.Text = App.path & "\SQC Reports" & "\"
                For i = 0 To 6
                    chkDay(i).Value = Unchecked
                Next
                txtSQL.Text = ""
                txtLogs.Text = ""
                txtPostProcess.Text = ""
                TabSQL.Enabled = False
                ClearTable
                AllScripts = ""
                
                QCTree.Refresh
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = "Deleted successfully"
         End If
    End If
End If
End Sub

Private Sub cmdNewFolder_Click()
Dim tmpKey, tmpName, FatherKey, i
On Error Resume Next
FatherKey = QCTree.SelectedItem.Key
If FatherKey = "" Then
     MsgBox "Please select a parent folder"
ElseIf stringFunct.StrIn(FatherKey, "R") = True Then
 
Else
     tmpName = stringFunct.OnlyFolderFriendly(InputBox("Enter new Folder Name"))
     If Trim(tmpName) <> "" Then
          tmpKey = FatherKey & "-" & Format(QCTree.Nodes.Item(FatherKey).Children + 1, "000")
          QCTree.Nodes.Add FatherKey, tvwChild, tmpKey, stringFunct.ProperCase(CStr(tmpName)), 1
          QCTree.Nodes.Item(FatherKey).Expanded = True
          QCTree.Nodes.Item(FatherKey).Selected = True
          FileFunct.ReadAndWriteFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "|" & CStr(tmpKey) & "|", "�NAME=" & stringFunct.ProperCase(CStr(tmpName)) & "��PARENTID=" & CStr(FatherKey) & "�"
            
            txtReportName.Text = ""
            optNo.Value = True
            lstHour.ListIndex = 11
            lstMinute.ListIndex = 0
            txtSavePath.Text = App.path & "\SQC Reports" & "\"
            For i = 0 To 6
                chkDay(i).Value = Unchecked
            Next
            txtSQL.Text = ""
            txtLogs.Text = ""
            txtPostProcess.Text = ""
            TabSQL.Enabled = False
            ClearTable
            AllScripts = ""
            
            QCTree.Refresh
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = "New folder created successfully"
     End If
End If
End Sub

Private Sub cmdNewFolder_Funct(strName, strParent, tmpKey)
Dim tmpName, FatherKey
On Error Resume Next
FatherKey = strParent
If FatherKey = "" Then
     MsgBox "Please select a parent folder"
ElseIf stringFunct.StrIn(FatherKey, "R") = True Then
 
Else
     tmpName = strName
     If Trim(tmpName) <> "" Then
          QCTree.Nodes.Add FatherKey, tvwChild, tmpKey, stringFunct.ProperCase(CStr(tmpName)), 1
     End If
End If
End Sub

Private Sub cmdPath_Click()
Dim tmp
    tmp = BrowseForFolder(hWnd, "Please select a target folder.")
    If tmp <> "" Then txtSavePath.Text = tmp
End Sub

Private Sub cmdReport_Click()
Dim tmpKey, tmpName, FatherKey
On Error Resume Next
FatherKey = QCTree.SelectedItem.Key
If FatherKey = "" Then
     MsgBox "Please select a parent folder"
ElseIf stringFunct.StrIn(FatherKey, "R") = True Then
     
Else
     tmpName = stringFunct.OnlyFolderFriendly(InputBox("Enter new Report Name"))
     If Trim(tmpName) <> "" Then
        tmpKey = FatherKey & "-R" & Format(QCTree.Nodes.Item(FatherKey).Children + 1, "000")
          QCTree.Nodes.Add FatherKey, tvwChild, tmpKey, stringFunct.ProperCase(CStr(tmpName)), 2
          QCTree.Nodes.Item(FatherKey).Expanded = True
          QCTree.Nodes.Item(FatherKey).Selected = True
          FileFunct.ReadAndWriteFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "~" & CStr(tmpKey) & "~", "�NAME=" & stringFunct.ProperCase(CStr(tmpName)) & "��SQL=��POSTP=��RUN=NO��TIME=12:00��TARGET=" & App.path & "\SQC Reports" & "\" & "��DAYS=0000000��PARENTID=" & CStr(FatherKey) & "�"
          stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = "New report created successfully"
          QCTree.Refresh
     End If
End If
End Sub

Private Sub cmdReport_Funct(tmpName, tmpParent, tmpKey)
Dim FatherKey
On Error Resume Next
FatherKey = tmpParent
If FatherKey = "" Then
     MsgBox "Please select a parent folder"
ElseIf stringFunct.StrIn(FatherKey, "R") = True Then
     
Else
     If Trim(tmpName) <> "" Then
          QCTree.Nodes.Add FatherKey, tvwChild, tmpKey, stringFunct.ProperCase(CStr(tmpName)), 2
     End If
End If
End Sub

Private Sub LoadFolders()
Dim tmpContent, i, tmp, j
Dim tmpName(), tmpParent(), tmpKey()
tmpContent = FileFunct.ReadFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh")
If tmpContent = "No Data Found" Then Exit Sub
tmp = Split(tmpContent, "|")
ReDim tmpName(0)
ReDim tmpParent(0)
ReDim tmpKey(0)
For i = LBound(tmp) To UBound(tmp)
    If Left(tmp(i), 1) = "F" Then
        tmpName(UBound(tmpName)) = stringFunct.GetValueFromKey(CStr(tmp(i + 1)), "NAME")
        tmpParent(UBound(tmpParent)) = stringFunct.GetValueFromKey(CStr(tmp(i + 1)), "PARENTID")
        tmpKey(UBound(tmpKey)) = tmp(i)
        ReDim Preserve tmpName(UBound(tmpName) + 1)
        ReDim Preserve tmpParent(UBound(tmpParent) + 1)
        ReDim Preserve tmpKey(UBound(tmpKey) + 1)
    End If
Next
    If UBound(tmpName) <> 0 Then
        ReDim Preserve tmpName(UBound(tmpName) - 1)
        ReDim Preserve tmpParent(UBound(tmpParent) - 1)
        ReDim Preserve tmpKey(UBound(tmpKey) - 1)

        BubbleSort_QC_Tree tmpName, tmpParent, tmpKey
        For j = 1 To 50 * 2
                For i = LBound(tmpName) To UBound(tmpName)
                    If Len(tmpKey(i)) = j Then
                        cmdNewFolder_Funct tmpName(i), tmpParent(i), tmpKey(i)
                    End If
                Next
        Next
    End If
    
QCTree.Refresh
End Sub

Private Sub LoadReports()
Dim tmpContent, i, tmp, j
Dim tmpName(), tmpParent(), tmpKey()
tmpContent = FileFunct.ReadFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh")
If tmpContent = "No Data Found" Then Exit Sub
tmp = Split(tmpContent, "~")
ReDim tmpName(0)
ReDim tmpParent(0)
ReDim tmpKey(0)
For i = LBound(tmp) To UBound(tmp)
    If Left(tmp(i), 1) = "F" And stringFunct.StrIn(tmp(i), "R") Then
        tmpName(UBound(tmpName)) = stringFunct.GetValueFromKey(CStr(tmp(i + 1)), "NAME")
        tmpParent(UBound(tmpParent)) = stringFunct.GetValueFromKey(CStr(tmp(i + 1)), "PARENTID")
        tmpKey(UBound(tmpKey)) = tmp(i)
        ReDim Preserve tmpName(UBound(tmpName) + 1)
        ReDim Preserve tmpParent(UBound(tmpParent) + 1)
        ReDim Preserve tmpKey(UBound(tmpKey) + 1)
    End If
Next
        On Error Resume Next
        ReDim Preserve tmpName(UBound(tmpName) - 1)
        ReDim Preserve tmpParent(UBound(tmpParent) - 1)
        ReDim Preserve tmpKey(UBound(tmpKey) - 1)
        
    BubbleSort_QC_Tree tmpName, tmpParent, tmpKey
    For j = 1 To 50 * 2
            For i = LBound(tmpName) To UBound(tmpName)
                If Len(tmpKey(i)) = j Then
                    cmdReport_Funct tmpName(i), tmpParent(i), tmpKey(i)
                End If
            Next
    Next
        On Error GoTo 0
QCTree.Refresh
End Sub

Private Sub flxImport_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 67 And Shift = vbCtrlMask Then
    Clipboard.Clear
    Clipboard.SetText flxImport.Clip
End If
End Sub



Private Sub Form_Resize()
On Error Resume Next
flxImport.height = stsBar.Top - flxImport.Top - 250
flxImport.width = Me.width - flxImport.Left - 350
TabSQL.width = Me.width - TabSQL.Left - 350
txtSQL.width = Me.width - TabSQL.Left - 650
txtLogs.width = Me.width - TabSQL.Left - 650
txtPostProcess.width = Me.width - TabSQL.Left - 650
TabSQL.TabMaxWidth = 2500
End Sub


Private Sub lstHour_Change()
ReportChanged
End Sub

Private Sub lstMinute_Change()
ReportChanged
End Sub

Private Sub optNo_Click()
ReportChanged
End Sub

Private Sub optYes_Click()
ReportChanged
End Sub

Private Sub QCTree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim tmpKey, tmpName, tmpContent, tmp, i
On Error Resume Next
tmpKey = QCTree.SelectedItem.Key
If stringFunct.StrIn(tmpKey, "R") = True Then
     txtReportName.Text = Node.Text
     tmpContent = FileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "~" & CStr(tmpKey) & "~")
     tmp = stringFunct.GetValueFromKey(CStr(tmpContent), "RUN")
     If UCase(tmp) = "YES" Then
        optYes.Value = True
     Else
        optNo.Value = True
     End If
     tmp = stringFunct.GetValueFromKey(CStr(tmpContent), "TIME")
     lstHour.ListIndex = Val(Left(tmp, 2)) - 1
     lstMinute.ListIndex = Val(Right(tmp, 2))
     tmp = stringFunct.GetValueFromKey(CStr(tmpContent), "TARGET")
     txtSavePath.Text = tmp
     tmp = stringFunct.GetValueFromKey(CStr(tmpContent), "DAYS")
     For i = 0 To 6
        If Mid(tmp, i + 1, 1) = 1 Then
            chkDay(i).Value = Checked
        Else
            chkDay(i).Value = Unchecked
        End If
     Next
     tmp = stringFunct.GetValueFromKey(CStr(tmpContent), "SQL")
     txtSQL.Text = Replace(tmp, ";", ":")
     tmp = stringFunct.GetValueFromKey(CStr(tmpContent), "POSP")
     txtPostProcess.Text = tmp
     txtLogs.Text = FileFunct.ReadFromFile(App.path & "\SQC Logs\" & CStr(tmpKey) & ".txt")
     TabSQL.Enabled = True
     TabSQL.Tab = 0
     ReportUpdated
Else
    txtReportName.Text = ""
    optNo.Value = True
    lstHour.ListIndex = 11
    lstMinute.ListIndex = 0
    txtSavePath.Text = App.path & "\SQC Reports" & "\"
    For i = 0 To 6
        chkDay(i).Value = Unchecked
    Next
    txtSQL.Text = ""
    txtPostProcess.Text = ""
    txtLogs.Text = ""
    TabSQL.Enabled = False
    ClearTable
    AllScripts = ""
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = ""
    ReportUpdated
End If
End Sub

Private Sub SaveReport()
Dim tmpRun, tmpTime, tmpTarget, tmpDays, tmpSQL, tmpPostP, tmpLogs, tmpFatherID, tmpName, tmpKey, i
Dim tmpReportFileName
On Error Resume Next
If stringFunct.StrIn(QCTree.SelectedItem.Key, "R") = True Then
        If Err.Number = 91 Then Exit Sub
        On Error GoTo 0
        tmpKey = "~" & QCTree.SelectedItem.Key & "~"
        tmpFatherID = QCTree.SelectedItem.Parent.Key
        If Trim(txtReportName.Text) <> "" Then
            tmpName = stringFunct.ProperCase(CStr(txtReportName.Text))
            QCTree.SelectedItem.Text = tmpName
        Else
            tmpName = stringFunct.ProperCase(CStr(QCTree.SelectedItem.Text))
        End If
        If optYes.Value = True Then
            tmpRun = "YES"
        Else
            tmpRun = "NO"
        End If
        tmpTime = Format(lstHour.ListIndex + 1, "00") & ":" & Format(lstMinute.ListIndex, "00")
        tmpTarget = Trim(txtSavePath.Text)
        For i = 0 To 6
            If chkDay(i).Value = Checked Then
                tmpDays = tmpDays & "1"
            Else
                tmpDays = tmpDays & "0"
            End If
        Next
        tmpSQL = Replace(Trim(stringFunct.RemoveAllEnter(txtSQL.Text)), ":", ";")
        tmpPostP = Trim(stringFunct.RemoveAllEnter(txtPostProcess.Text))
        FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", CStr(tmpKey), "�NAME=" & tmpName & "��SQL=" & tmpSQL & "��POSP=" & tmpPostP & "��RUN=" & tmpRun & "��TIME=" & tmpTime & "��TARGET=" & tmpTarget & "��DAYS=" & tmpDays & "��PARENTID=" & tmpFatherID & "�"
        ReportUpdated
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "cmdSave"
    SaveReport
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = "Report saved"
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
End Select
End Sub

Private Sub ClearForm()
Dim i
ClearTable
lstHour.Clear
lstMinute.Clear
For i = 1 To 24
    lstHour.AddItem Format(i, "00")
Next
lstHour.ListIndex = 0
For i = 0 To 59
    lstMinute.AddItem Format(i, "00")
Next
lstMinute.ListIndex = 0
txtSavePath.Text = App.path & "\SQC Reports" & "\"

For i = 0 To 6
    chkDay(i).Value = Unchecked
Next

txtSQL.Text = ""
txtLogs.Text = ""
txtPostProcess.Text = ""
optYes.Value = True
flxImport.Clear
QCTree.Nodes.Clear
QCTree.Nodes.Add , , "F", "SuperQC Reports", 1
mdiMain.pBar.Value = 0
mdiMain.pBar.Max = 100

LoadFolders
LoadReports

QCTree.Nodes("F").Expanded = True
ReportUpdated
TabSQL.Enabled = False
 Me.Caption = Me.Tag
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.Rows = 2
End Sub

Private Sub OutputTable(ColLetter As String)
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim i, Protections
Dim curTab
Dim w

FileWrite App.path & "\SQC DAT" & "\" & "REPORT01" & ".xls", AllScripts

Set xlObject = New Excel.Application

On Error Resume Next
For Each w In xlObject.Workbooks
   w.Close savechanges:=False
Next w
On Error GoTo 0

Set xlWB = xlObject.Workbooks.Open(App.path & "\SQC DAT" & "\" & "REPORT01" & ".xls")
    xlObject.Sheets.Add
    xlObject.Sheets(1).Range("A1").Value = "Report Name: " & txtReportName.Text
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
  xlObject.Workbooks(1).SaveAs FileFunct.AddBackslash(txtSavePath.Text) & "" & QCTree.SelectedItem.Text & "-" & Format(Now, "mm-dd-yyyy HH-MM AMPM")
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate

  Set xlWB = Nothing
  Set xlObject = Nothing
  FXGirl.EZPlay FXExportToExcel
  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export to MS Excel completed.": Exit Sub:
OutErr:     MsgBox Err.Description, vbCritical: xlObject.Visible = True: xlObject.ActiveWindow.Activate: Set xlWB = Nothing: Set xlObject = Nothing
On Error GoTo 0
End Sub

Private Sub txtPostProcess_Change()
ReportChanged
End Sub

Private Sub txtReportName_Change()
ReportChanged
End Sub

Private Sub txtSavePath_Change()
ReportChanged
End Sub

Private Sub txtSQL_Change()
ReportChanged
End Sub

Private Sub BubbleSort_QC_Tree(ByRef pvarArray1 As Variant, ByRef pvarArray2 As Variant, ByRef pvarArray3 As Variant)
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    Dim varSwap2 As Variant
    Dim varSwap3 As Variant
    Dim blnSwapped As Boolean
    
    iMin = LBound(pvarArray1)
    iMax = UBound(pvarArray1) - 1
    Do
        blnSwapped = False
        For i = iMin To iMax
            If pvarArray1(i) > pvarArray1(i + 1) Then
                varSwap = pvarArray1(i)
                pvarArray1(i) = pvarArray1(i + 1)
                pvarArray1(i + 1) = varSwap
                varSwap2 = pvarArray2(i)
                pvarArray2(i) = pvarArray2(i + 1)
                pvarArray2(i + 1) = varSwap2
                varSwap3 = pvarArray3(i)
                pvarArray3(i) = pvarArray3(i + 1)
                pvarArray3(i + 1) = varSwap3
                blnSwapped = True
            End If
        Next
        iMax = iMax - 1
    Loop Until Not blnSwapped
End Sub

Private Sub Form_Load()
ClearForm
End Sub

Private Sub PerformPostProcessing(strScript As String)
Dim tmp, LCol, RCol, i
strScript = UCase(Trim(strScript))
If stringFunct.StrIn(strScript, UCase("GetRequirementFolderPath")) = True Then
    tmp = Split(strScript, ",")
    tmp(0) = Replace(tmp(0), UCase("GetRequirementFolderPath"), "")
    tmp(0) = Replace(tmp(0), "(", "")
    tmp(0) = Replace(tmp(0), " ", "")
    tmp(1) = Replace(tmp(1), ")", "")
    tmp(1) = Replace(tmp(1), " ", "")
    LCol = CInt(tmp(0)) - 1
    RCol = CInt(tmp(1)) - 1
    For i = 1 To flxImport.Rows - 1
        flxImport.TextMatrix(i, LCol) = GetRequirementFolderPath(flxImport.TextMatrix(i, RCol))
    Next
ElseIf stringFunct.StrIn(strScript, UCase("GetTestSetFolderPath")) = True Then
    tmp = Split(strScript, ",")
    tmp(0) = Replace(tmp(0), UCase("GetTestSetFolderPath"), "")
    tmp(0) = Replace(tmp(0), "(", "")
    tmp(0) = Replace(tmp(0), " ", "")
    tmp(1) = Replace(tmp(1), ")", "")
    tmp(1) = Replace(tmp(1), " ", "")
    LCol = CInt(tmp(0)) - 1
    RCol = CInt(tmp(1)) - 1
    For i = 1 To flxImport.Rows - 1
        flxImport.TextMatrix(i, LCol) = GetTestSetFolderPath(flxImport.TextMatrix(i, RCol))
    Next
ElseIf stringFunct.StrIn(strScript, UCase("GetBusinessComponentFolderPath")) = True Then
    tmp = Split(strScript, ",")
    tmp(0) = Replace(tmp(0), UCase("GetBusinessComponentFolderPath"), "")
    tmp(0) = Replace(tmp(0), "(", "")
    tmp(0) = Replace(tmp(0), " ", "")
    tmp(1) = Replace(tmp(1), ")", "")
    tmp(1) = Replace(tmp(1), " ", "")
    LCol = CInt(tmp(0)) - 1
    RCol = CInt(tmp(1)) - 1
    For i = 1 To flxImport.Rows - 1
        flxImport.TextMatrix(i, LCol) = GetBusinessComponentFolderPath(flxImport.TextMatrix(i, RCol))
    Next
ElseIf stringFunct.StrIn(strScript, UCase("GetTestFolderPath")) = True Then
    tmp = Split(strScript, ",")
    tmp(0) = Replace(tmp(0), UCase("GetTestFolderPath"), "")
    tmp(0) = Replace(tmp(0), "(", "")
    tmp(0) = Replace(tmp(0), " ", "")
    tmp(1) = Replace(tmp(1), ")", "")
    tmp(1) = Replace(tmp(1), " ", "")
    LCol = CInt(tmp(0)) - 1
    RCol = CInt(tmp(1)) - 1
    For i = 1 To flxImport.Rows - 1
        flxImport.TextMatrix(i, LCol) = GetTestFolderPath(flxImport.TextMatrix(i, RCol))
    Next
End If
End Sub
