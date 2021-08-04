VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFunctionControl 
   Caption         =   "Procedures Management Module"
   ClientHeight    =   11355
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   16590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11355
   ScaleWidth      =   16590
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFilter 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   29
      Top             =   600
      Width           =   2355
   End
   Begin VB.TextBox txtLibraryName 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   14115
   End
   Begin VB.ComboBox cmbLibraryFileName 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4740
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   11775
   End
   Begin TabDlg.SSTab tabSource 
      Height          =   9435
      Left            =   120
      TabIndex        =   0
      Top             =   1500
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   16642
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   3528
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Details"
      TabPicture(0)   =   "frmFunctionControl.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtLibraryDesc"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtOwner"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCreationDate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Actual Source"
      TabPicture(1)   =   "frmFunctionControl.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmbProcName"
      Tab(1).Control(1)=   "cmdGo_1"
      Tab(1).Control(2)=   "txtActual"
      Tab(1).Control(3)=   "Label1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Last Versions"
      TabPicture(2)   =   "frmFunctionControl.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtModifiedBy"
      Tab(2).Control(1)=   "txtVersionDesc"
      Tab(2).Control(2)=   "cmbVerProcName"
      Tab(2).Control(3)=   "cmbVersion"
      Tab(2).Control(4)=   "cmdGo2_1"
      Tab(2).Control(5)=   "cmdGo2_2"
      Tab(2).Control(6)=   "txtVersionCode"
      Tab(2).Control(7)=   "Label10"
      Tab(2).Control(8)=   "Label6"
      Tab(2).Control(9)=   "Label4"
      Tab(2).Control(10)=   "Label5"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "History Logs"
      TabPicture(3)   =   "frmFunctionControl.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtHistoryLog"
      Tab(3).ControlCount=   1
      Begin VB.TextBox txtModifiedBy 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73020
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1980
         Width           =   13815
      End
      Begin VB.TextBox txtCreationDate 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   780
         Width           =   14355
      End
      Begin VB.TextBox txtOwner 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   420
         Width           =   14355
      End
      Begin VB.TextBox txtVersionDesc 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -73020
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   780
         Width           =   13815
      End
      Begin VB.ComboBox cmbProcName 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73020
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   420
         Width           =   13815
      End
      Begin VB.CommandButton cmdGo_1 
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -59160
         TabIndex        =   15
         Top             =   420
         Width           =   435
      End
      Begin VB.ComboBox cmbVerProcName 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73020
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2400
         Width           =   13815
      End
      Begin VB.ComboBox cmbVersion 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73020
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   420
         Width           =   13815
      End
      Begin VB.CommandButton cmdGo2_1 
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -59160
         TabIndex        =   9
         Top             =   420
         Width           =   435
      End
      Begin VB.CommandButton cmdGo2_2 
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -59160
         TabIndex        =   8
         Top             =   2400
         Width           =   435
      End
      Begin RichTextLib.RichTextBox txtHistoryLog 
         Height          =   8775
         Left            =   -74880
         TabIndex        =   7
         Top             =   420
         Width           =   16215
         _ExtentX        =   28601
         _ExtentY        =   15478
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmFunctionControl.frx":0070
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtVersionCode 
         Height          =   6435
         Left            =   -74940
         TabIndex        =   12
         Top             =   2820
         Width           =   16215
         _ExtentX        =   28601
         _ExtentY        =   11351
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmFunctionControl.frx":00F1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtActual 
         Height          =   8415
         Left            =   -74940
         TabIndex        =   17
         Top             =   840
         Width           =   16215
         _ExtentX        =   28601
         _ExtentY        =   14843
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmFunctionControl.frx":0172
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtLibraryDesc 
         Height          =   7815
         Left            =   120
         TabIndex        =   19
         Top             =   1500
         Width           =   16215
         _ExtentX        =   28601
         _ExtentY        =   13785
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmFunctionControl.frx":01F3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         Caption         =   "Modified By:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         TabIndex        =   28
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Creation Date:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Owner:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         TabIndex        =   20
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Procedure Name:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Procedure Name:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         TabIndex        =   14
         Top             =   2460
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Version:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16590
      _ExtentX        =   29263
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
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdAdd"
            Object.ToolTipText     =   "Add New Library"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdGenerate"
            Object.ToolTipText     =   "Generate"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "F_GET"
                  Text            =   "Get From Source File"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Picture         =   "frmFunctionControl.frx":0274
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctionControl.frx":0506
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctionControl.frx":0798
            Key             =   ""
         EndProperty
      EndProperty
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
            Picture         =   "frmFunctionControl.frx":0A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctionControl.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctionControl.frx":1576
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctionControl.frx":1C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctionControl.frx":239A
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
            Picture         =   "frmFunctionControl.frx":2AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctionControl.frx":2D8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctionControl.frx":32DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctionControl.frx":3830
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgOpenFunction 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Visual Basic Script | *.vbs*"
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   10980
      Width           =   16590
      _ExtentX        =   29263
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmFunctionControl.frx":3D81
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   28037
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Library Filename:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   660
      Width           =   2235
   End
   Begin VB.Label Label2 
      Caption         =   "Library Name:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1020
      Width           =   2235
   End
End
Attribute VB_Name = "frmFunctionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbSource As New clsDatabase
Dim Changed_ As Boolean

Private Sub cmdGo_1_Click()
If cmbProcName.ListIndex <> -1 Then
    txtActual.SelStart = InStr(1, txtActual.Text, cmbProcName.Text)
    txtActual.SelLength = Len(cmbProcName.Text)
End If
End Sub

Private Sub cmdGo2_1_Click()
Dim rs As New ADODB.Recordset
Dim stringFunct As New clsStrings
Dim AllLines() As String, i
Set rs = dbSource.GetRecordSetSQL("SELECT * FROM tblVersions WHERE txtVersionName = '" & Trim((cmbVersion.Text)) & "'")
If rs.RecordCount = 1 Then
    txtVersionCode.Text = rs.Fields("memLibraryCode")
    txtVersionDesc.Text = rs.Fields("memVersionDescription")
    txtModifiedBy.Text = rs.Fields("txtModifiedBy") & " on " & rs.Fields("dteModifyDate")
    stringFunct.VBTextColored txtVersionCode
    AllLines = Split(txtVersionCode.Text, vbCrLf)
    cmbVerProcName.Clear
    For i = LBound(AllLines) To UBound(AllLines)
        If InStr(1, AllLines(i), "Function") <> 0 And InStr(1, AllLines(i), "(") <> 0 And InStr(1, AllLines(i), ")") <> 0 Then
            cmbVerProcName.AddItem Trim(AllLines(i))
        ElseIf InStr(1, AllLines(i), "Sub") <> 0 And InStr(1, AllLines(i), "(") <> 0 And InStr(1, AllLines(i), ")") <> 0 Then
            cmbVerProcName.AddItem Trim(AllLines(i))
        End If
    Next
ElseIf rs.RecordCount > 1 Then
    MsgBox "Multiple data entry"
End If
End Sub

Private Sub cmdGo2_2_Click()
If cmbVerProcName.ListIndex <> -1 Then
    txtVersionCode.SelStart = InStr(1, txtVersionCode.Text, cmbVerProcName.Text)
    txtVersionCode.SelLength = Len(cmbVerProcName.Text)
End If
End Sub

Private Sub Form_Load()
Dim DbName As String
Dim FileFunct As New clsFiles
DbName = FileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "<DBNAME>")
If Trim(DbName) = "" Then
    DbName = InputBox("Enter database path", "No database connection", "X:\BAK\Function Tracker DB\Master Database.mdb")
    If Trim(DbName) = "" Then Unload Me: Exit Sub
    FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "<DBNAME>", DbName
End If
If FileFunct.FileExists(DbName) = False Then
    DbName = InputBox("Enter database path", "No database connection", "X:\BAK\Function Tracker DB\Master Database.mdb")
    If Trim(DbName) = "" Then Unload Me: Exit Sub
    FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "<DBNAME>", DbName
End If
    ClearForm
    dbSource.ConnectToMDB DbName, , "Welcome$1"
    Refresh_
End Sub

Private Sub Form_Resize()
On Error Resume Next
tabSource.height = stsBar.Top - 1600
tabSource.width = Me.width - tabSource.Left - 350
With Me
    .txtLibraryDesc.height = stsBar.Top - 3200
    .txtActual.height = stsBar.Top - 2500
    .txtVersionCode.height = stsBar.Top - 4500
    .txtHistoryLog.height = stsBar.Top - 2150
End With
End Sub

Private Sub Form_Terminate()
    dbSource.CloseDBADO
End Sub

Private Sub Refresh_()
Dim rs
Set rs = dbSource.GetRecordSetSQL("SELECT txtFileName FROM tblLibraries ORDER BY txtFileName ASC")
ClearForm
cmbLibraryFileName.Clear
cmbLibraryFileName.Locked = False
If rs.RecordCount <> 0 Then
    Do While rs.EOF <> True
        cmbLibraryFileName.AddItem rs.Fields("txtFileName").Value
        rs.MoveNext
    Loop
End If
End Sub

Private Sub Refresh_Filter(Filter As String)
Dim rs
Set rs = dbSource.GetRecordSetSQL("SELECT txtFileName FROM tblLibraries WHERE txtFileName LIKE '%" & Filter & "%' ORDER BY txtFileName ASC")
cmbLibraryFileName.Clear
If rs.RecordCount <> 0 Then
    Do While rs.EOF <> True
        cmbLibraryFileName.AddItem rs.Fields("txtFileName").Value
        rs.MoveNext
    Loop
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim LibFileName As String
Dim LibName As String
Dim SourceCode As String
Dim FileFunct As New clsFiles
Dim stringFunct As New clsStrings
Dim rs As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim OldCode As String
Dim i As Long, tmpLine As String, AllLines() As String

Select Case Button.Key
    Case "cmdRefresh"
        If Changed_ = False Then
            Refresh_
        Else
            If MsgBox("There are some changes made in the code. Are you sure you want to abandon these changes?", vbYesNo) = vbYes Then
                Refresh_
                Changed_ = False
            End If
        End If
    Case "cmdAdd"
        If Changed_ = False Then
cmdAddNow:
            dlgOpenFunction.filename = ""
            dlgOpenFunction.ShowOpen
            LibFileName = dlgOpenFunction.filename
            LibName = dlgOpenFunction.FileTitle
            If Trim(LibFileName) <> "" Then
                SourceCode = FileFunct.ReadFromFile(LibFileName)
                Set rs = dbSource.GetRecordSetSQL("SELECT * FROM tblLibraries WHERE txtFileName = '" & Trim((LibFileName)) & "'")
                If rs.RecordCount = 0 Then
                    rsAdd.Open "SELECT * FROM tblLibraries", dbSource.cnn, adOpenDynamic, adLockOptimistic
                    rsAdd.AddNew
                    frmLibraryDialog.txtLibraryFileName = LibFileName
                    frmLibraryDialog.txtLibraryName = LibName
                    frmLibraryDialog.txtOwner = curUser
                    frmLibraryDialog.txtDescription = ""
                    frmLibraryDialog.Show 1
                    rsAdd.Fields("txtLibraryName") = Trim(LibName)
                    rsAdd.Fields("memLibraryDescription") = frmLibraryDialog.Description
                    txtLibraryDesc.Text = rsAdd.Fields("memLibraryDescription")
                    rsAdd.Fields("txtOwner") = frmLibraryDialog.Owner
                    txtOwner.Text = rsAdd.Fields("txtOwner")
                    rsAdd.Fields("dteCreationDate") = Now
                    txtCreationDate.Text = rsAdd.Fields("dteCreationDate")
                    rsAdd.Fields("memLibraryCode") = SourceCode
                    txtActual.Text = SourceCode
                    rsAdd.Fields("txtFileName") = LibFileName
                    rsAdd.Update
                    rsAdd.Close
                    stringFunct.VBTextColored txtActual
                    Changed_ = False
                    MsgBox LibFileName & " is now added in the system"
                    txtLibraryDesc.Enabled = True
                    txtActual.Enabled = True
                    cmbProcName.Enabled = True
                    cmbVerProcName.Enabled = True
                Else
                    MsgBox LibFileName & " is already existing in the system."
                End If
            End If
        Else
            If MsgBox("There are some changes made in the code. Are you sure you want to abandon these changes?", vbYesNo) = vbYes Then
                GoTo cmdAddNow
            End If
        End If
    Case "cmdGenerate"
        If Changed_ = False Then
cmdGenerateNow:
            cmbLibraryFileName.Locked = True
            If cmbLibraryFileName.ListIndex <> -1 Then
                GenClearForm
                Set rs = dbSource.GetRecordSetSQL("SELECT * FROM tblLibraries WHERE txtFileName = '" & Trim((cmbLibraryFileName.Text)) & "'")
                If rs.RecordCount = 1 Then
                    With Me
                        .txtLibraryName = rs.Fields("txtLibraryName")
                        .txtActual.Text = rs.Fields("memLibraryCode")
                        .txtLibraryName.Tag = rs.Fields("lngLibraryID")
                        .txtLibraryDesc.Text = rs.Fields("memLibraryDescription")
                        .txtOwner.Text = rs.Fields("txtOwner")
                        .txtCreationDate.Text = rs.Fields("dteCreationDate")
                        stringFunct.VBTextColored .txtActual
                        Changed_ = False
                        .txtLibraryDesc.Enabled = True
                        .txtActual.Enabled = True
                        .cmbProcName.Enabled = True
                        .cmbVerProcName.Enabled = True
                    End With
                ElseIf rs.RecordCount > 1 Then
                    MsgBox "Duplicate entries found"
                End If
                Set rs = dbSource.GetRecordSetSQL("SELECT TOP 25 txtVersionName FROM tblVersions WHERE lngLibraryID = " & Trim((txtLibraryName.Tag)) & " ORDER BY dteModifyDate DESC")
                With Me
                    .cmbVersion.Clear
                    If rs.RecordCount > 0 Then
                        Do While rs.EOF <> True
                            .cmbVersion.AddItem rs.Fields("txtVersionName")
                            rs.MoveNext
                        Loop
                    End If
                    Set rs = dbSource.GetRecordSetSQL("SELECT lngVersionID, txtVersionName, txtModifiedBy, dteModifyDate FROM tblVersions WHERE lngLibraryID = " & Trim((txtLibraryName.Tag)) & " ORDER BY dteModifyDate DESC")
                    .txtHistoryLog.Text = ""
                    .txtHistoryLog.Text = "Version ID" & vbTab & vbTab & "Modified By" & vbTab & vbTab & "Date" & vbTab & vbTab & vbTab & vbTab & "Version Name" & vbCrLf
                    If rs.RecordCount > 0 Then
                        Do While rs.EOF <> True
                            .txtHistoryLog.Text = .txtHistoryLog.Text & Format(rs.Fields("lngVersionID"), "V0000000000") & vbTab & vbTab & rs.Fields("txtModifiedBy") & vbTab & vbTab & vbTab & rs.Fields("dteModifyDate") & vbTab & vbTab & rs.Fields("txtVersionName") & vbCrLf
                            rs.MoveNext
                        Loop
                    End If
                End With
                AllLines = Split(txtActual.Text, vbCrLf)
                cmbProcName.Clear
                For i = LBound(AllLines) To UBound(AllLines)
                    If InStr(1, AllLines(i), "Function") <> 0 And InStr(1, AllLines(i), "(") <> 0 And InStr(1, AllLines(i), ")") <> 0 Then
                        cmbProcName.AddItem Trim(AllLines(i))
                    ElseIf InStr(1, AllLines(i), "Sub") <> 0 And InStr(1, AllLines(i), "(") <> 0 And InStr(1, AllLines(i), ")") <> 0 Then
                        cmbProcName.AddItem Trim(AllLines(i))
                    End If
                Next
            End If
            Changed_ = False
        Else
            If MsgBox("There are some changes made in the code. Are you sure you want to abandon these changes?", vbYesNo) = vbYes Then
                GoTo cmdGenerateNow
            End If
        End If
    Case "cmdUpload"
        If cmbLibraryFileName.ListIndex <> -1 Then
            If MsgBox("Are you sure you want to upload your changes to the system?", vbYesNo) = vbYes Then
                Set rs = dbSource.GetRecordSetSQL("SELECT * FROM tblLibraries WHERE lngLibraryID = " & Trim((txtLibraryName.Tag)) & "")
                If rs.RecordCount = 1 Then
                    SourceCode = Trim(txtActual.Text)
                    LibName = txtLibraryName.Text
                    LibFileName = cmbLibraryFileName.Text
                    rsAdd.Open "SELECT * FROM tblLibraries WHERE lngLibraryID = " & Trim((txtLibraryName.Tag)), dbSource.cnn, adOpenDynamic, adLockOptimistic
                    OldCode = rs.Fields("memLibraryCode")
                    rsAdd.Fields("memLibraryDescription") = txtLibraryDesc.Text
                    rsAdd.Fields("memLibraryCode") = SourceCode
                    rsAdd.Update
                    rsAdd.Close
                    rsAdd.Open "SELECT * FROM tblVersions", dbSource.cnn, adOpenDynamic, adLockOptimistic
                    rsAdd.AddNew
                    rsAdd.Fields("lngLibraryID") = txtLibraryName.Tag
                    rsAdd.Fields("txtVersionName") = LibName & " (" & Now & ")"
                    frmLibraryDialog.txtLibraryFileName = LibFileName
                    frmLibraryDialog.txtLibraryName = rsAdd.Fields("txtVersionName")
                    frmLibraryDialog.lblOwner = "Modified By:"
                    frmLibraryDialog.txtOwner.Locked = True
                    frmLibraryDialog.txtOwner = curUser
                    frmLibraryDialog.txtDescription = ""
                    frmLibraryDialog.Show 1
                    rsAdd.Fields("memVersionDescription") = frmLibraryDialog.Description
                    rsAdd.Fields("txtModifiedBy") = curUser
                    rsAdd.Fields("dteModifyDate") = Now
                    rsAdd.Fields("memLibraryCode") = SourceCode
                    rsAdd.Update
                    rsAdd.Close
                    FileFunct.FileWrite LibFileName, SourceCode
                    Changed_ = False
                    MsgBox LibFileName & " is now updated in the system"
                Else
                    MsgBox "Duplicate entries found."
                End If
            End If
        End If
End Select
End Sub

Private Sub ClearForm()
With Me
    .cmbLibraryFileName.Clear
    .txtLibraryName = ""
    .cmbProcName.Clear
    .txtActual.Text = ""
    .cmbVerProcName.Clear
    .cmbVersion.Clear
    .txtVersionCode.Text = ""
    .txtHistoryLog.Text = ""
    .txtLibraryName.Tag = ""
    .txtOwner.Text = ""
    .txtCreationDate.Text = ""
    .txtLibraryDesc.Text = ""
    .txtVersionDesc.Text = ""
    .txtModifiedBy.Text = ""
    .tabSource.Tab = 0
    .txtLibraryDesc.Enabled = False
    .txtActual.Enabled = False
    .cmbProcName.Enabled = False
    .cmbVerProcName.Enabled = False
    .txtFilter.Text = ""
    Changed_ = False
End With
End Sub

Private Sub GenClearForm()
With Me
    .txtLibraryName = ""
    .cmbProcName.Clear
    .txtActual.Text = ""
    .cmbVerProcName.Clear
    .cmbVersion.Clear
    .txtVersionCode.Text = ""
    .txtHistoryLog.Text = ""
    .txtLibraryName.Tag = ""
    .txtOwner.Text = ""
    .txtCreationDate.Text = ""
    .txtLibraryDesc.Text = ""
    .txtVersionDesc.Text = ""
    .txtModifiedBy.Text = ""
    .tabSource.Tab = 0
    .txtLibraryDesc.Enabled = False
    .txtActual.Enabled = False
    .cmbProcName.Enabled = False
    .cmbVerProcName.Enabled = False
    .txtFilter.Text = ""
    Changed_ = False
End With
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim FileFunct As New clsFiles
Dim stringFunct As New clsStrings
If ButtonMenu.Key = "F_GET" Then
    If cmbLibraryFileName.ListIndex <> -1 And Trim(txtLibraryName.Text) <> "" Then
        If MsgBox("Are you sure you want to get the code from the this location and overwrite the latest code in the system?", vbYesNo) = vbYes Then
            txtActual.Text = FileFunct.ReadFromFile(cmbLibraryFileName.Text)
            stringFunct.VBTextColored txtActual
            Changed_ = True
        End If
    Else
        MsgBox "Select and generate a Function Library first"
    End If
End If
End Sub

Private Sub txtActual_Change()
    Changed_ = True
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Refresh_Filter txtFilter.Text
End If
End Sub

Private Sub txtLibraryDesc_Change()
    Changed_ = True
End Sub
