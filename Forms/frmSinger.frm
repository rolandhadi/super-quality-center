VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSinger 
   Caption         =   "Singer Module"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   Icon            =   "frmSinger.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10770
   ScaleWidth      =   12660
   Tag             =   "Business Component Module"
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkInvertDate 
      Caption         =   "Invert Date"
      Height          =   315
      Left            =   4740
      TabIndex        =   12
      Top             =   660
      Width           =   1155
   End
   Begin VB.CheckBox chkAutoDates 
      Caption         =   "Auto Compute Dates"
      Height          =   315
      Left            =   6000
      TabIndex        =   11
      Top             =   660
      Width           =   1875
   End
   Begin VB.CommandButton cmdImportQTP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2340
      Picture         =   "frmSinger.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Import step description and expected results from an excel file"
      Top             =   600
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      Height          =   435
      Left            =   60
      TabIndex        =   4
      Top             =   1140
      Width           =   5115
      Begin VB.CommandButton cmdRemove 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdUp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Move Up"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdDown 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Move Down"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
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
      Height          =   495
      Left            =   60
      Picture         =   "frmSinger.frx":18D3
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Import step description and expected results from an excel file"
      Top             =   600
      Width           =   2205
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12660
      _ExtentX        =   22331
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinger.frx":2B99
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinger.frx":2E2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinger.frx":861D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinger.frx":88AF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView QCTree 
      Height          =   8655
      Left            =   60
      TabIndex        =   0
      Top             =   1620
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   15266
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
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
      OLEDragMode     =   1
      OLEDropMode     =   1
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
            Picture         =   "frmSinger.frx":8B3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinger.frx":924F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinger.frx":9961
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinger.frx":A073
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
      TabIndex        =   3
      Top             =   10395
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmSinger.frx":A785
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   21105
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
            Picture         =   "frmSinger.frx":ACD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinger.frx":AFB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinger.frx":B509
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinger.frx":BA5A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgQTP 
      Left            =   12180
      Top             =   780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "QTP Script | *.usr*"
   End
   Begin MSComDlg.CommonDialog dlgOpenXML 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "TAO Logs | *.xml*"
   End
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   8655
      Left            =   6900
      TabIndex        =   10
      Top             =   1620
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   15266
      _Version        =   393216
      Cols            =   8
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "frmSinger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NodeType
   Key As String
   Tag As String
   Text As String
End Type

Dim indrag As Boolean ' Flag that signals a Drag Drop operation.
Dim nodx As Node ' Item that is being dragged. -- Changed to Node
Dim tmpPars() As TestPlan_Pars
Dim FromClick As Boolean

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


Private Sub cmdAdd_Click()
Dim cnt
cnt = Val(Replace(Left(QCTree.SelectedItem.Text, 6), "[", ""))
frmControlLogs.EXECUTION_TIME = flxImport.TextMatrix(cnt, 1)
frmControlLogs.ELAPSED_TIME = flxImport.TextMatrix(cnt, 2)
frmControlLogs.STEP_RESULT = flxImport.TextMatrix(cnt, 3)
frmControlLogs.STEP_SUMMARY = flxImport.TextMatrix(cnt, 4)
frmControlLogs.Component_Name = flxImport.TextMatrix(cnt, 5)
frmControlLogs.STEP_DESCRIPTION = flxImport.TextMatrix(cnt, 6)
frmControlLogs.Show
PushToFlex
End Sub


Private Sub cmdImportQTP_Click()
Dim fname As String
Dim lastrow
Dim i, j, FileFunct As New clsFiles, stringFunct As New clsStrings
Dim tmpSplit, X, tmpDate As String, c, CNV
'On Error GoTo ErrLoad
    dlgOpenXML.filename = "": dlgOpenXML.ShowOpen
    fname = dlgOpenXML.filename
    If Trim(fname) = "" Then Exit Sub
    If FileFunct.FileExists(fname) = False Then MsgBox "File does not exist": Exit Sub
    cmdUp.Enabled = True
    cmdDown.Enabled = True
    cmdAdd.Enabled = True
    cmdRemove.Enabled = True
    FileFunct.LoadXMLDocument_v2 fname
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
    QCTree.Nodes.Clear
    QCTree.Nodes.Add , , "Root", Replace(fname, ".xml", ""), 1
    QCTree.Nodes(1).Selected = True
    c = 1
    X = 1
    CNV = 1
    For i = LBound(tmpSplit) To UBound(tmpSplit) - 1
        flxImport.TextMatrix(i + 1, 0) = i + 1
        If InStr(1, tmpSplit(i), "TAOVERSION:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "TAOVERSION:", ""))
            flxImport.TextMatrix(X, 3) = "TAOVERSION"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TAOVERSION", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TAOVERSION", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "TYPE:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "TYPE:", ""))
            flxImport.TextMatrix(X, 3) = "TYPE"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TYPE", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TYPE", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "QCUSER:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "QCUSER:", ""))
            flxImport.TextMatrix(X, 3) = "QCUSER"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - QCUSER", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "QCUSER", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "QCDOMAIN:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "QCDOMAIN:", ""))
            flxImport.TextMatrix(X, 3) = "QCDOMAIN"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - QCDOMAIN", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "QCDOMAIN", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "QCPROJECT:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "QCPROJECT:", ""))
            flxImport.TextMatrix(X, 3) = "QCPROJECT"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - QCPROJECT", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "QCPROJECT", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "TESTSETNAME:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "TESTSETNAME:", ""))
            flxImport.TextMatrix(X, 3) = "TESTSETNAME"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TESTSETNAME", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TESTSETNAME", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "TESTSETID:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "TESTSETID:", ""))
            flxImport.TextMatrix(X, 3) = "TESTSETID"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TESTSETID", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TESTSETID", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "TESTNAME:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "TESTNAME:", ""))
            flxImport.TextMatrix(X, 3) = "TESTNAME"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TESTNAME", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TESTNAME", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "TESTID:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "TESTID:", ""))
            flxImport.TextMatrix(X, 3) = "TESTID"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TESTID", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TESTID", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "TITLE:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "TITLE:", ""))
            flxImport.TextMatrix(X, 3) = "TITLE"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TITLE", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TITLE", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "TABLE_HEADER_TEXT:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "TABLE_HEADER_TEXT:", ""))
            flxImport.TextMatrix(X, 3) = "TABLE_HEADER_TEXT"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TABLE_HEADER_TEXT", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TABLE_HEADER_TEXT", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "TABLE_HEADER_TIME:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "TABLE_HEADER_TIME:", ""))
            flxImport.TextMatrix(X, 3) = "TABLE_HEADER_TIME"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TABLE_HEADER_TIME", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TABLE_HEADER_TIME", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
         ElseIf InStr(1, tmpSplit(i), "DEBUG_LOG:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "DEBUG_LOG:", ""))
            flxImport.TextMatrix(X, 3) = "DEBUG_LOG"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - DEBUG_LOG", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "DEBUG_LOG", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "LOG_FOLDER:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "LOG_FOLDER:", ""))
            flxImport.TextMatrix(X, 3) = "LOG_FOLDER"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - LOG_FOLDER", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "LOG_FOLDER", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
         ElseIf InStr(1, tmpSplit(i), "EXECUTED_BY:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "EXECUTED_BY:", ""))
            flxImport.TextMatrix(X, 3) = "EXECUTED_BY"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - EXECUTED_BY", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "EXECUTED_BY", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "PC:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "PC:", ""))
            flxImport.TextMatrix(X, 3) = "PC"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - PC", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "PC", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "STATUS:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "STATUS:", ""))
            flxImport.TextMatrix(X, 3) = "STATUS"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - STATUS", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STATUS", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "EXECUTION_TIME:") <> 0 Then
            If chkInvertDate.Value = Checked Then
                tmpDate = stringFunct.Switch_Month_Day_DateFormat(CDate((Trim(Replace(tmpSplit(i), "EXECUTION_TIME:", "")))), "/", "dd/mmm/yyyy")
            Else
                tmpDate = Format(CDate((Trim(Replace(tmpSplit(i), "EXECUTION_TIME:", "")))), "dd/mmm/yyyy")
            End If
            flxImport.TextMatrix(X, 1) = tmpDate & " " & Format(CDate((Trim(Replace(tmpSplit(i), "EXECUTION_TIME:", "")))), "hh:mm:ss")
            On Error Resume Next
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] Log Entry", 4
            If Err.Number <> 0 Then
                On Error GoTo 0
                c = c + 1
                QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] Log Entry", 4
            End If
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "EXECUTION_TIME", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 1), 3
            CNV = CNV + 1
        ElseIf InStr(1, tmpSplit(i), "ELAPSED_TIME:") <> 0 Then
            flxImport.TextMatrix(X, 2) = Trim(Replace(tmpSplit(i), "ELAPSED_TIME:", ""))
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "ELAPSED_TIME", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 2), 3
            CNV = CNV + 1
        ElseIf InStr(1, tmpSplit(i), "STEP_RESULT:") <> 0 Then
            flxImport.TextMatrix(X, 3) = Trim(Replace(tmpSplit(i), "STEP_RESULT:", ""))
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STEP_RESULT", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 3), 3
            CNV = CNV + 1
        ElseIf InStr(1, tmpSplit(i), "STEP_SUMMARY:") <> 0 Then
            flxImport.TextMatrix(X, 4) = Trim(Replace(tmpSplit(i), "STEP_SUMMARY:", ""))
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STEP_SUMMARY", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            CNV = CNV + 1
        ElseIf InStr(1, tmpSplit(i), "COMPONENT_NAME:") <> 0 Then
            flxImport.TextMatrix(X, 5) = Trim(Replace(tmpSplit(i), "COMPONENT_NAME:", ""))
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "COMPONENT_NAME", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 5), 3
            CNV = CNV + 1
        ElseIf InStr(1, tmpSplit(i), "STEP_DESCRIPTION:") <> 0 Then
            flxImport.TextMatrix(X, 6) = Trim(Replace(tmpSplit(i), "STEP_DESCRIPTION:", ""))
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STEP_DESCRIPTION", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 6), 3
            CNV = CNV + 1
            c = c + 1
            X = X + 1
        ElseIf InStr(1, tmpSplit(i), "IMAGE_PATH:") <> 0 Then
            flxImport.TextMatrix(X - 1, 7) = Trim(flxImport.TextMatrix(X, 2) & " Captured Image Stored in: " & Trim(Replace(tmpSplit(i), "IMAGE_PATH:", "")))
            QCTree.Nodes.Add CStr("C" & c - 1), tvwChild, CStr("N" & CNV), "IMAGE_PATH", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X - 1, 7), 3
            CNV = CNV + 1
        End If
        If chkAutoDates.Value = Checked Then
            If X <> "1" And flxImport.TextMatrix(X - 1, 1) <> "" And flxImport.TextMatrix(X - 1, 1) <> "-" Then
                flxImport.TextMatrix(X, 1) = Format(DateAdd("s", Val(flxImport.TextMatrix(X, 2)), flxImport.TextMatrix(X - 1, 1)), "dd/mmm/yyyy hh:mm:ss")
            End If
        End If
    Next
    flxImport.Rows = X + 1
    flxImport.TextMatrix(X, 1) = Format(DateAdd("s", 10, flxImport.TextMatrix(X - 1, 1)), "dd/mmm/yyyy hh:mm:ss")
    flxImport.TextMatrix(X, 4) = "End Business Component"
    On Error GoTo 0
    On Error Resume Next
    flxImport.TextMatrix(X, 6) = "End Business Component " & flxImport.TextMatrix(10, 4) 'Uncomment
    flxImport.TextMatrix(X, 2) = "10"
    flxImport.TextMatrix(X, 3) = "DONE"
    QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] FOOTER - END BUSINESS COMPONENT", 4
    QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "EXECUTION_TIME", 2
    QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 1), 3
    CNV = CNV + 1
    QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "ELAPSED_TIME", 2
    QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 2), 3
    CNV = CNV + 1
    QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STEP_RESULT", 2
    QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 3), 3
    CNV = CNV + 1
    QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STEP_SUMMARY", 2
    QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
    CNV = CNV + 1
    QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "COMPONENT_NAME", 2
    QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 5), 3
    CNV = CNV + 1
    QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STEP_DESCRIPTION", 2
    QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 6), 3
    CNV = CNV + 1
Exit Sub
ErrLoad:
MsgBox "There was an error while importing the file. Please refresh and close all excel and try again" & vbCrLf & Err.Description, vbCritical
End Sub

Private Function GetNodeKey(X As String)
Dim tmpNode
Set tmpNode = QCTree.Nodes("Root").Child.FirstSibling
 Do While True
        If tmpNode Is Nothing Then Exit Do
        If Left(tmpNode.Key, 1) = "C" Then
            If InStr(1, tmpNode.Text, X) <> 0 Then
                GetNodeKey = tmpNode.Key
                Exit Function
            End If
        End If
        Set tmpNode = tmpNode.Next
Loop
End Function

Public Sub PushToFlex()
Dim w, cnt, curBC, curOrder, curID, tmpNode, tmpFName, curRow, curCol
cnt = 1
    Set tmpNode = QCTree.Nodes("Root")
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
    Do While True
        If tmpNode Is Nothing Then Exit Do
        If Left(tmpNode.Key, 1) = "C" Then
            curRow = CInt(Replace(Left(tmpNode.Text, 5), "[", ""))
            flxImport.TextMatrix(curRow, 0) = curRow
            If Not (tmpNode.Child.FirstSibling Is Nothing) Then Set tmpNode = tmpNode.Child.FirstSibling
        ElseIf Left(tmpNode.Key, 1) = "N" Then
            If InStr(1, tmpNode.Text, "TAOVERSION") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "TAOVERSION"
            ElseIf InStr(1, tmpNode.Text, "TYPE") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "TYPE"
            ElseIf InStr(1, tmpNode.Text, "QCUSER") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "QCUSER"
            ElseIf InStr(1, tmpNode.Text, "QCDOMAIN") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "QCDOMAIN"
            ElseIf InStr(1, tmpNode.Text, "QCPROJECT") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "QCPROJECT"
            ElseIf InStr(1, tmpNode.Text, "TESTSETNAME") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "TESTSETNAME"
            ElseIf InStr(1, tmpNode.Text, "TESTSETID") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "TESTSETID"
            ElseIf InStr(1, tmpNode.Text, "TESTNAME") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "TESTNAME"
            ElseIf InStr(1, tmpNode.Text, "TESTID") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "TESTID"
            ElseIf InStr(1, tmpNode.Text, "TITLE") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "TITLE"
            ElseIf InStr(1, tmpNode.Text, "TABLE_HEADER_TEXT") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "TABLE_HEADER_TEXT"
            ElseIf InStr(1, tmpNode.Text, "TABLE_HEADER_TIME") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "TABLE_HEADER_TIME"
            ElseIf InStr(1, tmpNode.Text, "DEBUG_LOG") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "DEBUG_LOG"
            ElseIf InStr(1, tmpNode.Text, "LOG_FOLDER") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "LOG_FOLDER"
            ElseIf InStr(1, tmpNode.Text, "EXECUTED_BY") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "EXECUTED_BY"
            ElseIf InStr(1, tmpNode.Text, "PC") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "PC"
            ElseIf InStr(1, tmpNode.Text, "STATUS") Then
                curCol = 4
                flxImport.TextMatrix(curRow, 3) = "STATUS"
            ElseIf InStr(1, tmpNode.Text, "EXECUTION_TIME") Then '1
                curCol = 1
            ElseIf InStr(1, tmpNode.Text, "ELAPSED_TIME") Then '1
                curCol = 2
            ElseIf InStr(1, tmpNode.Text, "STEP_RESULT") Then '1
                curCol = 3
            ElseIf InStr(1, tmpNode.Text, "STEP_SUMMARY") Then '1
                curCol = 4
            ElseIf InStr(1, tmpNode.Text, "COMPONENT_NAME") Then '1
                curCol = 5
            ElseIf InStr(1, tmpNode.Text, "STEP_DESCRIPTION") Then '1
                curCol = 6
            ElseIf InStr(1, tmpNode.Text, "IMAGE_PATH") Then '1
                curCol = 7
            End If
            If Not (tmpNode.Child.FirstSibling Is Nothing) Then Set tmpNode = tmpNode.Child.FirstSibling
        ElseIf Left(tmpNode.Key, 1) = "V" Then
            flxImport.TextMatrix(curRow, curCol) = tmpNode.Text
            If Not (tmpNode.Parent.Next Is Nothing) Then
                Set tmpNode = tmpNode.Parent.Next
            Else
                Set tmpNode = tmpNode.Parent.Parent.Next
            End If
        ElseIf tmpNode.FirstSibling.Index = 1 Then
            Set tmpNode = tmpNode.Child.FirstSibling
        End If
    Loop
End Sub

Private Sub cmdLoadExcel_Click()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim fname As String
Dim lastrow
Dim i, j, tmpParam, k, X, Node_BC, Node_Param, Y
Dim tmpSts, tmpDate
Dim strFunct As New clsFiles
Dim stringFunct As New clsStrings
Dim intFunct As New clsInternet
Dim oldNode, tmpName, c, CNV

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
    tmpName = Trim(xlObject.ActiveWorkbook.Sheets(2).Range("B6").Value)
    With xlObject.ActiveWorkbook.ActiveSheet
         If InStr(1, GetCommentText(.Range("A1")), "Singer Module") = 0 Then
            MsgBox "Import file is invalid. Please use only sheets generated by the SuperQC"
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
        lastrow = .Range("A" & .Rows.Count).End(xlUp).row
        mdiMain.pBar.Max = lastrow + 2
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
        flxImport.Rows = lastrow
        QCTree.Nodes.Clear
        QCTree.Nodes.Add , , "Root", Replace(fname, ".xml", ""), 1
        QCTree.Nodes(1).Selected = True
        c = 1
        X = 1
        CNV = 1
        Me.AutoRedraw = False
        flxImport.Visible = False
        For i = 1 To lastrow - 1
        mdiMain.pBar.Value = i
        flxImport.TextMatrix(i, 0) = i
        If Trim(.Range("D" & i + 1).Value) = "TAOVERSION" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("D" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "TAOVERSION"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TAOVERSION", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TAOVERSION", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "TYPE" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "TYPE"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TYPE", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TYPE", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "QCUSER" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "QCUSER"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - QCUSER", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "QCUSER", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "QCDOMAIN" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "QCDOMAIN"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - QCDOMAIN", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "QCDOMAIN", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "QCPROJECT" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "QCPROJECT"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - QCPROJECT", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "QCPROJECT", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "TESTSETNAME" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "TESTSETNAME"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TESTSETNAME", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TESTSETNAME", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "TESTSETID" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "TESTSETID"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TESTSETID", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TESTSETID", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "TESTNAME" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "TESTNAME"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TESTNAME", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TESTNAME", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "TESTID" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "TESTID"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TESTID", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TESTID", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "TITLE" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "TITLE"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TITLE", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TITLE", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "TABLE_HEADER_TEXT" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "TABLE_HEADER_TEXT"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TABLE_HEADER_TEXT", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TABLE_HEADER_TEXT", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "TABLE_HEADER_TIME" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "TABLE_HEADER_TIME"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - TABLE_HEADER_TIME", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "TABLE_HEADER_TIME", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
         ElseIf Trim(.Range("D" & i + 1).Value) = "DEBUG_LOG" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "DEBUG_LOG"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - DEBUG_LOG", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "DEBUG_LOG", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "LOG_FOLDER" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "LOG_FOLDER"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - LOG_FOLDER", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "LOG_FOLDER", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
         ElseIf Trim(.Range("D" & i + 1).Value) = "EXECUTED_BY" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "EXECUTED_BY"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - EXECUTED_BY", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "EXECUTED_BY", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "PC" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "PC"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - PC", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "PC", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        ElseIf Trim(.Range("D" & i + 1).Value) = "STATUS" Then
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            flxImport.TextMatrix(X, 3) = "STATUS"
            flxImport.TextMatrix(X, 1) = "-"
            flxImport.TextMatrix(X, 2) = "-"
            flxImport.TextMatrix(X, 5) = "-"
            flxImport.TextMatrix(X, 6) = "-"
            flxImport.TextMatrix(X, 7) = "-"
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] HEADER - STATUS", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STATUS", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            c = c + 1
            CNV = CNV + 1
            X = X + 1
        Else
            If chkInvertDate.Value = Checked Then
                tmpDate = stringFunct.Switch_Month_Day_DateFormat(CDate((Trim(.Range("B" & i + 1).Value))), "/", "dd/mmm/yyyy")
            Else
                tmpDate = Format(CDate((Trim(.Range("B" & i + 1).Value))), "dd/mmm/yyyy")
            End If
            flxImport.TextMatrix(X, 1) = tmpDate & " " & Format(CDate((Trim(.Range("B" & i + 1).Value))), "hh:mm:ss")
            QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] Log Entry", 4
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "EXECUTION_TIME", 2
            If chkAutoDates.Value = Checked Then
                If X <> "1" And flxImport.TextMatrix(X - 1, 1) <> "" And flxImport.TextMatrix(X - 1, 1) <> "-" Then
                    flxImport.TextMatrix(X, 1) = Format(DateAdd("s", Val(flxImport.TextMatrix(X, 2)), flxImport.TextMatrix(X - 1, 1)), "dd/mmm/yyyy hh:mm:ss")
                End If
            End If
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 1), 3
            CNV = CNV + 1
            flxImport.TextMatrix(X, 2) = Trim(.Range("C" & i + 1).Value)
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "ELAPSED_TIME", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 2), 3
            CNV = CNV + 1
            flxImport.TextMatrix(X, 3) = Trim(.Range("D" & i + 1).Value)
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STEP_RESULT", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 3), 3
            CNV = CNV + 1
            flxImport.TextMatrix(X, 4) = Trim(.Range("E" & i + 1).Value)
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STEP_SUMMARY", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 4), 3
            CNV = CNV + 1
            flxImport.TextMatrix(X, 5) = Trim(.Range("F" & i + 1).Value)
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "COMPONENT_NAME", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 5), 3
            CNV = CNV + 1
            flxImport.TextMatrix(X, 6) = Trim(.Range("G" & i + 1).Value)
            QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STEP_DESCRIPTION", 2
            QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 6), 3
            CNV = CNV + 1
            If Trim(.Range("H" & i + 1).Value) <> "" Then
                flxImport.TextMatrix(X, 7) = Trim(.Range("H" & i + 1).Value)
                QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "IMAGE_PATH", 2
                QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), flxImport.TextMatrix(X, 7), 3
                CNV = CNV + 1
            End If
            c = c + 1
            X = X + 1
        End If
    Next
    End With
    mdiMain.pBar.Value = mdiMain.pBar.Max
    Me.AutoRedraw = True
    flxImport.Visible = True
    QCTree.Nodes(1).Expanded = True
    QCTree.Nodes(1).Selected = True
    Me.QCTree.Visible = True
    mdiMain.pBar.Value = mdiMain.pBar.Max
    xlObject.DisplayAlerts = False 'To avoid "Save woorkbook" messagebox
    
    'Close Excel
    xlWB.Close
    xlObject.Application.Quit
    Set xlWB = Nothing
    Set xlObject = Nothing
    
        cmdUp.Enabled = True
        cmdDown.Enabled = True
        cmdAdd.Enabled = True
        cmdRemove.Enabled = True
Exit Sub
ErrLoad:
MsgBox "There was an error while importing the file. Please refresh and close all excel and try again" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub cmdRemove_Click()
Dim All_C()
Dim tmp, i, tmp_
Me.AutoRedraw = False
If Left(QCTree.SelectedItem.Key, 1) = "C" Then
    If MsgBox("Are you sure you want to delete this Component?", vbYesNo) = vbYes Then
        Control_Auto = False
        Me.QCTree.Visible = False
        ReDim All_C(0)
        QCTree.Nodes.Remove QCTree.SelectedItem.Index
        Set tmp = QCTree.Nodes("Root")
        Do While True
            If tmp Is Nothing Then Exit Do
            If Left(tmp.Key, 1) = "C" Then
                All_C(UBound(All_C)) = tmp.Key
                ReDim Preserve All_C(UBound(All_C) + 1)
            End If
            If tmp.FirstSibling.Index = 1 Then
                Set tmp = tmp.Child.FirstSibling
            Else
                Set tmp = tmp.Next
            End If
        Loop
        ReDim Preserve All_C(UBound(All_C) - 1)
        For i = LBound(All_C) To UBound(All_C)
                tmp = "[" & Format(i + 1, "0000") & "]"
                tmp_ = Trim(Right(QCTree.Nodes(All_C(i)).Text, Len(QCTree.Nodes(All_C(i)).Text) - 6))
                QCTree.Nodes(All_C(i)).Text = Replace(tmp, "]]", "]") & " " & Replace(tmp_, "] ", "")
        Next
        flxImport.Rows = flxImport.Rows - 1
    End If
End If
PushToFlex
Me.AutoRedraw = True
Me.QCTree.Visible = True
End Sub

Private Sub cmdUp_Click()
Dim All_C(), tmpNode
Dim NodeA, tmp, i, tmp_
Me.AutoRedraw = False
If Left(QCTree.SelectedItem.Key, 1) <> "C" Then Exit Sub
If (QCTree.SelectedItem.Previous Is Nothing) Then Exit Sub
ReDim All_C(0)
Me.QCTree.Visible = False
Set tmp = QCTree.Nodes("Root")
Do While True
    If tmp Is Nothing Then Exit Do
    If Left(tmp.Key, 1) = "C" Then
        All_C(UBound(All_C)) = tmp.Key
        ReDim Preserve All_C(UBound(All_C) + 1)
    End If
    If tmp.FirstSibling.Index = 1 Then
        Set tmp = tmp.Child.FirstSibling
    Else
        Set tmp = tmp.Next
    End If
Loop



ReDim Preserve All_C(UBound(All_C) - 1)
If Left(QCTree.SelectedItem.Key, 1) = "C" And Not (QCTree.SelectedItem.Previous Is Nothing) Then
    NodeA = GetArrInxFromInx(All_C, QCTree.SelectedItem.Key)

    tmp = Trim(Left(QCTree.Nodes(All_C(NodeA)).Text, 6))
    tmp_ = Trim(Right(QCTree.Nodes(All_C(NodeA - 1)).Text, Len(QCTree.Nodes(All_C(NodeA - 1)).Text) - 6))
    QCTree.Nodes(All_C(NodeA)).Text = Trim(Left(QCTree.Nodes(All_C(NodeA - 1)).Text, 6)) & " " & Trim(Right(QCTree.Nodes(All_C(NodeA)).Text, Len(QCTree.Nodes(All_C(NodeA)).Text) - 6))
    QCTree.Nodes(All_C(NodeA - 1)).Text = tmp & " " & tmp_
    
    tmp = All_C(NodeA)
    All_C(NodeA) = All_C(NodeA - 1)
    All_C(NodeA - 1) = tmp
End If
mdiMain.pBar.Max = UBound(All_C)
tmp = 0
For i = UBound(All_C) To LBound(All_C) Step -1
    Set QCTree.Nodes(All_C(i)).Parent = QCTree.Nodes("Root")
    mdiMain.pBar.Value = tmp
    tmp = tmp + 1
Next
mdiMain.pBar.Value = mdiMain.pBar.Max
PushToFlex
Me.AutoRedraw = True
Me.QCTree.Visible = True
FromClick = True
flxImport.col = 0
flxImport.row = Val(Replace(Left(QCTree.Nodes(QCTree.SelectedItem.Index).Text, 6), "[", ""))
flxImport.ColSel = 7
End Sub

Private Sub cmdDown_Click()
Dim All_C(), tmpNode
Dim NodeA, tmp, i, tmp_
Me.AutoRedraw = False
If Left(QCTree.SelectedItem.Key, 1) <> "C" Then Exit Sub
If (QCTree.SelectedItem.Next Is Nothing) Then Exit Sub
ReDim All_C(0)
Me.QCTree.Visible = False
Set tmp = QCTree.Nodes("Root")
Do While True
    If tmp Is Nothing Then Exit Do
    If Left(tmp.Key, 1) = "C" Then
        All_C(UBound(All_C)) = tmp.Key
        ReDim Preserve All_C(UBound(All_C) + 1)
    End If
    If tmp.FirstSibling.Index = 1 Then
        Set tmp = tmp.Child.FirstSibling
    Else
        Set tmp = tmp.Next
    End If
Loop



ReDim Preserve All_C(UBound(All_C) - 1)
If Left(QCTree.SelectedItem.Key, 1) = "C" And Not (QCTree.SelectedItem.Next Is Nothing) Then
    NodeA = GetArrInxFromInx(All_C, QCTree.SelectedItem.Key)

    tmp = Trim(Left(QCTree.Nodes(All_C(NodeA)).Text, 6))
    tmp_ = Trim(Right(QCTree.Nodes(All_C(NodeA + 1)).Text, Len(QCTree.Nodes(All_C(NodeA + 1)).Text) - 6))
    QCTree.Nodes(All_C(NodeA)).Text = Trim(Left(QCTree.Nodes(All_C(NodeA + 1)).Text, 6)) & " " & Trim(Right(QCTree.Nodes(All_C(NodeA)).Text, Len(QCTree.Nodes(All_C(NodeA)).Text) - 6))
    QCTree.Nodes(All_C(NodeA + 1)).Text = tmp & " " & tmp_
    
    tmp = All_C(NodeA)
    All_C(NodeA) = All_C(NodeA + 1)
    All_C(NodeA + 1) = tmp
End If
mdiMain.pBar.Max = UBound(All_C)
tmp = 0
For i = UBound(All_C) To LBound(All_C) Step -1
    Set QCTree.Nodes(All_C(i)).Parent = QCTree.Nodes("Root")
    mdiMain.pBar.Value = tmp
    tmp = tmp + 1
Next
mdiMain.pBar.Value = mdiMain.pBar.Max
PushToFlex
Me.AutoRedraw = True
Me.QCTree.Visible = True
FromClick = True
flxImport.col = 0
flxImport.row = Val(Replace(Left(QCTree.Nodes(QCTree.SelectedItem.Index).Text, 6), "[", ""))
flxImport.ColSel = 7
End Sub

Public Sub DragMove(Selected, Target)
Dim All_C(), tmpNode
Dim NodeA, tmp, i, tmp_, NodeB, tmpA

Me.AutoRedraw = False

If Left(Selected, 1) <> "C" Then Exit Sub
If Left(Target, 1) <> "C" Then Exit Sub
If MsgBox("Do you want to move this item?", vbYesNo) <> vbYes Then Exit Sub
ReDim All_C(0)
Me.QCTree.Visible = False
Set tmp = QCTree.Nodes("Root")
Do While True
    If tmp Is Nothing Then Exit Do
    If Left(tmp.Key, 1) = "C" Then
        All_C(UBound(All_C)) = tmp.Key
        ReDim Preserve All_C(UBound(All_C) + 1)
    End If
    If tmp.FirstSibling.Index = 1 Then
        Set tmp = tmp.Child.FirstSibling
    Else
        Set tmp = tmp.Next
    End If
Loop

ReDim Preserve All_C(UBound(All_C) - 1)
If Left(Selected, 1) = "C" And Left(Target, 1) = "C" Then
    NodeA = GetArrInxFromInx(All_C, Selected)
    NodeB = GetArrInxFromInx(All_C, Target)
    If NodeA > NodeB Then
        tmp = All_C(NodeA)
        All_C(NodeA) = ""
        For i = NodeA To NodeB + 1 Step -1
            All_C(i) = All_C(i - 1)
        Next
        All_C(NodeB) = tmp
    Else
        tmp = All_C(NodeA)
        All_C(NodeA) = ""
        For i = NodeA To NodeB - 1
            All_C(i) = All_C(i + 1)
        Next
        All_C(NodeB - 1) = tmp
    End If
    NodeA = GetArrInxFromInx(All_C, Selected)
    NodeB = GetArrInxFromInx(All_C, Target)
    For i = LBound(All_C) To UBound(All_C)
        tmp = "[" & Format(i + 1, "0000") & "]"
        tmp_ = Trim(Right(QCTree.Nodes(All_C(i)).Text, Len(QCTree.Nodes(All_C(i)).Text) - 6))
        QCTree.Nodes(All_C(i)).Text = Replace(tmp, "]]", "]") & " " & Replace(tmp_, "] ", "")
    Next
End If
mdiMain.pBar.Max = UBound(All_C)
tmp = 0
For i = UBound(All_C) To LBound(All_C) Step -1
    Set QCTree.Nodes(All_C(i)).Parent = QCTree.Nodes("Root")
    mdiMain.pBar.Value = tmp
    tmp = tmp + 1
Next
mdiMain.pBar.Value = mdiMain.pBar.Max
Me.AutoRedraw = True
Me.QCTree.Visible = True
End Sub

Public Sub DragMove2(Selected, Target)
Dim All_C(), tmpNode
Dim NodeA, tmp, i, tmp_, NodeB, tmpA

Me.AutoRedraw = False

If Left(Selected, 1) <> "C" Then Exit Sub
If Left(Target, 1) <> "C" Then Exit Sub
ReDim All_C(0)
Me.QCTree.Visible = False
Set tmp = QCTree.Nodes("Root")
Do While True
    If tmp Is Nothing Then Exit Do
    If Left(tmp.Key, 1) = "C" Then
        All_C(UBound(All_C)) = tmp.Key
        ReDim Preserve All_C(UBound(All_C) + 1)
    End If
    If tmp.FirstSibling.Index = 1 Then
        Set tmp = tmp.Child.FirstSibling
    Else
        Set tmp = tmp.Next
    End If
Loop

ReDim Preserve All_C(UBound(All_C) - 1)
If Left(Selected, 1) = "C" And Left(Target, 1) = "C" Then
    NodeA = GetArrInxFromInx(All_C, Selected) + 1
    NodeB = GetArrInxFromInx(All_C, Target)
    If NodeA <> NodeB Then
        tmp = All_C(NodeB)
        For i = NodeB To NodeA Step -1
            All_C(i) = All_C(i - 1)
        Next
        All_C(NodeA) = tmp
    End If
    For i = LBound(All_C) To UBound(All_C)
        tmp = "[" & Format(i + 1, "0000") & "]"
        tmp_ = Trim(Right(QCTree.Nodes(All_C(i)).Text, Len(QCTree.Nodes(All_C(i)).Text) - 6))
        QCTree.Nodes(All_C(i)).Text = Replace(tmp, "]]", "]") & " " & Replace(tmp_, "] ", "")
    Next
End If
mdiMain.pBar.Max = UBound(All_C)
tmp = 0
For i = UBound(All_C) To LBound(All_C) Step -1
    Set QCTree.Nodes(All_C(i)).Parent = QCTree.Nodes("Root")
    mdiMain.pBar.Value = tmp
    tmp = tmp + 1
Next
mdiMain.pBar.Value = mdiMain.pBar.Max
Me.AutoRedraw = True
Me.QCTree.Visible = True
End Sub

Public Sub RenumberTree()
Dim All_C()
Dim tmp, i, tmp_
Me.AutoRedraw = False
Me.QCTree.Visible = False
ReDim All_C(0)
Set tmp = QCTree.Nodes("Root")
Do While True
    If tmp Is Nothing Then Exit Do
    If Left(tmp.Key, 1) = "C" Then
        All_C(UBound(All_C)) = tmp.Key
        ReDim Preserve All_C(UBound(All_C) + 1)
    End If
    If tmp.FirstSibling.Index = 1 Then
        Set tmp = tmp.Child.FirstSibling
    Else
        Set tmp = tmp.Next
    End If
Loop
For i = LBound(All_C) To UBound(All_C) - 1
        tmp = "[" & Format(i + 1, "0000") & "]"
        tmp_ = Trim(Right(QCTree.Nodes(All_C(i)).Text, Len(QCTree.Nodes(All_C(i)).Text) - 6))
        QCTree.Nodes(All_C(i)).Text = Replace(tmp, "]]", "]") & " " & Replace(tmp_, "] ", "")
Next
Me.AutoRedraw = True
Me.QCTree.Visible = True
End Sub

Private Function GetArrInxFromInx(X, strFind)
Dim i
For i = LBound(X) To UBound(X)
    If X(i) = strFind Then
        GetArrInxFromInx = i
        Exit Function
    End If
Next
End Function

Private Sub flxImport_Click()
If flxImport.TextMatrix(flxImport.row, 0) <> "" Then
QCTree.Nodes(GetNodeKey(Format(flxImport.row, "0000"))).Selected = True
FromClick = False
End If
End Sub

Private Sub flxImport_EnterCell()
If FromClick = True Then Exit Sub
If flxImport.TextMatrix(flxImport.row, 0) <> "" Then
QCTree.Nodes(GetNodeKey(Format(flxImport.row, "0000"))).Selected = True
FromClick = False
End If
End Sub

Private Sub Form_Load()
Dim imgX As ListImage
Dim BitmapPath As String
Dim FileFunct As New clsFiles
BitmapPath = "icons\mail\mail01a.ico"
ClearForm
End Sub

Private Sub Form_Resize()
On Error Resume Next
QCTree.height = stsBar.Top - 1650
flxImport.height = stsBar.Top - 1650
flxImport.width = Me.width - flxImport.Left - 500
End Sub

Private Sub Form_Terminate()
On Error Resume Next
Unload frmControl
End Sub

Private Sub QCTree_Click()
On Error Resume Next
If Left(QCTree.SelectedItem.Key, 1) = "C" Then
    QCTree.Nodes(QCTree.SelectedItem.Index).Selected = True
    QCTree.Nodes(QCTree.SelectedItem.Index).Selected = True
    QCTree.Nodes(QCTree.SelectedItem.Index).Selected = True
    FromClick = True
    flxImport.col = 0
    flxImport.row = Val(Replace(Left(QCTree.Nodes(QCTree.SelectedItem.Index).Text, 6), "[", ""))
    flxImport.ColSel = 7
End If
End Sub

Private Sub QCTree_DblClick()
Dim tmp
On Error Resume Next
If Left(QCTree.SelectedItem.Key, 1) = "V" Then
   tmp = InputBox("Enter the Parameter Value", "Enter Parameter Value", QCTree.SelectedItem.Text)
   If Trim(tmp) <> "" Then QCTree.SelectedItem.Text = tmp: PushToFlex
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpR, i
Select Case Button.Key
Case "cmdRefresh"
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
    ClearForm
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
Case "cmdOutput"
    If QCTree.Nodes.Count > 1 Then OutputTable_XML
Case "cmdUpload"
    If QCTree.Nodes.Count > 1 Then
      If MsgBox("Are you sure you want to consolidate " & QCTree.Nodes(1).Text & "?", vbYesNo) = vbYes Then
        Randomize: tmpR = CInt(Rnd(1000) * 10000)
        If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
            CreateXML
        Else
            MsgBox "Invalid pass key", vbCritical
        End If
      End If
    End If
End Select
End Sub

Private Sub CreateXML()
Dim HEADER_ As String
Dim BODY_ As String
Dim FOOTER_ As String
Dim tmpLOg As String
Dim tmp As String
Dim i As Integer, lastCnt As Integer
Dim FileFunct As New clsFiles
mdiMain.pBar.Max = 100
mdiMain.pBar.Value = 10
tmpLOg = "<LOG_ENTRY>" & vbCrLf
tmpLOg = tmpLOg & "<ROW_COLOR>#000000</ROW_COLOR>" & vbCrLf
tmpLOg = tmpLOg & "<EXECUTION_TIME>|1|</EXECUTION_TIME>" & vbCrLf
tmpLOg = tmpLOg & "<ELAPSED_TIME>|2|</ELAPSED_TIME>" & vbCrLf
tmpLOg = tmpLOg & "<STEP_RESULT>|3|</STEP_RESULT>" & vbCrLf
tmpLOg = tmpLOg & "<STEP_SUMMARY>|4|</STEP_SUMMARY>" & vbCrLf
tmpLOg = tmpLOg & "<COMPONENT_NAME> |5| </COMPONENT_NAME>" & vbCrLf
tmpLOg = tmpLOg & "<STEP_DESCRIPTION>|6|</STEP_DESCRIPTION>" & vbCrLf
tmpLOg = tmpLOg & "|7|" & vbCrLf
tmpLOg = tmpLOg & "</LOG_ENTRY>"

HEADER_ = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?>" & vbCrLf
HEADER_ = HEADER_ & "<?xml-stylesheet href=""LOG.XSLT"" type=""Text/XSL""?>" & vbCrLf
HEADER_ = HEADER_ & "<REPORT_LOG>" & vbCrLf
HEADER_ = HEADER_ & "<TAOVERSION>TAOVERSION</TAOVERSION>" & vbCrLf
HEADER_ = HEADER_ & "<TESTINFO>" & vbCrLf

mdiMain.pBar.Value = 30

    For i = 2 To 10
        If flxImport.TextMatrix(i, 3) = "TITLE" Then
            lastCnt = i
            Exit For
        Else
            HEADER_ = HEADER_ & "<" & CleanHTML_v1(flxImport.TextMatrix(i, 3)) & ">" & flxImport.TextMatrix(i, 4) & "</" & flxImport.TextMatrix(i, 3) & ">" & vbCrLf
        End If
    Next
mdiMain.pBar.Value = 60
HEADER_ = HEADER_ & "<TESTPATH/>" & vbCrLf
HEADER_ = HEADER_ & "</TESTINFO>" & vbCrLf
HEADER_ = HEADER_ & "<TITLE>" & CleanHTML_v1(flxImport.TextMatrix(lastCnt, 4)) & "</TITLE>" & vbCrLf: lastCnt = lastCnt + 1
HEADER_ = HEADER_ & "<TABLE_HEADER_TEXT>" & CleanHTML_v1(flxImport.TextMatrix(lastCnt, 4)) & "</TABLE_HEADER_TEXT>" & vbCrLf: lastCnt = lastCnt + 1
HEADER_ = HEADER_ & "<TABLE_HEADER_TIME>" & CleanHTML_v1(flxImport.TextMatrix(lastCnt, 4)) & "</TABLE_HEADER_TIME>" & vbCrLf: lastCnt = lastCnt + 1
HEADER_ = HEADER_ & "<LOG_DATA>" & vbCrLf
    HEADER_ = HEADER_ & "<DEBUG_LOG>" & CleanHTML_v1(flxImport.TextMatrix(lastCnt, 4)) & "</DEBUG_LOG>" & vbCrLf: lastCnt = lastCnt + 1
    HEADER_ = HEADER_ & "<LOG_FOLDER>" & CleanHTML_v1(flxImport.TextMatrix(lastCnt, 4)) & "</LOG_FOLDER>" & vbCrLf: lastCnt = lastCnt + 1
    HEADER_ = HEADER_ & "<EXECUTED_BY>" & CleanHTML_v1(flxImport.TextMatrix(lastCnt, 4)) & "</EXECUTED_BY>" & vbCrLf: lastCnt = lastCnt + 1
    HEADER_ = HEADER_ & "<PC>" & CleanHTML_v1(flxImport.TextMatrix(lastCnt, 4)) & "</PC>" & vbCrLf: lastCnt = lastCnt + 1
    HEADER_ = HEADER_ & "<STATUS>" & CleanHTML_v1(flxImport.TextMatrix(lastCnt, 4)) & "</STATUS>" & vbCrLf: lastCnt = lastCnt + 1
HEADER_ = HEADER_ & "</LOG_DATA>" & vbCrLf
mdiMain.pBar.Value = 90
For i = lastCnt To flxImport.Rows - 2
    tmp = Replace(tmpLOg, "|1|", CleanHTML_v1(flxImport.TextMatrix(i, 1)))
    tmp = Replace(tmp, "|2|", CleanHTML_v1(flxImport.TextMatrix(i, 2)))
    tmp = Replace(tmp, "|3|", CleanHTML_v1(flxImport.TextMatrix(i, 3)))
    tmp = Replace(tmp, "|4|", CleanHTML_v1(flxImport.TextMatrix(i, 4)))
    tmp = Replace(tmp, "|5|", CleanHTML_v1(flxImport.TextMatrix(i, 5)))
    tmp = Replace(tmp, "|6|", CleanHTML_v1(flxImport.TextMatrix(i, 6)))
    If Trim(flxImport.TextMatrix(i, 7)) <> "" Then
        tmp = Replace(tmp, "|7|", "<IMAGE_PATH>" & CleanHTML_v1(Replace(flxImport.TextMatrix(i, 7), "Captured Image Stored in: ", "")) & "</IMAGE_PATH>")
    Else
        tmp = Replace(tmp, "|7|", "<IMAGE_PATH/>")
    End If
    BODY_ = BODY_ & tmp & vbCrLf
Next
FOOTER_ = "</REPORT_LOG>"
dlgOpenXML.ShowSave
If dlgOpenXML.filename <> "" Then
FileFunct.WriteNewFile dlgOpenXML.filename, HEADER_ & BODY_ & FOOTER_
stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = "Report Logs successfully created - " & dlgOpenXML.filename
End If
mdiMain.pBar.Value = 100
End Sub

Private Sub ClearForm()
Dim FileFunct As New clsFiles
cmdUp.Enabled = False
cmdDown.Enabled = False
cmdAdd.Enabled = False
cmdRemove.Enabled = False
Me.QCTree.Visible = False
QCTree.Nodes.Clear
Me.QCTree.Visible = True
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
flxImport.FixedCols = 1
flxImport.FixedRows = 1
flxImport.Rows = 2
End Sub

Function GetCommentText(rCommentCell As Range)
     Dim strGotIt As String
         On Error Resume Next
         strGotIt = WorksheetFunction.Clean _
             (rCommentCell.Comment.Text)
         GetCommentText = strGotIt
         On Error GoTo 0
End Function

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
  curTab = "SINGER01"
  xlObject.Sheets("Sheet1").Name = curTab
  flxImport.FixedCols = 0
  flxImport.FixedRows = 0
  flxImport.row = 0
  flxImport.col = 0
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
    
  xlObject.Workbooks(1).SaveAs "SINGER01-" & Replace(Replace(QCTree.Nodes("Root").Text, ":", "-"), "\", "-") & "-" & Format(Now, "mmddyyyy HHMM AMPM")
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate
  
  Set xlWB = Nothing
  Set xlObject = Nothing
  
  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export to MS Excel completed.": Exit Sub:
OutErr:     MsgBox Err.Description, vbCritical: xlObject.Visible = True: xlObject.ActiveWindow.Activate: Set xlWB = Nothing: Set xlObject = Nothing
On Error GoTo 0
End Sub

Private Sub QCTree_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Changed to that the item that the mouse is over is dragged, not the currently
    ' selected item which could be else where
    Set nodx = QCTree.HitTest(X, Y) ' Set the item being dragged.
    ' Make sure that the Root node is not selected
    If Not nodx Is Nothing Then
        If nodx.Parent Is Nothing Then Set nodx = Nothing
    End If
End Sub

Private Sub QCTree_DragDrop(Source As Control, X As Single, Y As Single)
    If QCTree.DropHighlight Is Nothing Then
        Set QCTree.DropHighlight = Nothing
        indrag = False
    Else
        If nodx = QCTree.DropHighlight Then Exit Sub
        Cls
        Print nodx.Text & " dropped on " & QCTree.DropHighlight.Text

        ' This line actually moves the nodes, all you have to do is change the parent to move nodes
        ' Assign the parent to nothing to move it to the root of the tree assign it's parent to Nothing
        If Not nodx.Parent Is QCTree.DropHighlight Then
            DragMove nodx.Key, QCTree.DropHighlight.Key
            'Set nodX.Parent = QCTree.DropHighlight
        End If
         
        Set QCTree.DropHighlight = Nothing
        indrag = False
    End If
End Sub

Private Sub QCTree_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = vbLeftButton And Not nodx Is Nothing Then ' Signal a Drag operation.
        indrag = True ' Set the flag to true.
        ' Set the drag icon with the CreateDragImage method.
        QCTree.DragIcon = nodx.CreateDragImage
        QCTree.Drag vbBeginDrag ' Drag operation.
    End If
End Sub


Private Sub QCTree_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If indrag = True Then
        ' Set DropHighlight to the mouse's coordinates.
        Set QCTree.DropHighlight = QCTree.HitTest(X, Y)
    End If
End Sub
