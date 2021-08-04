VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsolidate 
   Caption         =   "TAO Automation Module"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12705
   Icon            =   "frmConsolidate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   12705
   Tag             =   "Business Component Module"
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkAddYN 
      Caption         =   "Add Optional Execution"
      Height          =   255
      Left            =   6540
      TabIndex        =   14
      Top             =   1260
      Width           =   2115
   End
   Begin VB.CommandButton cmdGenerateDataTable 
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
      Left            =   9180
      Picture         =   "frmConsolidate.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Extract Test Plan tests"
      Top             =   600
      Width           =   2205
   End
   Begin VB.CommandButton cmdLoadToQC 
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
      Left            =   6900
      Picture         =   "frmConsolidate.frx":18F7
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Extract Test Plan tests"
      Top             =   600
      Width           =   2205
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
      Picture         =   "frmConsolidate.frx":2CCE
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Import step description and expected results from an excel file"
      Top             =   600
      Width           =   2205
   End
   Begin VB.Frame Frame1 
      Height          =   435
      Left            =   60
      TabIndex        =   5
      Top             =   1140
      Width           =   6375
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "New"
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
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
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
         Left            =   5100
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdLoadFromQC 
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
      Left            =   4620
      Picture         =   "frmConsolidate.frx":47DC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Extract Test Plan tests"
      Top             =   600
      Width           =   2205
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
      Picture         =   "frmConsolidate.frx":5C68
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
      Width           =   12705
      _ExtentX        =   22410
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
            Picture         =   "frmConsolidate.frx":6F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsolidate.frx":71C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsolidate.frx":C9B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsolidate.frx":CC44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView QCTree 
      Height          =   4335
      Left            =   60
      TabIndex        =   0
      Top             =   1620
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   7646
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
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
            Picture         =   "frmConsolidate.frx":CED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsolidate.frx":D5E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsolidate.frx":DCF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsolidate.frx":E408
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
      Top             =   6060
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmConsolidate.frx":EB1A
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   21184
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
            Picture         =   "frmConsolidate.frx":F06B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsolidate.frx":F34D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsolidate.frx":F89E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsolidate.frx":FDEF
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
End
Attribute VB_Name = "frmConsolidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BComponent
    Group_No As Integer
    Component_Path As String
    Component_Name As String
    Component_Description As String
    Scripter As String
    Peer_Reviewer As String
    QA_Reviewer As String
    Planned_Start_Date As String
    Planned_End_Date As String
    Status As String
    Parameters() As String
    DefaultValue() As String
    Log As String
End Type

Private Type NodeType
   Key As String
   Tag As String
   Text As String
End Type

Dim indrag As Boolean ' Flag that signals a Drag Drop operation.
Dim nodx As Node ' Item that is being dragged. -- Changed to Node
Dim tmpPars() As TestPlan_Pars

Private Sub chkAddYN_Click()
If chkAddYN.Value = Checked Then
    If MsgBox("Are you sure you want to add optional execution components?", vbYesNo) = vbYes Then
        chkAddYN.Value = Checked
    Else
        chkAddYN.Value = Unchecked
    End If
Else
    chkAddYN.Value = Unchecked
End If
End Sub

Private Sub cmdAdd_Click()
Control_Auto = False
If REALTIME = True Then
frmControlRunTime.Show
Else
frmControl.Show
End If
End Sub

Private Sub cmdGenerateDataTable_Click()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim i, Protections
Dim curTab, dt_cnt As Integer
Dim w, cnt, curBC, curOrder, curID, tmpNode, tmpFName

If QCTree.Nodes.Count > 1 Then
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
    
    xlObject.Sheets("Sheet2").Range("A6").Value = "Filename:"
    xlObject.Sheets("Sheet2").Range("B6").Value = QCTree.Nodes(1).Text
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
  curTab = "SAPTAO01"
  xlObject.Sheets("Sheet1").Name = curTab

    cnt = 1
    Set tmpNode = QCTree.Nodes("Root")
    Do While True
        If tmpNode Is Nothing Then Exit Do
        If Left(tmpNode.Key, 1) = "C" Then
            If Not (tmpNode.Child.FirstSibling Is Nothing) Then Set tmpNode = tmpNode.Child.FirstSibling
        ElseIf Left(tmpNode.Key, 1) = "N" Then
            If Not (tmpNode.Child.FirstSibling Is Nothing) Then Set tmpNode = tmpNode.Child.FirstSibling
        ElseIf Left(tmpNode.Key, 1) = "V" Then
            If Left(tmpNode.Text, 3) = "DT_" Then
                xlObject.Sheets(curTab).Range(ColumnLetter(CInt(cnt)) & 1) = tmpNode.Text
                xlObject.Sheets(curTab).Range(ColumnLetter(CInt(cnt)) & 2) = tmpNode.Tag
                cnt = cnt + 1
            End If
            If Not (tmpNode.Parent.Next Is Nothing) Then
                Set tmpNode = tmpNode.Parent.Next
            Else
                Set tmpNode = tmpNode.Parent.Parent.Next
            End If
        ElseIf tmpNode.FirstSibling.Index = 1 Then
            Set tmpNode = tmpNode.Child.FirstSibling
        End If
    Loop
'On Error Resume Next
    xlObject.Sheets(curTab).Select
    xlObject.Sheets(curTab).Range("A:" & ColumnLetter(CInt(cnt))).Select
        
    xlObject.Sheets(curTab).Range("A:" & ColumnLetter(CInt(cnt))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlObject.Sheets(curTab).Range("A:" & ColumnLetter(CInt(cnt))).Borders(xlDiagonalUp).LineStyle = xlNone
    With xlObject.Sheets(curTab).Range("A:" & ColumnLetter(CInt(cnt))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:" & ColumnLetter(CInt(cnt))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:" & ColumnLetter(CInt(cnt))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:" & ColumnLetter(CInt(cnt))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:" & ColumnLetter(CInt(cnt))).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:" & ColumnLetter(CInt(cnt))).Borders(xlInsideHorizontal)
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
    xlObject.Sheets(curTab).Range("A:" & ColumnLetter(CInt(cnt))).Select
    xlObject.Sheets(curTab).Range("A:" & ColumnLetter(CInt(cnt))).EntireColumn.AutoFit
    xlObject.Sheets(curTab).Range("A1").Select
    
    xlObject.Sheets(curTab).Range("A1").AddComment
    xlObject.Sheets(curTab).Range("A1").Comment.Visible = False
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & "TAO Automation Module " & curDomain & "-" & curProject & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""
    
  tmpFName = "SAPTAO01-" & CleanTheString_PARAMS(QCTree.Nodes(1).Text) & "-" & Format(Now, "mmddyyyy HHMM AMPM")
  On Error Resume Next
  xlObject.Workbooks(1).SaveAs tmpFName
  dlgOpenExcel.filename = tmpFName: dlgOpenExcel.ShowSave
  If dlgOpenExcel.filename <> "" Then xlObject.Workbooks(1).SaveAs dlgOpenExcel.filename
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate
  
  Set xlWB = Nothing
  Set xlObject = Nothing
  FXGirl.EZPlay FXExportToExcel
  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export to MS Excel completed.": Exit Sub:
OutErr:     MsgBox Err.Description, vbCritical: xlObject.Visible = True: xlObject.ActiveWindow.Activate: Set xlWB = Nothing: Set xlObject = Nothing
On Error GoTo 0
End If
End Sub

Private Sub cmdImportQTP_Click()
Dim FolderName, filename, tmp, AllSteps() As QTP_Script, tmp2
Dim fileFunct As New clsFiles, i
dlgQTP.filename = "": dlgQTP.ShowOpen
If dlgQTP.filename <> "" Then
    QCTree.Visible = False
    QCTree.Nodes.Clear
    QCTree.Nodes.Add , , "Root", Replace(dlgQTP.FileTitle, ".usr", ""), 1
    QCTree.Nodes(1).Selected = True
    filename = dlgQTP.FileTitle
    FolderName = Replace(dlgQTP.filename, filename, "")
    filename = FolderName & "Action1\Script.mts"
    GetURIs FolderName & "Action1\ObjectRepository.bdb"
    tmp = fileFunct.ReadFromFileToArray(CStr(filename))
    ReDim AllSteps(0)
    For i = LBound(tmp) To UBound(tmp)
        If Trim(tmp(i)) <> "" Then
            tmp2 = Trim(GetObjText(CStr(tmp(i))))
            AllSteps(UBound(AllSteps)).Class = GetClass(CStr(tmp2))
            AllSteps(UBound(AllSteps)).Name = GetObjName(CStr(tmp2))
            AllSteps(UBound(AllSteps)).Action = GetObjAction(CStr(tmp2))
            AllSteps(UBound(AllSteps)).Value = GetObjValue(CStr(tmp2))
            AllSteps(UBound(AllSteps)).URI = GetURI(AllSteps(UBound(AllSteps)).Name)
            ReDim Preserve AllSteps(UBound(AllSteps) + 1)
        End If
    Next
    ReDim Preserve AllSteps(UBound(AllSteps) - 1)
    tmp2 = ""
    Control_Auto = True
    For i = LBound(AllSteps) To UBound(AllSteps)
        tmp2 = tmp2 & "[" & Format(i, "000") & "] " & "(" & Trim(AllSteps(i).Action) & ")" & vbCrLf & "[" & Format(i, "000") & "] " & "{" & Trim(AllSteps(i).Value) & "}" & vbCrLf & "[" & Format(i, "000") & "] " & "|micclass:=" & AllSteps(i).Class & ";logical name/text/name:=" & AllSteps(i).Name & ";index:=0|" & vbCrLf & vbCrLf
        If Left(AllSteps(i).Class, 6) = "SAPGui" Then
            frmControl.txtProperties.Text = AllSteps(i).URI
            frmControl.tmpProperties_ = frmControl.txtProperties.Text
            frmControl.txtFieldName.Text = AllSteps(i).Name
            If Trim(AllSteps(i).Action) = "Set" Then
                frmControl.txtFilter.Text = "SetText"
            ElseIf Trim(AllSteps(i).Action) = "Click" Then
                frmControl.txtFilter.Text = "Press"
            ElseIf Trim(AllSteps(i).Action) = "SendKey" Then
                frmControl.txtFilter.Text = "PressKey"
            Else
                frmControl.txtFilter.Text = Trim(AllSteps(i).Action)
            End If
        Else
            frmControl.txtProperties.Text = "micclass:=" & AllSteps(i).Class & ";logical name/text/name:=" & AllSteps(i).Name & ";index:=0"
            frmControl.tmpProperties_ = frmControl.txtProperties.Text
            frmControl.txtFieldName.Text = AllSteps(i).Name
            frmControl.optTag(frmControl_Option).Value = True
            If Trim(AllSteps(i).Action) = "Set" Then
                frmControl.txtFilter.Text = "Control_Set"
            ElseIf Trim(AllSteps(i).Action) = "Click" Then
                frmControl.txtFilter.Text = "Control_Click"
            ElseIf Trim(AllSteps(i).Action) = "Select" Then
                frmControl.txtFilter.Text = "Control_Select"
            Else
                frmControl.txtFilter.Text = Trim(AllSteps(i).Action)
            End If
        End If
        frmControl.txtValue.Text = Trim(Trim(AllSteps(i).Value))
        frmControl.txtFilter_KeyPress 13
        frmControl.Show 1
    Next
    fileFunct.WriteNewFile App.path & "\SQC Logs" & "\" & Replace(dlgQTP.FileTitle, ".usr", "") & ".txt", CStr(tmp2)
    QCTree.Visible = True
    cmdUp.Enabled = True
    cmdDown.Enabled = True
    cmdAdd.Enabled = True
    cmdRemove.Enabled = True
    MsgBox "QTP Script extracted to " & App.path & "\SQC Logs" & "\" & Replace(dlgQTP.FileTitle, ".usr", "") & ".txt"
End If
End Sub

Private Function GetURI(LogicalName As String)
Dim i
For i = LBound(All_URI) To UBound(All_URI)
    If All_URI(i).LogicalName = LogicalName Then
        GetURI = All_URI(i).URI
        Exit Function
    End If
Next
End Function

Private Function GetClass(X As String) As String
    Dim i, tmp, L
    For i = 1 To Len(X)
        L = Mid(X, i, 1)
        If L = "(" Then
            GetClass = tmp
            Exit Function
        Else
            tmp = tmp & L
        End If
    Next
End Function

Private Function GetObjName(X As String) As String
    Dim i, tmp, L, Start, parOpen, parClosed, parClosedCount
    Dim stringFunct As New clsStrings
    Start = InStr(1, X, "(")
    parOpen = 1
    parClosed = 0
    parClosedCount = stringFunct.CountInstance(X, ")")
    For i = Start + 1 To Len(X)
        L = Mid(X, i, 1)
        If L = ")" Then
            parClosed = parClosed + 1
            If L = "(" Then parOpen = parOpen + 1
            tmp = tmp & L
            If parOpen = parClosed Or parClosedCount = parClosed Then
                GetObjName = Left(Replace(tmp, """", ""), Len(Replace(tmp, """", "")) - 1)
                Exit Function
            End If
        Else
            If L = "(" Then parOpen = parOpen + 1
            tmp = tmp & L
        End If
    Next
    GetObjName = Replace(tmp, """", "")
End Function

Private Function GetObjAction(X As String) As String
    Dim i, tmp, L, Start
    tmp = ""
    Start = InStr(1, X, """)") + 2
    For i = Start + 1 To Len(X)
        L = Mid(X, i, 1)
        If L = " " Then
            GetObjAction = Replace(tmp, ".", "")
            Exit Function
        Else
            tmp = tmp & L
        End If
    Next
    GetObjAction = Replace(tmp, ".", "")
End Function

Private Function GetObjValue(X As String) As String
    Dim tmp
    tmp = GetClass(X) & "(""" & GetObjName(X) & """)." & GetObjAction(X)
    GetObjValue = Replace(Replace(X, tmp, ""), """", "")
End Function

Private Function GetObjText(X As String)
    Dim tmpString, tmp, i, Go, L
    If InStr(1, X, "@@") <> 0 Then
        tmpString = Trim(Left(X, InStr(1, X, "@@") - 1))
    Else
        tmpString = Trim(X)
    End If
    For i = Len(tmpString) To 1 Step -1
        L = Mid(tmpString, i, 1)
        If L = "(" Then Go = True
        If L = "." And Go = True Then
            GetObjText = tmp
            Exit Function
        Else
            tmp = L & tmp
        End If
    Next
End Function

Private Sub cmdLoadExcel_Click()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim fname As String
Dim lastrow
Dim i, j, tmpParam, k, X, Node_BC, Node_Param, Y
Dim tmpSts
Dim strFunct As New clsFiles
Dim stringFunct As New clsStrings
Dim intFunct As New clsInternet
Dim oldNode, tmpName

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
         If InStr(1, GetCommentText(.Range("A1")), "TAO Automation Module " & curDomain & "-" & curProject) = 0 Then
            MsgBox "Import file is invalid. Please use only sheets generated by the SuperQC"
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
         lastrow = .Range("A" & .Rows.Count).End(xlUp).row
        '.Range("A3:M" & LastRow).Copy 'Set selection to Copy
        
        'A - Load HPQC Folder Path
        'Should not be blank
        mdiMain.pBar.Max = lastrow + 2
        j = -1
        k = -1
        X = X + 1
        Y = 0
        oldNode = X
        ReDim All_BC_Param(0)
        QCTree.Visible = False
        QCTree.Nodes.Clear
        If Trim(tmpName) = "" Then
            tmpName = Replace(Replace(Replace(dlgOpenExcel.FileTitle, ".xlsx", ""), ".xls", ""), "SAPTAO01-", "")
            QCTree.Nodes.Add , , "Root", Left(tmpName, Len(tmpName) - 17), 1
        Else
            QCTree.Nodes.Add , , "Root", Trim(tmpName), 1
        End If
        QCTree.Nodes.Item(X).Selected = True
        For i = 2 To lastrow
            If .Range("A" & i).Value <> .Range("A" & i - 1).Value Then
                k = k + 1
                j = -1
                Node_BC = X
                X = 1
                QCTree.Nodes.Item(X).Selected = True
                QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("C" & k), "[" & Format(Trim(.Range("A" & i).Value), "000") & "] " & Trim(.Range("C" & i).Value), 4
                X = Node_BC
                X = QCTree.Nodes.Count
                j = j + 1
                Y = Y + 1
                QCTree.Nodes.Item(X).Selected = True
                QCTree.Nodes.Item(X).Tag = Trim(.Range("B" & i).Value)
                QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("N" & Y), Trim(.Range("D" & i).Value), 2
                Node_BC = X
                QCTree.Nodes.Item(X + 1).Selected = True
                QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("V" & Y), Trim(.Range("E" & i).Value), 3
                QCTree.Nodes.Item(Node_BC).Expanded = False
                QCTree.Nodes.Item(X + 1).Selected = True
                QCTree.SelectedItem.Tag = Trim(.Range("F" & i).Value)
                QCTree.Nodes.Item(Node_BC).Expanded = False
            Else
                j = j + 1
                X = Node_BC
                Y = Y + 1
                QCTree.Nodes.Item(X).Selected = True
                QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("N" & Y), Trim(.Range("D" & i).Value), 2
                X = QCTree.Nodes.Count
                QCTree.Nodes.Item(X).Selected = True
                QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("V" & Y), Trim(.Range("E" & i).Value), 3
                QCTree.Nodes.Item(Node_BC).Expanded = False
                QCTree.Nodes.Item(X + 1).Selected = True
                QCTree.SelectedItem.Tag = Trim(.Range("F" & i).Value)
                QCTree.Nodes.Item(Node_BC).Expanded = False
                X = Node_BC
            End If
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = i - 1 & " out of " & lastrow - 1 & " validated " & Format(i / lastrow, "0.0%") & " (" & tmpSts & ") errors found."
            mdiMain.pBar.Value = i
        Next
        
    End With
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

Private Sub cmdLoadFromQC_Click()
frmLoadTAOBC.Show 1
End Sub

Private Sub cmdLoadToQC_Click()
Dim tp_id As String, tp_name As String, tmpNode, BC_ID As String, tmpR, i
Dim LinkStack()
If QCTree.Nodes.Count > 1 Then
      If MsgBox("Are you sure you want to consolidate " & QCTree.Nodes(1).Text & " to " & BCTARGET & "?", vbYesNo) = vbYes Then
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
            ReDim tmpPars(0)
            tp_name = CleanTheString(QCTree.Nodes(1).Text & " " & Format(Now, "mmmddyyyy hhmmss"))
            If QCTree.Nodes.Count > 1 Then
                frmQCDialogBox.curSaveModule = "TESTPLAN"
                frmQCDialogBox.txtName.Text = tp_name
                frmQCDialogBox.Show 1
                Randomize: tmpR = CInt(Rnd(1000) * 10000)
                If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
                    If frmQCDialogBox.curOutput_ID = "NEW" Then
                        tp_id = CreateTest(frmQCDialogBox.curOutput_Path, CleanTheString(frmQCDialogBox.curOutput_FName))
                    Else
                        RemoveBPTTest frmQCDialogBox.curOutput_ID
                        tp_id = frmQCDialogBox.curOutput_ID
                    End If
                        Set tmpNode = QCTree.Nodes("Root")
                        mdiMain.pBar.Max = QCTree.Nodes.Count + 10
                        ReDim LinkStack(0)
                        Do While True
                            If tmpNode Is Nothing Then Exit Do
                            If Left(tmpNode.Key, 1) = "C" Then
                                BC_ID = tmpNode.Tag
                                tmpPars(UBound(tmpPars)).BC_ID = BC_ID
                                LinkStack(UBound(LinkStack)) = tmpNode.Tag
                                ReDim Preserve LinkStack(UBound(LinkStack) + 1)
                                'LinkBPTTest tp_id, tmpNode.Tag
                                If Not (tmpNode.Child.FirstSibling Is Nothing) Then Set tmpNode = tmpNode.Child.FirstSibling
                            ElseIf Left(tmpNode.Key, 1) = "N" Then
                                tmpPars(UBound(tmpPars)).Name = tmpNode.Text
                                 tmpPars(UBound(tmpPars)).ID = BC_ID & tmpPars(UBound(tmpPars)).Name
                                If Not (tmpNode.Child.FirstSibling Is Nothing) Then Set tmpNode = tmpNode.Child.FirstSibling
                            ElseIf Left(tmpNode.Key, 1) = "V" Then
                                tmpPars(UBound(tmpPars)).Value = tmpNode.Text
                                tmpPars(UBound(tmpPars)).ID = BC_ID & tmpPars(UBound(tmpPars)).Name
                                ReDim Preserve tmpPars(UBound(tmpPars) + 1)
                                If Not (tmpNode.Parent.Next Is Nothing) Then
                                    Set tmpNode = tmpNode.Parent.Next
                                Else
                                    Set tmpNode = tmpNode.Parent.Parent.Next
                                End If
                                On Error Resume Next
                                mdiMain.pBar.Value = mdiMain.pBar.Value + 1
                                On Error GoTo 0
                            ElseIf tmpNode.FirstSibling.Index = 1 Then
                                Set tmpNode = tmpNode.Child.FirstSibling
                            End If
                        Loop
                        ReDim Preserve tmpPars(UBound(tmpPars) - 1)
                        ReDim Preserve LinkStack(UBound(LinkStack) - 1)
                        For i = LBound(LinkStack) To UBound(LinkStack)
                            If chkAddYN.Value = Checked Then
                                LinkBPTTest tp_id, CStr(LinkStack(i))
                                PromoteParamBPTTestRunBC tp_id, CStr(LinkStack(i)), i + 1 & "_" & CleanTheString_PARAMS(GetBusinessComponentName(CStr(LinkStack(i)))), CInt(i) + 1
                            End If
                            LinkBPTTest tp_id, CStr(LinkStack(i))
                            If chkAddYN.Value = Checked Then
                                LinkBPTTest tp_id, CStr(LinkStack(i))
                            End If
                        Next
                        PromoteParamBPTTest tp_id
                        stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Test Plan " & tp_name & " successfully created in " & TPTARGET
                        FXGirl.EZPlay FXScriptConsolidationCompleted
                        mdiMain.pBar.Value = mdiMain.pBar.Max
                Else
                    MsgBox "Invalid pass key", vbCritical
            End If
          End If
        End If
End If
End Sub

Private Sub cmdNew_Click()
Dim tmp
tmp = InputBox("Enter Consolidated Component name", "New Consolidated Component", "New Component")
tmp = CleanTheString(tmp)
If Trim((tmp)) <> "" Then
        Me.AutoRedraw = False
        QCTree.Visible = False
        QCTree.Nodes.Clear
        QCTree.Nodes.Add , , "Root", tmp, 1
        QCTree.Nodes.Item("Root").Selected = True
        cmdUp.Enabled = True
        cmdDown.Enabled = True
        cmdAdd.Enabled = True
        cmdRemove.Enabled = True
        Me.AutoRedraw = True
        QCTree.Visible = True
        Control_Auto = False
End If
End Sub

Private Sub cmdRemove_Click()
Dim All_C()
Dim tmp, i, tmp_
Me.AutoRedraw = False
On Error GoTo Err1
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
                tmp = "[" & Format(i + 1, "000") & "]"
                tmp_ = Trim(Right(QCTree.Nodes(All_C(i)).Text, Len(QCTree.Nodes(All_C(i)).Text) - 5))
                QCTree.Nodes(All_C(i)).Text = Replace(tmp, "]]", "]") & " " & Replace(tmp_, "] ", "")
        Next
    End If
End If
Me.AutoRedraw = True
Me.QCTree.Visible = True
Exit Sub
Err1:
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

    tmp = Trim(Left(QCTree.Nodes(All_C(NodeA)).Text, 5))
    tmp_ = Trim(Right(QCTree.Nodes(All_C(NodeA - 1)).Text, Len(QCTree.Nodes(All_C(NodeA - 1)).Text) - 5))
    QCTree.Nodes(All_C(NodeA)).Text = Trim(Left(QCTree.Nodes(All_C(NodeA - 1)).Text, 5)) & " " & Trim(Right(QCTree.Nodes(All_C(NodeA)).Text, Len(QCTree.Nodes(All_C(NodeA)).Text) - 5))
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
Me.AutoRedraw = True
Me.QCTree.Visible = True
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

    tmp = Trim(Left(QCTree.Nodes(All_C(NodeA)).Text, 5))
    tmp_ = Trim(Right(QCTree.Nodes(All_C(NodeA + 1)).Text, Len(QCTree.Nodes(All_C(NodeA + 1)).Text) - 5))
    QCTree.Nodes(All_C(NodeA)).Text = Trim(Left(QCTree.Nodes(All_C(NodeA + 1)).Text, 5)) & " " & Trim(Right(QCTree.Nodes(All_C(NodeA)).Text, Len(QCTree.Nodes(All_C(NodeA)).Text) - 5))
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
Me.AutoRedraw = True
Me.QCTree.Visible = True
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
        tmp = "[" & Format(i + 1, "000") & "]"
        tmp_ = Trim(Right(QCTree.Nodes(All_C(i)).Text, Len(QCTree.Nodes(All_C(i)).Text) - 5))
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
        tmp = "[" & Format(i + 1, "000") & "]"
        tmp_ = Trim(Right(QCTree.Nodes(All_C(i)).Text, Len(QCTree.Nodes(All_C(i)).Text) - 5))
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
        tmp = "[" & Format(i + 1, "000") & "]"
        tmp_ = Trim(Right(QCTree.Nodes(All_C(i)).Text, Len(QCTree.Nodes(All_C(i)).Text) - 5))
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

Private Sub Form_Load()
Dim imgX As ListImage
Dim BitmapPath As String
Dim fileFunct As New clsFiles
BitmapPath = "icons\mail\mail01a.ico"
SPEF = fileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦SPEF" & curDomain & "-" & curProject & "¦")
If fileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦REALTIME" & curDomain & "-" & curProject & "¦") = "01" Then
    REALTIME = True
    LoadAllComponentsFromQC
Else
    REALTIME = False
    LoadAllBusinessComponents
End If
TPTARGET = fileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦TPTARGET" & curDomain & "-" & curProject & "¦")
ClearForm
End Sub

Private Sub Form_Resize()
On Error Resume Next
QCTree.height = stsBar.Top - 1650
QCTree.width = Me.width - QCTree.Left - 200
End Sub

Private Sub Form_Terminate()
On Error Resume Next
Unload frmControl
End Sub

Private Sub QCTree_Click()
On Error Resume Next
QCTree.Nodes(QCTree.SelectedItem.Index).Selected = True
End Sub

Private Sub QCTree_DblClick()
Dim tmp
If Left(QCTree.SelectedItem.Key, 1) = "V" Then
   tmp = InputBox("Enter the Parameter Value", "Enter Parameter Value", QCTree.SelectedItem.Text)
   If Trim(tmp) <> "" Then QCTree.SelectedItem.Text = tmp
ElseIf QCTree.SelectedItem.Key = "Root" Then
    tmp = InputBox("Enter the Component Name", "Enter Component Name", QCTree.SelectedItem.Text)
   If Trim(tmp) <> "" Then QCTree.SelectedItem.Text = tmp
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpR, i
Select Case Button.Key
Case "cmdRefresh"
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
    ClearForm
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
Case "cmdGenerate"
    On Error GoTo OutputErr1
    frmSourceBusinessComponents.Show 1
    Exit Sub
OutputErr1:
    MsgBox "Data was been truncated because of an error." & vbCrLf & Err.Description
Case "cmdOutput"
    On Error GoTo OutputErr2
    If QCTree.Nodes.Count > 1 Then OutputTable
    Exit Sub
OutputErr2:
    MsgBox "Data was been truncated because of an error." & vbCrLf & Err.Description
Case "cmdUpload"
    On Error GoTo OutputErr3
    If QCTree.Nodes.Count > 1 Then
      If MsgBox("Are you sure you want to consolidate " & QCTree.Nodes(1).Text & " to " & BCTARGET & "?", vbYesNo) = vbYes Then
        Randomize: tmpR = CInt(Rnd(1000) * 10000)
        If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
            UploadBC
        Else
            MsgBox "Invalid pass key", vbCritical
        End If
      End If
    End If
    Exit Sub
OutputErr3:
    MsgBox "Data was been truncated because of an error." & vbCrLf & Err.Description
End Select
End Sub

Private Sub UploadBC()
Dim tmpCode As String, fileFunct As New clsFiles
Dim BC_ID As String, BC_Order As String, BC_Name As String
Dim tmpBC As BComponent, cnt, tmp, tmpNode, DumpedList As String
If QCTree.Nodes.Count > 1 Then
        frmQCDialogBox.curSaveModule = "BCCOMP"
        frmQCDialogBox.txtName.Text = CleanTheString(QCTree.Nodes(1).Text & " " & Format(Now, "mmmddyyyy hhmmss"))
        frmQCDialogBox.Show 1
        If Trim(frmQCDialogBox.curOutput_ID) = "" Then Exit Sub
        If Trim(frmQCDialogBox.curOutput_FName) = "" Then Exit Sub
        If MsgBox("Are you sure you want to save " & frmQCDialogBox.curOutput_FName & "?", vbYesNo) <> vbYes Then Exit Sub
        ReDim tmpBC.Parameters(0)
        ReDim tmpBC.DefaultValue(0)
        mdiMain.pBar.Max = QCTree.Nodes("Root").Children
        mdiMain.pBar.Value = 0
        Set tmpNode = QCTree.Nodes("Root")
            Do While True
                If tmpNode Is Nothing Then Exit Do
                If Left(tmpNode.Key, 1) = "C" Then
                    cnt = cnt + 1
                    tmpBC.Component_Description = Right(tmpNode.Text, Len(tmpNode.Text) - 6)
                    If Not (tmpNode.Child.FirstSibling Is Nothing) Then Set tmpNode = tmpNode.Child.FirstSibling
                ElseIf Left(tmpNode.Key, 1) = "N" Then
                    BC_Order = (Replace(Left(tmpNode.Text, 4), "[", ""))
                    BC_Name = tmpNode.Text
                    tmpBC.Parameters(UBound(tmpBC.Parameters)) = "C" & Format(cnt, "000") & "_" & CleanTheString_PARAMS(BC_Name)
                    ReDim Preserve tmpBC.Parameters(UBound(tmpBC.Parameters) + 1)
                    If Not (tmpNode.Child.FirstSibling Is Nothing) Then Set tmpNode = tmpNode.Child.FirstSibling
                ElseIf Left(tmpNode.Key, 1) = "V" Then
                    tmpBC.DefaultValue(UBound(tmpBC.DefaultValue)) = tmpNode.Text
                    ReDim Preserve tmpBC.DefaultValue(UBound(tmpBC.DefaultValue) + 1)
                    If Not (tmpNode.Parent.Next Is Nothing) Then
                        Set tmpNode = tmpNode.Parent.Next
                    Else
                        Set tmpNode = tmpNode.Parent.Parent.Next
                    End If
                ElseIf tmpNode.FirstSibling.Index = 1 Then
                    Set tmpNode = tmpNode.Child.FirstSibling
                End If
                mdiMain.pBar.Value = cnt
            Loop
            mdiMain.pBar.Value = mdiMain.pBar.Max - 1
            Set tmpNode = QCTree.Nodes("Root")
            Do While True
                If tmpNode Is Nothing Then Exit Do
                If Left(tmpNode.Key, 1) = "C" Then
                    BC_ID = tmpNode.Tag
                    BC_Order = CInt(Replace(Left(tmpNode.Text, 4), "[", ""))
                    BC_Name = Right(tmpNode.Text, Len(tmpNode.Text) - 6)
                    If REALTIME = True Then
                        If InStr(1, DumpedList, BC_ID, vbTextCompare) = 0 Then
                            DumpBusinessComponent CStr(BC_ID), CStr(BC_Name)
                            DumpedList = DumpedList & BC_ID & " "
                            DumpedList = Trim(DumpedList)
                        End If
                    End If
                    tmpCode = tmpCode & Load_BC(BC_ID, CInt(BC_Order), BC_Name, tmpBC) & vbCrLf & vbCrLf
                    Set tmpNode = tmpNode.Next
                ElseIf tmpNode.FirstSibling.Index = 1 Then
                    Set tmpNode = tmpNode.Child.FirstSibling
                End If
            Loop
            mdiMain.pBar.Value = mdiMain.pBar.Max / 2
        If frmQCDialogBox.curOutput_ID = "NEW" Then
            tmpBC.Component_Name = CleanTheString(frmQCDialogBox.curOutput_FName)
            FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[CONSOLIDATE START: " & Now & " " & tmpBC.Component_Name & "] "
            tmpBC.Component_Description = CleanTheString(QCTree.Nodes(1).Text & " " & Format(Now, "mmmddyyyy hhmmss"))
            tmpBC.Component_Path = frmQCDialogBox.curOutput_Path
            tmpBC.Peer_Reviewer = curUser
            tmpBC.QA_Reviewer = curUser
            tmpBC.Scripter = curUser
            tmpBC.Planned_Start_Date = Format(Now, "dd/mm/yyyy")
            tmpBC.Planned_End_Date = Format(Now, "dd/mm/yyyy")
            tmpBC.Status = "040 Ready For QA Review"
            fileFunct.WriteNewFile App.path & "\SQC DAT\BC Template\Action1\Script.mts", tmpCode
            tmp = Create_New_Component(tmpBC)
            Save_BC tmp
        Else
            tmp = (frmQCDialogBox.curOutput_ID)
            tmpBC.Component_Description = CleanTheString(QCTree.Nodes(1).Text & " " & Format(Now, "mmmddyyyy hhmmss"))
            tmpBC.Component_Path = frmQCDialogBox.curOutput_Path
            tmpBC.Peer_Reviewer = curUser
            tmpBC.QA_Reviewer = curUser
            tmpBC.Scripter = curUser
            tmpBC.Planned_Start_Date = Format(Now, "dd/mm/yyyy")
            tmpBC.Planned_End_Date = Format(Now, "dd/mm/yyyy")
            tmpBC.Status = "040 Ready For QA Review"
            FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[CONSOLIDATE START: " & Now & " " & tmpBC.Component_Name & "] "
            fileFunct.WriteNewFile App.path & "\SQC DAT\BC Template\Action1\Script.mts", tmpCode
            Update_Component tmpBC, tmp
            Save_BC tmp
        End If
        stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Component " & tmpBC.Component_Name & " successfully created in " & BCTARGET
        FXGirl.EZPlay FXScriptConsolidationCompleted
        mdiMain.pBar.Value = mdiMain.pBar.Max
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[CONSOLIDATE END: " & Now & " " & tmpBC.Component_Name & "] "
End If
End Sub

'########################### Create New Component ###########################
Private Function Create_New_Component(tmpComp As BComponent)
    Dim myComp As Component, j
    Dim compFactory As ComponentFactory
    Dim compParamFactory As ComponentParamFactory
    Dim compParam() As ComponentParam
    Dim generalComponentFolderFactory As ComponentFolderFactory
    Dim cFolder As ComponentFolder, rootCFolder As ComponentFolder
    Dim tmpList() As String
    On Error GoTo NewComponentErr
    ' Get Component Folder
    ' Get a ComponentFolderFactory from the QCConnectiononnection object
    Set generalComponentFolderFactory = QCConnection.ComponentFolderFactory
    Set rootCFolder = generalComponentFolderFactory.FolderByPath(tmpComp.Component_Path)
    If IsMissing(rootCFolder) Then
        ' Get the root folder
        Set rootCFolder = generalComponentFolderFactory.Root
    End If
    ' Get a ComponentFactory
    Set compFactory = rootCFolder.ComponentFactory
    ' Add the component
    Set myComp = compFactory.AddItem(Null)
    Dim errString As String
    If (compFactory.IsComponentNameValid(tmpComp.Component_Name, errString)) Then
        myComp.Name = tmpComp.Component_Name
    End If
    myComp.Field("CO_RESPONSIBLE") = tmpComp.Scripter  'Scripter
    myComp.Field("CO_DESC") = "" 'tmpComp.Component_Description   'Component Description
    myComp.Field("CO_USER_TEMPLATE_01") = tmpComp.Status  'Status
    myComp.Field("CO_USER_TEMPLATE_02") = tmpComp.Peer_Reviewer  'Peer Reviewer
    myComp.Field("CO_USER_TEMPLATE_03") = tmpComp.QA_Reviewer  'QA Reviewer
    myComp.Field("CO_USER_TEMPLATE_05") = tmpComp.Planned_End_Date  'Planned End
    myComp.Field("CO_USER_TEMPLATE_04") = tmpComp.Planned_Start_Date  'Planned Start
    myComp.Field("CO_USER_TEMPLATE_08") = "AUTOMATED"
    myComp.ScriptType = "QT-SCRIPTED"
    myComp.ApplicationAreaID = 11866
    myComp.Post
    Create_New_Component = myComp.ID
    'Return the new component.
    Set compParamFactory = myComp.ComponentParamFactory
    ReDim tmpList(0)
    mdiMain.pBar.Max = UBound(tmpComp.Parameters) + 1
    For j = LBound(tmpComp.Parameters) To UBound(tmpComp.Parameters)
        If tmpComp.Parameters(j) <> "" And IsParameterDeclared(tmpList, Replace(tmpComp.Parameters(j), "-", "_")) = False And Trim(tmpComp.DefaultValue(j)) <> "" Then
            ReDim Preserve compParam(j)
            Set compParam(j) = compParamFactory.AddItem(Null)
            compParam(j).IsOut = 0
            compParam(j).Name = Replace((tmpComp.Parameters(j)), "-", "_")
            compParam(j).Desc = (tmpComp.Parameters(j))
            compParam(j).Value = tmpComp.DefaultValue(j)
            compParam(j).ValueType = "String"
            compParam(j).Order = j + 1
            compParam(j).Post
            tmpList(UBound(tmpList)) = compParam(j).Name
            ReDim Preserve tmpList(UBound(tmpList) + 1)
        End If
        Debug.Print j
        mdiMain.pBar.Value = j
        If mdiMain.pBar.Max > 30 Then
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
    ReDim Preserve compParam(j)
    Set compParam(j) = compParamFactory.AddItem(Null)
    compParam(j).IsOut = 0
    compParam(j).Name = "EMPTY"
    compParam(j).Desc = "Placeholder for all EMPTY parameters"
    compParam(j).Value = ""
    compParam(j).ValueType = "String"
    compParam(j).Order = j + 1
    compParam(j).Post
    tmpList(UBound(tmpList)) = compParam(j).Name
    ReDim Preserve tmpList(UBound(tmpList) + 1)
    myComp.Post
    myComp.UnLockObject
    Create_New_Component = myComp.ID
Exit Function
NewComponentErr:
    Dim tmpFile As New clsFiles
    Debug.Print "Error in creating component in line " & tmpComp.Component_Name
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[BC CREATE: (FAILED) " & Now & " " & tmpComp.Component_Name & "] " & Err.Description
End Function
'########################### End Of Create New Component ##########################

'########################### Update Component ###########################
Private Function Update_Component(tmpComp As BComponent, compID)
    Dim myComp As Component, j
    Dim compFactory As ComponentFactory
    Dim compParamFactory As ComponentParamFactory
    Dim compParam() As ComponentParam
    Dim generalComponentFolderFactory As ComponentFolderFactory
    Dim cFolder As ComponentFolder, rootCFolder As ComponentFolder
    Dim tmpList() As String
    Dim compList As List, X
    On Error GoTo NewComponentErr
    ' Get Component Folder
    ' Get a ComponentFolderFactory from the QCConnectiononnection object
    Set compFactory = QCConnection.ComponentFactory
    ' Add the component
    Set myComp = compFactory.Item(compID)
    Dim errString As String
    myComp.Field("CO_RESPONSIBLE") = tmpComp.Scripter  'Scripter
    myComp.Field("CO_DESC") = "" 'tmpComp.Component_Description   'Component Description
    myComp.Field("CO_USER_TEMPLATE_01") = tmpComp.Status  'Status
    myComp.Field("CO_USER_TEMPLATE_02") = tmpComp.Peer_Reviewer  'Peer Reviewer
    myComp.Field("CO_USER_TEMPLATE_03") = tmpComp.QA_Reviewer  'QA Reviewer
    myComp.Field("CO_USER_TEMPLATE_05") = tmpComp.Planned_End_Date  'Planned End
    myComp.Field("CO_USER_TEMPLATE_04") = tmpComp.Planned_Start_Date  'Planned Start
    myComp.Field("CO_USER_TEMPLATE_08") = "AUTOMATED"
    myComp.ScriptType = "QT-SCRIPTED"
    myComp.ApplicationAreaID = 11866
    myComp.Post
    'Return the new component.
        Set compParamFactory = myComp.ComponentParamFactory
        Set compList = compParamFactory.NewList("") ' HERE!!!
        For X = 1 To compList.Count
                compParamFactory.RemoveItem compList.Item(X).ID
        Next
    ReDim tmpList(0)
    mdiMain.pBar.Max = UBound(tmpComp.Parameters) + 1
    For j = LBound(tmpComp.Parameters) To UBound(tmpComp.Parameters)
        If tmpComp.Parameters(j) <> "" And IsParameterDeclared(tmpList, Replace(tmpComp.Parameters(j), "-", "_")) = False And Trim(tmpComp.DefaultValue(j)) <> "" Then
            ReDim Preserve compParam(j)
            Set compParam(j) = compParamFactory.AddItem(Null)
            compParam(j).IsOut = 0
            compParam(j).Name = Replace((tmpComp.Parameters(j)), "-", "_")
            compParam(j).Desc = (tmpComp.Parameters(j))
            compParam(j).Value = tmpComp.DefaultValue(j)
            compParam(j).ValueType = "String"
            compParam(j).Order = j + 1
            compParam(j).Post
            tmpList(UBound(tmpList)) = compParam(j).Name
            ReDim Preserve tmpList(UBound(tmpList) + 1)
        End If
        Debug.Print j
        mdiMain.pBar.Value = j
        If mdiMain.pBar.Max > 30 Then
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
    ReDim Preserve compParam(j)
    Set compParam(j) = compParamFactory.AddItem(Null)
    compParam(j).IsOut = 0
    compParam(j).Name = "EMPTY"
    compParam(j).Desc = "Placeholder for all EMPTY parameters"
    compParam(j).Value = ""
    compParam(j).ValueType = "String"
    compParam(j).Order = j + 1
    compParam(j).Post
    tmpList(UBound(tmpList)) = compParam(j).Name
    ReDim Preserve tmpList(UBound(tmpList) + 1)
    myComp.Post
    myComp.UnLockObject
Exit Function
NewComponentErr:
    Dim tmpFile As New clsFiles
    Debug.Print "Error in creating component in line " & tmpComp.Component_Name
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[BC CREATE: (FAILED) " & Now & " " & tmpComp.Component_Name & "] " & Err.Description
End Function
'########################### End Of Create New Component ##########################

Private Sub GenerateDescription(tmpComp As BComponent)
Dim tmp
tmp = "Summary" & vbCrLf
tmp = tmp & "This is a Consolidated component generated by SAP Test Acceleration and Optimization. " & vbCrLf
tmp = tmp & "Full Component Name : " & tmpComp.Component_Name & vbCrLf
tmp = tmp & "Date of creation: " & Format(Now, "mmm-dd-yyyy") & vbCrLf
tmp = tmp & "Time of creation: " & Format(Now, "hh:mm:ss") & vbCrLf & vbCrLf
tmp = tmp & "Contained Component Info" & vbCrLf

End Sub

Private Sub ClearForm()
Dim fileFunct As New clsFiles
BCTARGET = fileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦BCTARGET" & curDomain & "-" & curProject & "¦")
cmdUp.Enabled = False
cmdDown.Enabled = False
cmdAdd.Enabled = False
cmdRemove.Enabled = False
QCTree.Nodes.Clear
Me.QCTree.Visible = True
Control_Auto = False
End Sub

Function GetCommentText(rCommentCell As Range)
     Dim strGotIt As String
         On Error Resume Next
         strGotIt = WorksheetFunction.Clean _
             (rCommentCell.Comment.Text)
         GetCommentText = strGotIt
         On Error GoTo 0
End Function

Private Sub OutputTable()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim i, Protections
Dim curTab, dt_cnt As Integer
Dim w, cnt, curBC, curOrder, curID, tmpNode, tmpFName


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
    
    xlObject.Sheets("Sheet2").Range("A6").Value = "Filename:"
    xlObject.Sheets("Sheet2").Range("B6").Value = QCTree.Nodes(1).Text
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
  xlObject.Sheets("Sheet3").Name = "DATA SHEET"
  'xlObject.Visible = True
  curTab = "SAPTAO01"
  xlObject.Sheets("Sheet1").Name = curTab

    cnt = 2
    Set tmpNode = QCTree.Nodes("Root")
    Do While True
        If tmpNode Is Nothing Then Exit Do
        If Left(tmpNode.Key, 1) = "C" Then
            xlObject.Sheets(curTab).Range("A" & cnt) = CInt(Replace(Left(tmpNode.Text, 4), "[", ""))
            xlObject.Sheets(curTab).Range("B" & cnt) = tmpNode.Tag
            xlObject.Sheets(curTab).Range("C" & cnt) = Right(tmpNode.Text, Len(tmpNode.Text) - 6)
            curOrder = xlObject.Sheets(curTab).Range("A" & cnt)
            curID = xlObject.Sheets(curTab).Range("B" & cnt)
            curBC = xlObject.Sheets(curTab).Range("C" & cnt)
            If Not (tmpNode.Child.FirstSibling Is Nothing) Then Set tmpNode = tmpNode.Child.FirstSibling
        ElseIf Left(tmpNode.Key, 1) = "N" Then
            xlObject.Sheets(curTab).Range("A" & cnt) = curOrder
            xlObject.Sheets(curTab).Range("B" & cnt) = curID
            xlObject.Sheets(curTab).Range("C" & cnt) = curBC
            xlObject.Sheets(curTab).Range("D" & cnt) = tmpNode.Text
            If Not (tmpNode.Child.FirstSibling Is Nothing) Then Set tmpNode = tmpNode.Child.FirstSibling
        ElseIf Left(tmpNode.Key, 1) = "V" Then
            xlObject.Sheets(curTab).Range("E" & cnt) = tmpNode.Text
            xlObject.Sheets(curTab).Range("F" & cnt) = tmpNode.Tag
            cnt = cnt + 1
            If Not (tmpNode.Parent.Next Is Nothing) Then
                Set tmpNode = tmpNode.Parent.Next
            Else
                Set tmpNode = tmpNode.Parent.Parent.Next
            End If
            'MsgBox UCase(Left(xlObject.Sheets(curTab).Range("E" & cnt), 3))
            If UCase(Left(xlObject.Sheets(curTab).Range("E" & cnt - 1), 3)) = "DT_" Then
                dt_cnt = dt_cnt + 1
                xlObject.Sheets("DATA SHEET").Range("A" & dt_cnt) = xlObject.Sheets(curTab).Range("E" & cnt - 1)
            End If
        ElseIf tmpNode.FirstSibling.Index = 1 Then
            Set tmpNode = tmpNode.Child.FirstSibling
        End If
    Loop
'On Error Resume Next
    xlObject.Sheets(curTab).Select
    xlObject.Sheets(curTab).Range("A1") = "Order"
    xlObject.Sheets(curTab).Range("B1") = "Component ID"
    xlObject.Sheets(curTab).Range("C1") = "Component Name"
    xlObject.Sheets(curTab).Range("D1") = "Parameter Name"
    xlObject.Sheets(curTab).Range("E1") = "Parameter Value"
    xlObject.Sheets(curTab).Range("F1") = "Parameter Tag"
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
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & "TAO Automation Module " & curDomain & "-" & curProject & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""
    
  tmpFName = "SAPTAO01-" & CleanTheString_PARAMS(QCTree.Nodes(1).Text) & "-" & Format(Now, "mmddyyyy HHMM AMPM")
  On Error Resume Next
  xlObject.Workbooks(1).SaveAs tmpFName
  dlgOpenExcel.filename = tmpFName: dlgOpenExcel.ShowSave
  If dlgOpenExcel.filename <> "" Then xlObject.Workbooks(1).SaveAs dlgOpenExcel.filename
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate
  
  Set xlWB = Nothing
  Set xlObject = Nothing
  FXGirl.EZPlay FXExportToExcel
  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export to MS Excel completed.": Exit Sub:
OutErr:     MsgBox Err.Description, vbCritical: xlObject.Visible = True: xlObject.ActiveWindow.Activate: Set xlWB = Nothing: Set xlObject = Nothing
On Error GoTo 0
End Sub

Private Sub OutputTable_ToTestPlan()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim i, Protections
Dim curTab, dt_cnt As Integer
Dim w, cnt, curBC, curOrder, curID, tmpNode, tmpFName


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
    
    xlObject.Sheets("Sheet2").Range("A6").Value = "Filename:"
    xlObject.Sheets("Sheet2").Range("B6").Value = QCTree.Nodes(1).Text
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
  curTab = "TP_LINKBPT-01"
  xlObject.Sheets("Sheet1").Name = curTab

    cnt = 2
    Set tmpNode = QCTree.Nodes("Root")
    For Each tmpNode In QCTree.Nodes
       If Left(tmpNode.Key, 1) = "C" Then
            xlObject.Sheets(curTab).Range("C" & cnt) = QCTree.Nodes("Root").Text
            xlObject.Sheets(curTab).Range("D" & cnt) = Right(tmpNode.Text, Len(tmpNode.Text) - 6)
            cnt = cnt + 1
        End If
    Next
'On Error Resume Next
    xlObject.Sheets(curTab).Select
    xlObject.Sheets(curTab).Range("A1") = "Test Case ID"
    xlObject.Sheets(curTab).Range("B1") = "Business Component ID"
    xlObject.Sheets(curTab).Range("C1") = "Test Case Name (TEST PLAN)"
    xlObject.Sheets(curTab).Range("D1") = "Business Component Name (BUSINESS COMPONENTS)"
    xlObject.Sheets(curTab).Range("E1") = "Validation"
    xlObject.Sheets(curTab).Range("A:E").Select
        
xlObject.Sheets(curTab).Range("A:E").Select

    xlObject.Sheets(curTab).Range("A:E").Borders(xlDiagonalDown).LineStyle = xlNone
    xlObject.Sheets(curTab).Range("A:E").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlObject.Sheets(curTab).Range("A:E").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:E").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:E").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:E").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:E").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:E").Borders(xlInsideHorizontal)
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
    xlObject.Sheets(curTab).Range("A:E").Select
    xlObject.Sheets(curTab).Range("A:E").EntireColumn.AutoFit
    xlObject.Sheets(curTab).Range("A1").Select

    xlObject.Sheets(curTab).Range("A1").AddComment
    xlObject.Sheets(curTab).Range("A1").Comment.Visible = False
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & mdiMain.Caption & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""

    xlObject.Sheets(curTab).Range("A:B").Interior.ColorIndex = 3
    xlObject.Sheets(curTab).Range("E:E").Interior.ColorIndex = 3
    
  xlObject.Workbooks(1).SaveAs "TP_LINKBPT-01" & "-" & Format(Now, "mmddyyyy HHMMSS AMPM")
  dlgOpenExcel.filename = tmpFName: dlgOpenExcel.ShowSave
  If dlgOpenExcel.filename <> "" Then xlObject.Workbooks(1).SaveAs dlgOpenExcel.filename
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate
  FXGirl.EZPlay FXExportToExcel
  Set xlWB = Nothing
  Set xlObject = Nothing
  
  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export to Test Plan MS Excel completed.": Exit Sub:
OutErr:     MsgBox Err.Description, vbCritical: xlObject.Visible = True: xlObject.ActiveWindow.Activate: Set xlWB = Nothing: Set xlObject = Nothing
End Sub

Private Function Load_BC(BC_ID As String, BC_Order As Integer, BC_Name As String, tmpComp As BComponent) As String
Dim FileStruct As New clsFiles
Dim tmp, i
Dim tmpBC
Dim comp As Component
Dim compStorage As ExtendedStorage
Dim CompDownLoadPath As String
Dim compFact As ComponentFactory

If FileStruct.FolderExists(App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & BC_ID) = False Then
    Set compFact = QCConnection.ComponentFactory
    Set comp = compFact.Item(BC_ID)
    Set compStorage = comp.ExtendedStorage(0)
    compStorage.ClientPath = App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & BC_ID
    CompDownLoadPath = compStorage.Load("Action1\Script.mts,Action1\Resource.mtr", True)
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Dumping BC Script (PASSED) " & Now & " " & BC_ID
End If

    tmp = FileStruct.ReadFromFileToArray(App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & BC_ID & "\Action1\Script.mts")
    tmpBC = "ComponentName = "" " & BC_Order & " - " & BC_Name & " """ & vbCrLf
    tmpBC = tmpBC & "'ComponentPath =" & GetBusinessComponentFolderPath(BC_ID) & " [" & BC_ID & "]" & vbCrLf
    If BC_Order = 1 Then
        tmpBC = tmpBC & "CBASE_BOOTSTRAP_ONLY = True" & vbCrLf
        tmpBC = tmpBC & "ExecuteFile ""[QualityCenter] Subject\BPT Resources\Libraries\CBASE_Init.vbs.txt""" & vbCrLf
    End If
    tmpBC = tmpBC & "SetExitConsolidatedComponentOnFailure(False)" & vbCrLf & vbCrLf & vbCrLf
    For i = LBound(tmp) To UBound(tmp)
        If InStr(1, tmp(i), "ExecuteFile ""[QualityCenter]", vbTextCompare) = 0 And InStr(1, tmp(i), "CBASE_BOOTSTRAP_ONLY = True", vbTextCompare) = 0 Then
             tmp(i) = Replace(tmp(i), "ï»¿", "")
             tmp(i) = UpdateParametersInBC(CStr(tmp(i)), BC_Name, BC_Order, tmpComp)
             tmpBC = tmpBC & tmp(i) & vbCrLf
        End If
    Next
    'FileStruct.WriteNewFile "C:\" & BC_Name & ".txt", CStr(tmpBC)
    Load_BC = CStr(tmpBC)
End Function

Private Function Dump_BC(BC_ID)
Dim FileStruct As New clsFiles
Dim tmp, i
Dim tmpBC
Dim comp As Component
Dim compStorage As ExtendedStorage
Dim CompDownLoadPath As String
Dim compFact As ComponentFactory
Dim NullList As List

Dim compParamFactory As ComponentParamFactory
Dim compParam As ComponentParam
Dim tmpList As List
    Set compFact = QCConnection.ComponentFactory
    Set comp = compFact.Item(BC_ID)
    Set compParamFactory = comp.ComponentParamFactory
    Set tmpList = compParamFactory.NewList("")
    tmp = ""
    tmp = "{" & comp.ID & "}" & vbCrLf
    tmp = "<" & GetBusinessComponentFolderPath(comp.ID) & ">" & vbCrLf
    tmp = "|" & comp.Name & "|" & vbCrLf
    For i = 1 To tmpList.Count
        tmp = tmp & "[" & tmpList.Item(i).Name & "] " & tmpList.Item(i).Value & vbCrLf
    Next
    Set compStorage = comp.ExtendedStorage(0)
    compStorage.ClientPath = App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & BC_ID
    CompDownLoadPath = compStorage.Load("Action1\Script.mts", True)
    FileStruct.WriteNewFile App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & BC_ID & "\Params.txt", CStr(tmp)
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Downloaded BC Script (PASSED) " & Now & " " & BC_ID

End Function

Private Function Save_BC(BC_ID)
Dim FileStruct As New clsFiles
Dim tmp, i
Dim tmpBC
Dim comp As Component
Dim compStorage As ExtendedStorage
Dim CompDownLoadPath As String
Dim compFact As ComponentFactory
Dim NullList As List

If FileStruct.FolderExists(App.path & "\SQC DAT\BC Template") = True Then
    Set compFact = QCConnection.ComponentFactory
    Set comp = compFact.Item(BC_ID)
    Set compStorage = comp.ExtendedStorage(0)
    compStorage.ClientPath = App.path & "\SQC DAT\BC Template"
    CompDownLoadPath = compStorage.SaveEx("*.*", True, NullList)       'SAVE TO QC
    CompDownLoadPath = compStorage.SaveEx("Action0\*.*", True, NullList)       'SAVE TO QC
    CompDownLoadPath = compStorage.SaveEx("Action1\*.*", True, NullList)       'SAVE TO QC
    'Dump_BC BC_ID
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploaded BC Script (PASSED) " & Now & " " & BC_ID
End If
End Function

Private Function UpdateParametersInBC(txtStep As String, BC_Name As String, BC_Order As Integer, tmpBC As BComponent) As String
Dim tmp, tmpParam, i
If InStr(1, txtStep, "GetParameterValue", vbTextCompare) <> 0 Then
  tmpParam = (GetParameterNameInStep(txtStep, "GetParameterValue"))
  For i = LBound(tmpParam) To UBound(tmpParam)
    If Trim(GetParValue("""C" & Format(BC_Order, "000") & "_" & CleanTheString_PARAMS(tmpParam(i)) & """", tmpBC)) <> "" Then
      tmp = Replace(txtStep, """" & tmpParam(i) & """", """C" & Format(BC_Order, "000") & "_" & CleanTheString_PARAMS(tmpParam(i)) & """")
    Else
      tmp = Replace(txtStep, """" & tmpParam(i) & """", """EMPTY""")
    End If
  Next
  UpdateParametersInBC = tmp
  Exit Function
ElseIf InStr(1, txtStep, "ResolveParameter", vbTextCompare) <> 0 Then
  tmpParam = (GetParameterNameInStep(txtStep, "ResolveParameter"))
  For i = LBound(tmpParam) To UBound(tmpParam)
    If Trim(GetParValue("""C" & Format(BC_Order, "000") & "_" & CleanTheString_PARAMS(tmpParam(i)) & """", tmpBC)) <> "" Then
      tmp = Replace(txtStep, """" & tmpParam(i) & """", """C" & Format(BC_Order, "000") & "_" & CleanTheString_PARAMS(tmpParam(i)) & """")
    Else
      tmp = Replace(txtStep, """" & tmpParam(i) & """", """EMPTY""")
    End If
  Next
  UpdateParametersInBC = tmp
  Exit Function
ElseIf InStr(1, txtStep, "Parameter", vbTextCompare) <> 0 Then
  tmpParam = (GetParameterNameInStep(txtStep, "Parameter"))
  tmp = txtStep
  For i = LBound(tmpParam) To UBound(tmpParam)
    If Trim(GetParValue("""C" & Format(BC_Order, "000") & "_" & CleanTheString_PARAMS(tmpParam(i)) & """", tmpBC)) <> "" Then
      tmp = Replace(tmp, """" & tmpParam(i) & """", """C" & Format(BC_Order, "000") & "_" & CleanTheString_PARAMS(tmpParam(i)) & """")
    Else
      tmp = Replace(tmp, """" & tmpParam(i) & """", """EMPTY""")
    End If
  Next
  UpdateParametersInBC = tmp
  Exit Function
Else
  UpdateParametersInBC = txtStep
End If
End Function

Private Function GetParValue(strFind As String, tmpBC As BComponent) As String
Dim i
    For i = LBound(tmpBC.Parameters) To UBound(tmpBC.Parameters)
        If UCase(Trim(tmpBC.Parameters(i))) = Replace(UCase(Trim(strFind)), """", "") Then
            GetParValue = tmpBC.DefaultValue(i)
            Exit Function
        End If
    Next
End Function

Private Function GetParameterNameInStep(txtStep As String, ParIdentifier As String)
Dim Start, tmpParam(), stringFunct As New clsStrings, ParNum, LastParPos, i
On Error Resume Next
ParNum = stringFunct.CountWordInstance(txtStep, ParIdentifier)
LastParPos = 1
ReDim tmpParam(0)
For i = 1 To ParNum
    Start = InStr(LastParPos, txtStep, ParIdentifier, vbTextCompare)
    LastParPos = Start + Len(ParIdentifier) - 1
    tmpParam(UBound(tmpParam)) = Right(txtStep, Len(txtStep) - Start - Len(ParIdentifier) - 1)
    tmpParam(UBound(tmpParam)) = Replace(Left(tmpParam(UBound(tmpParam)), InStr(1, tmpParam(UBound(tmpParam)), """")), """", "")
    ReDim Preserve tmpParam(UBound(tmpParam) + 1)
Next
ReDim Preserve tmpParam(UBound(tmpParam) - 1)
GetParameterNameInStep = tmpParam
End Function

Private Sub QCTree_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmp
    If Button = 1 Then
        ' Changed to that the item that the mouse is over is dragged, not the currently
        ' selected item which could be else where
        Set nodx = QCTree.HitTest(X, Y) ' Set the item being dragged.
        ' Make sure that the Root node is not selected
        If Not nodx Is Nothing Then
            If nodx.Parent Is Nothing Then Set nodx = Nothing
        End If
    Else
        Set nodx = QCTree.HitTest(X, Y) ' Set the item being dragged.
        ' Make sure that the Root node is not selected
        If Not nodx Is Nothing Then
            If Left(nodx.Key, 1) = "V" Then
                tmp = InputBox("Enter a new Parameter Tag", "Parameter Tag", nodx.Tag)
                nodx.Tag = Trim(tmp)
            End If
        End If
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
    If Left(QCTree.HitTest(X, Y).Key, 1) = "V" Then
        QCTree.ToolTipText = QCTree.HitTest(X, Y).Tag & " [Right Click to Edit]"
    End If
End Sub


Private Sub QCTree_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If indrag = True Then
        ' Set DropHighlight to the mouse's coordinates.
        Set QCTree.DropHighlight = QCTree.HitTest(X, Y)
    End If
End Sub

'########################### Create New Test Plan BPT ###########################
Private Function CreateTest(tmpBPTPath_Location As String, tmpBPTTest_Suit_Case_Name As String)   'PathLocation As String, TestSuitCaseName As String, Scripter As String, _
    'PeerReviewer As String, QAReviewer As String, PlanStart As String, PlanEnd As String, Status As String)
Dim i
Dim test1 As Test
Dim NewTest As Test
Dim folder As SubjectNode
Dim testF As TestFactory
Dim treeM As TreeManager
Set treeM = QCConnection.TreeManager

    Set folder = treeM.NodeByPath(tmpBPTPath_Location & "\")
    Set testF = folder.TestFactory
    '*****CREATING THE TESTS*****
    Set NewTest = testF.AddItem(Null)
    NewTest.Name = tmpBPTTest_Suit_Case_Name
    NewTest.Type = "BUSINESS-PROCESS"
    'Put the test in the new subject folder
    NewTest.Field("TS_SUBJECT") = folder.NodeID
    '-- Enter Scripter
    NewTest.Field("TS_RESPONSIBLE") = curUser
    '-- Enter Peer Reviewer
    NewTest.Field("TS_USER_TEMPLATE_02") = curUser
    '-- Enter QA Reviewer
    NewTest.Field("TS_USER_TEMPLATE_03") = curUser
    '-- Enter Planned Scripting End Date
    NewTest.Field("TS_USER_TEMPLATE_04") = Format(Now, "mm/dd/yyyy")
    '-- Enter Planned Scripting Start Date
    NewTest.Field("TS_USER_TEMPLATE_05") = Format(Now, "mm/dd/yyyy")
    '-- Enter Status
    NewTest.Field("TS_USER_TEMPLATE_07") = "040 Ready For Test"
    NewTest.Field("TS_STATUS") = "020 Ready for Review"
    NewTest.Post
    CreateTest = NewTest.ID
Set treeM = Nothing
Set folder = Nothing
Set testF = Nothing
Set NewTest = Nothing
End Function
'########################### End Of Create New Test Plan BPT ###########################

'########################### Link New Test Plan Folder ###########################
Private Sub LinkBPTTest(tmp_BPTTest_Plan_ID As String, tmp_BPTBusiness_Component_ID As String)
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

'Get the Test Factory
Set tfact = QCConnection.TestFactory

'Linking Business Component to BPT Test
Set compFact = QCConnection.ComponentFactory
Set comp = compFact.Item(tmp_BPTBusiness_Component_ID)
Set mytest = tfact.Item(tmp_BPTTest_Plan_ID)
Set myBPTest = mytest
Set myBPComponent = myBPTest.AddBPComponent(comp)
myBPComponent.FailureCondition = "Continue"
myBPComponent.Iterations.Item(1).DeleteIterationParams
myBPComponent.Iterations.Item(1).Order = 0
myBPComponent.Post

        Set com = QCConnection.Command
        com.CommandText = "select count(*) from BPTEST_TO_COMPONENTS where bc_bpt_id=" & mytest.ID
        Set recset = com.Execute
        bpcount = recset.FieldValue(0)
       
        com.CommandText = "update BPTEST_TO_COMPONENTS set bc_order=" & bpcount & " where bc_id=" & myBPComponent.ID
        Set recset = com.Execute
        
mytest.Post
Set comp = Nothing
Set mytest = Nothing
Set myBPTest = Nothing
Set myBPComponent = Nothing
Set myTempIterationParam = Nothing
End Sub
'########################### End Of Link New Test Plan Folder ###########################

'########################### Link New Test Plan Folder ###########################
Private Sub RemoveBPTTest(tmp_BPTTest_Plan_ID As String)
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

'Get the Test Factory
Set tfact = QCConnection.TestFactory

        Set mytest = tfact.Item(tmp_BPTTest_Plan_ID)
        Set myBPTest = mytest
        myBPTest.Load
        For Each bpComp In myBPTest.BPComponents
            myBPTest.DeleteBPComponent bpComp
            myBPTest.Save
            myBPTest.Refresh
            mytest.Post
        Next
        
Set comp = Nothing
Set mytest = Nothing
Set myBPTest = Nothing
Set myBPComponent = Nothing
Set myTempIterationParam = Nothing
End Sub
'########################### End Of Link New Test Plan Folder ###########################

'########################### Promote Test Plan Parameters ###########################
Private Function PromoteParamBPTTest(tmp_BPTTest_Plan_ID As String)
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
        Set mytest = tfact.Item(tmp_BPTTest_Plan_ID)
        Set myBPTest = mytest
        myBPTest.Load
        For Each bpComp In myBPTest.BPComponents
            For Each iter In bpComp.Iterations
                X = 1
                For Each bpParam In bpComp.BPParams
                        For Each myTempIteration In bpComp.Iterations
                            Set myTempIterationParam = myTempIteration.IterationParams.Item(X)
                                    Set myRTParam = myBPTest.AddRTParam
                                    myRTParam.Name = myTempIterationParam.BPParameter.ComponentParamName
                                    myRTParam.ValueType = "String"
                                    myTempIterationParam.Value = GetParamValueByID(bpComp.Component.ID & myTempIterationParam.BPParameter.ComponentParamName)
                        Next myTempIteration
                        X = X + 1
                Next bpParam
                myBPTest.Refresh
                myBPTest.Save
                X = 0
            Next iter
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

Private Function GetParamValueByID(parCode As String) As String
Dim i
For i = LBound(tmpPars) To UBound(tmpPars)
    If tmpPars(i).ID = parCode And tmpPars(i).Used = False Then
        GetParamValueByID = tmpPars(i).Value
        tmpPars(i).Used = True
        Exit Function
    End If
Next
End Function

'########################### Promote Test Plan Parameters ###########################
Private Function PromoteParamBPTTestRunBC(Test_Plan_ID As String, BC_ID As String, ParName As String, ParPos As Integer)
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
Dim X, myRTParam As RTParam, Changed As Boolean, CurPos As Integer
        
        CurPos = 1
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
                                    If CurPos = ParPos Then
                                        If Trim(myTempIterationParam.Value) = "" Then
                                            Set myRTParam = myBPTest.AddRTParam
                                            myRTParam.Name = ParName
                                            myRTParam.ValueType = "String"
                                            myTempIterationParam.Value = "{" & myRTParam.Name & "}"
                                            FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Promoting Test Parameters [BLANK] (PASSED) " & Now & " (TEST ID:" & Test_Plan_ID & "-BC ID:" & BC_ID & ") - " & myTempIterationParam.Value
                                            Changed = True
                                        ElseIf "{" & Trim(UCase(myTempIterationParam.BPParameter.ComponentParamName)) & "}" <> Trim(UCase(myTempIterationParam.Value)) Then
                                            Set myRTParam = myBPTest.AddRTParam
                                            myRTParam.Name = Trim(ParName)
                                            myRTParam.Name = Replace(myRTParam.Name, "{", "")
                                            myRTParam.Name = Replace(myRTParam.Name, "}", "")
                                            myRTParam.Name = Replace(myRTParam.Name, " ", "")
                                            myRTParam.ValueType = "String"
                                            myTempIterationParam.Value = "{" & myRTParam.Name & "}"
                                            FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Promoting Test Parameters (PASSED) [INCONSISTENT] " & Now & " (TEST ID:" & Test_Plan_ID & "-BC ID:" & BC_ID & ") - " & myTempIterationParam.Value
                                            Changed = True
                                        End If
                                    End If
                                    CurPos = CurPos + 1
                                Else
                                     FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Promoting Test Parameters (FAILED) " & Now & " (TEST ID:" & Test_Plan_ID & "-BC ID:" & BC_ID & ") - " & myTempIterationParam.Value & " " & Err.Description
                                End If
                                Err.Clear
                                On Error GoTo 0
                            Next myTempIteration
                            X = X + 1
                    Next bpParam
                    If Changed = True Then
                        myBPTest.Refresh
                        myBPTest.Save
                    End If
                    Changed = False
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

PromoteParamBPTTestRunBC = True

Set comp = Nothing
Set mytest = Nothing
Set myBPTest = Nothing
Set myBPComponent = Nothing
Set myTempIterationParam = Nothing
Set myRTParam = Nothing
End Function
'########################### End Of Promote Test Plan Parameters ###########################
