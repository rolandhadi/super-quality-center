VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSourceBusinessComponents 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extract Business Component Module"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13620
   Icon            =   "frmSourceBusinessComponents.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   13620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Business Component Module"
   Begin VB.TextBox txtSPEF 
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Text            =   "AAAAAA%"
      Top             =   7740
      Width           =   11955
   End
   Begin VB.CommandButton cmdRemove 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Remove"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12060
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   1455
   End
   Begin VB.ListBox lstSelected 
      Height          =   5460
      Left            =   7800
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   840
      Width           =   5715
   End
   Begin VB.TextBox txtPlan 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   7320
      Width           =   11955
   End
   Begin VB.CheckBox chkRealTime 
      Caption         =   "Real time scripting (Dumps every component from HPQC everytime)"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   8220
      Value           =   1  'Checked
      Width           =   5055
   End
   Begin VB.TextBox txtTarget 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   6840
      Width           =   11955
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13620
      _ExtentX        =   24024
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
            Object.ToolTipText     =   "Save Configuration"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "cmdOutput"
            Object.ToolTipText     =   "Export to Excel"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdUpload"
            Object.ToolTipText     =   "Extract Business Components from HPQC"
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
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSourceBusinessComponents.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSourceBusinessComponents.frx":0B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSourceBusinessComponents.frx":0DEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView QCTree 
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   9551
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
            Picture         =   "frmSourceBusinessComponents.frx":107C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSourceBusinessComponents.frx":178E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSourceBusinessComponents.frx":1EA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSourceBusinessComponents.frx":25B2
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
            Picture         =   "frmSourceBusinessComponents.frx":2CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSourceBusinessComponents.frx":2FA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSourceBusinessComponents.frx":34F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSourceBusinessComponents.frx":3A48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   8625
      Width           =   13620
      _ExtentX        =   24024
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmSourceBusinessComponents.frx":3F99
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   23275
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "SPE Folder Code"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   7860
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "Test Plan Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   7440
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Click here to get Folder Path"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Component Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6960
      Width           =   1395
   End
End
Attribute VB_Name = "frmSourceBusinessComponents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BC_Folders
    path As String
    ID As String
End Type

Dim CheckedItems_() As BC_Folders

Private Sub Form_Load()
ClearForm
End Sub

Private Sub Label1_Click()
Dim tmpPath, tmpID: On Error Resume Next
tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "FC_ID", "FC_PATH", "COMPONENT_FOLDER") & "%"
frmLogs.txtLogs.Text = "Component ID: " & tmpID & vbCrLf & "FC_PATH: " & tmpPath & vbCrLf & "Folder Path: " & QCTree.SelectedItem.FullPath
frmLogs.Show 1
End Sub

Private Sub QCTree_DblClick()
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long, j
Dim nodx As Node

    If QCTree.SelectedItem.Children <> 0 Then Exit Sub

    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT FC_ID, FC_NAME FROM COMPONENT_FOLDER WHERE FC_FATHER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY FC_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("F" & rs.FieldValue("FC_ID")), rs.FieldValue("FC_NAME"), 1
        rs.Next
    Next
        For i = 1 To QCTree.Nodes.Count
            For j = LBound(CheckedItems_) To UBound(CheckedItems_)
                If QCTree.Nodes(i).Key = CheckedItems_(j).ID Then
                    QCTree.Nodes(i).Checked = True
                End If
            Next
        Next
End Sub

Private Sub QCTree_NodeCheck(ByVal Node As MSComctlLib.Node)
Node.Selected = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpR
Select Case Button.Key
Case "cmdRefresh"
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
    ClearForm
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
Case "cmdGenerate"
    SaveSelected
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Cofiguration Saved"
    MsgBox "SuperQC will now refresh the records and exit this window"
    Unload frmConsolidate
    Unload Me
Case "cmdUpload"
    If MsgBox("Are you sure you want to extract Business Components from this directory?", vbYesNo) = vbYes Then
        stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[DOWNLOAD START: " & Now & " " & " ]"
        DumpBusinessComponents
        stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Business Components downloaded successfully"
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[DOWNLOAD END: " & Now & " " & " ]"
    End If
End Select
End Sub

Private Sub ClearForm()
QCTree.Nodes.Clear
Dim rs As TDAPIOLELib.Recordset
Dim objCommand, FileFunct As New clsFiles
Dim i As Long, tmp, j
    QCTree.Nodes.Add , , "Root", "Components", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT FC_ID, FC_NAME FROM COMPONENT_FOLDER WHERE FC_FATHER_ID = 1 ORDER BY FC_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("FC_ID")), rs.FieldValue("FC_NAME"), 1
        rs.Next
    Next
    Me.Caption = Me.Tag
    tmp = FileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦BCFOLDER" & curDomain & "-" & curProject & "¦")
    tmp = Split(tmp, ",")
    If UBound(tmp) <> -1 Then
        ReDim CheckedItems_(UBound(tmp))
        For i = LBound(CheckedItems_) To UBound(CheckedItems_)
            CheckedItems_(i).ID = tmp(i)
        Next
        For i = 1 To QCTree.Nodes.Count
            For j = LBound(CheckedItems_) To UBound(CheckedItems_)
                If QCTree.Nodes(i).Key = CheckedItems_(j).ID Then
                    QCTree.Nodes(i).Checked = True
                End If
            Next
        Next
    Else
        ReDim CheckedItems_(0)
    End If
    txtTarget.Text = FileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦BCTARGET" & curDomain & "-" & curProject & "¦")
    BCTARGET = txtTarget.Text
    txtPlan.Text = FileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦TPTARGET" & curDomain & "-" & curProject & "¦")
    TPTARGET = txtPlan.Text
    txtSPEF.Text = FileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦SPEF" & curDomain & "-" & curProject & "¦")
    SPEF = txtSPEF.Text
    If FileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦REALTIME" & curDomain & "-" & curProject & "¦") = "01" Then
        chkRealTime.Value = 1
        REALTIME = True
    Else
        chkRealTime.Value = 0
        REALTIME = False
    End If
End Sub

Private Sub SaveSelected()
Dim tmp, z, FileFunct As New clsFiles
GetAllCheckedItems_ QCTree.Nodes(1)
UpdateAllCheckedItems_ QCTree.Nodes(1)
tmp = ""
For z = LBound(CheckedItems_) To UBound(CheckedItems_)
    If CheckedItems_(z).ID <> "" Then tmp = tmp & CheckedItems_(z).ID & ","
Next
For z = 1 To 100
    tmp = Replace(tmp, ",,", ",")
Next
If Trim(tmp) <> "" Then
    tmp = Left(tmp, Len(tmp) - 1)
    FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦BCFOLDER" & curDomain & "-" & curProject & "¦", CStr(tmp)
End If
FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦BCTARGET" & curDomain & "-" & curProject & "¦", Trim(txtTarget.Text)
FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦TPTARGET" & curDomain & "-" & curProject & "¦", Trim(txtPlan.Text)
FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦REALTIME" & curDomain & "-" & curProject & "¦", Format(Trim(chkRealTime.Value), "00")
FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦SPEF" & curDomain & "-" & curProject & "¦", Format(Trim(txtSPEF.Text), "00")
BCTARGET = Trim(txtTarget.Text)
TPTARGET = Trim(txtPlan.Text)
SPEF = Trim(txtSPEF.Text)
If chkRealTime.Value = 1 Then
    REALTIME = True
Else
    REALTIME = False
End If
CreateTestFolder TPTARGET
End Sub

Private Sub DumpBusinessComponents()
On Error Resume Next
Dim com As Command
Dim rs As Recordset
Dim rs2 As Recordset
Dim rootCFolderID As Integer
Dim rootCFolderAbsolutePath As String
Dim i As Integer, i2 As Integer
Dim k As Integer
Dim z As Integer, tmp, FileFunct As New clsFiles

Dim generalComponentFolderFactory As ComponentFolderFactory
Dim rootComponentFolderFactory As ComponentFolderFactory
Dim cFolder As ComponentFolder, rootCFolder As ComponentFolder
Dim comp As Component
Dim compStorage As ExtendedStorage
Dim CompDownLoadPath As String
Dim NullList As List
Dim isFatalErr As Boolean
Dim compFact As ComponentFactory

Dim compParamFactory As ComponentParamFactory
Dim compParam As ComponentParam
Dim tmpList As List, FileStruct As New clsFiles

Dim IssueCount As Integer


GetAllCheckedItems_ QCTree.Nodes(1)
UpdateAllCheckedItems_ QCTree.Nodes(1)

tmp = ""
For z = LBound(CheckedItems_) To UBound(CheckedItems_)
    If CheckedItems_(z).ID <> "" Then tmp = tmp & CheckedItems_(z).ID & ","
Next
For z = 1 To 100
    tmp = Replace(tmp, ",,", ",")
Next
tmp = Left(tmp, Len(tmp) - 1)
'FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦BCFOLDER" & curDomain & "-" & curProject & "¦", CStr(tmp)
'FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "¦BCTARGET" & curDomain & "-" & curProject & "¦", Trim(txtTarget.Text)
BCTARGET = Trim(txtTarget.Text)
For z = LBound(CheckedItems_) To UBound(CheckedItems_)
    If CheckedItems_(z).ID <> "" Then
        ' Get a ComponentFolderFactory from the TDConnection object
        Set generalComponentFolderFactory = QCConnection.ComponentFolderFactory
        ' Get the root folder
        Set rootCFolder = generalComponentFolderFactory.Root
        'Example of path: "Components\myCompFolder"
        'Set rootCFolder = generalComponentFolderFactory.FolderPath(CLng(Right(CheckedItems_(z).ID, Len(CheckedItems_(z).ID) - 1)))
        'rootCFolderID = rootCFolder.ID
        rootCFolderID = CLng(Right(CheckedItems_(z).ID, Len(CheckedItems_(z).ID) - 1))
        'Get the Absolute Path of the root node
        Set com = QCConnection.Command
        com.CommandText = "select fc_path from component_folder where fc_id=" & rootCFolderID
        Set rs = com.Execute
        rootCFolderAbsolutePath = rs.FieldValue(0)
        
        'Get all the folders under the root node
        com.CommandText = "select fc_id from component_folder where fc_path like '" & rootCFolderAbsolutePath & "%'"
        Set rs = com.Execute
        
        'Loop through each folder and get the components
        For i = 1 To rs.RecordCount
            com.CommandText = "select co_id, co_name from component where CO_FOLDER_ID=" & rs.FieldValue(0) & " and co_script_type <> 'MANUAL'"
            Set rs2 = com.Execute
            Set cFolder = generalComponentFolderFactory.Item(rs.FieldValue(0))
            Set compFact = cFolder.ComponentFactory
            For k = 1 To rs2.RecordCount
                Set comp = compFact.Item(rs2.FieldValue(0))
                Set compFact = QCConnection.ComponentFactory
                Set comp = compFact.Item(rs2.FieldValue(0))
                Set compParamFactory = comp.ComponentParamFactory
                Set tmpList = compParamFactory.NewList("")
                tmp = ""
                tmp = "{" & rs2.FieldValue(0) & "}" & vbCrLf
                'tmp = tmp & "~" & rs2.FieldValue(2) & "~" & vbCrLf
                tmp = tmp & "<" & GetBusinessComponentFolderPath(rs2.FieldValue(0)) & ">" & vbCrLf
                tmp = tmp & "|" & rs2.FieldValue(1) & "|" & vbCrLf
                For i2 = 1 To tmpList.Count
                    tmp = tmp & "[" & tmpList.Item(i2).Name & "] " & tmpList.Item(i2).Value & vbCrLf
                Next
                Set compStorage = comp.ExtendedStorage(0)
                'compStorage.ClientPath = App.path & "\SQC Logs\bin\" & Format(rs2.FieldValue(0), "0000000000") & "-" & rs2.FieldValue(1) & "\" & rs2.FieldValue(0)
                compStorage.ClientPath = App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & rs2.FieldValue(0)
                CompDownLoadPath = compStorage.Load("Action1\Script.mts,Action1\Resource.mtr", True)
                FileStruct.WriteNewFile App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & rs2.FieldValue(0) & "\Params.txt", CStr(tmp)
                'CompDownLoadPath = compStorage.SaveEx("\Action1\Script.mts, Action1\Resource.mtr", True, NullList) 'SAVE TO QC
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Dumping BC Script " & i & " of " & rs.RecordCount & " (" & k & "-" & rs2.RecordCount & ") - " & rs2.FieldValue(1)
                If Err.Number = 0 Then
                    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Dumping BC Script (PASSED) " & Now & " " & Format(rs2.FieldValue(0), "0000000") & "-" & rs2.FieldValue(1)
                Else
                    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Dumping BC Script (FAILED) " & Now & " " & Format(rs2.FieldValue(0), "0000000") & "-" & rs2.FieldValue(1) & " (" & Err.Description & ")"
                    IssueCount = IssueCount + 1
                End If
                Err.Clear
                rs2.Next
            Next
            rs.Next
        Next
    End If
    Set com = Nothing
    Set rs = Nothing
    Set rs2 = Nothing
    Set rootCFolder = Nothing
    Set generalComponentFolderFactory = Nothing
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Dumping BC Script " & " completed (" & IssueCount & ") issues found. see logs"
Next
stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the list..."
GetAllBusinessComponents
stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Dumping BC Script " & " completed (" & IssueCount & ") issues found. see logs"
FXGirl.EZPlay FXSQCExtractCompleted
End Sub

Private Sub GetAllCheckedItems_(objNode As Node)
Dim objSiblingNode As Node
Set objSiblingNode = objNode
Do
     If objSiblingNode.Checked = True Then
        If NotYetPresent(objSiblingNode.Key) = False Then
            CheckedItems_(UBound(CheckedItems_)).path = objSiblingNode.FullPath
            CheckedItems_(UBound(CheckedItems_)).ID = objSiblingNode.Key
            ReDim Preserve CheckedItems_(UBound(CheckedItems_) + 1)
        End If
     End If
     If Not objSiblingNode.Child Is Nothing Then
         Call GetAllCheckedItems_(objSiblingNode.Child)
     End If
     Set objSiblingNode = objSiblingNode.Next
Loop While Not objSiblingNode Is Nothing
End Sub

Private Sub UpdateAllCheckedItems_(objNode As Node)
Dim objSiblingNode As Node, i
Set objSiblingNode = objNode
Do
     If objSiblingNode.Checked = False Then
        For i = LBound(CheckedItems_) To UBound(CheckedItems_)
            If objSiblingNode.Key = CheckedItems_(i).ID Then
                CheckedItems_(i).ID = ""
                CheckedItems_(i).path = ""
            End If
        Next
     End If
     If Not objSiblingNode.Child Is Nothing Then
         Call UpdateAllCheckedItems_(objSiblingNode.Child)
     End If
     Set objSiblingNode = objSiblingNode.Next
Loop While Not objSiblingNode Is Nothing
End Sub

Private Function NotYetPresent(X As String) As Boolean
Dim i
For i = LBound(CheckedItems_) To UBound(CheckedItems_)
    If CheckedItems_(i).ID = X Then
        NotYetPresent = True
        Exit Function
    End If
Next
NotYetPresent = False
End Function

'########################### Create New Test Plan Folder ###########################
Private Function CreateTestFolder(strPath As String)  'PathLocation As String, TestSuitCaseName As String, Scripter As String, _
    'PeerReviewer As String, QAReviewer As String, PlanStart As String, PlanEnd As String, Status As String)
Dim i
Dim strFol
Dim X
Dim stru

Dim folder As SubjectNode
Dim treeM As TreeManager

    Set treeM = QCConnection.TreeManager
    '*****SPLITTING THE PATH*****
    strFol = Split(strPath, "\")
    stru = UBound(strFol)
    strPath = "Subject\"
    For X = 1 To stru
        Set folder = treeM.NodeByPath(strPath)
        On Error Resume Next
        folder.AddNode (strFol(X))
        If Err.Number = "-2147220502" Then
            On Error GoTo 0
        End If
        strPath = strPath & strFol(X) & "\"
    Next
    CreateTestFolder = True
    Set treeM = Nothing
    Set folder = Nothing
End Function
'########################### End Of Create New Test Plan Folder ###########################

