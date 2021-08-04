VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "R2a-R3 Testing Team - SuperQC"
   ClientHeight    =   8730
   ClientLeft      =   6765
   ClientTop       =   3720
   ClientWidth     =   12030
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":08CA
   ScaleHeight     =   8730
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picUser 
      Height          =   315
      Left            =   11700
      ScaleHeight     =   255
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton optR4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Release 4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   12
      Top             =   5220
      Width           =   1275
   End
   Begin VB.OptionButton optR3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Release 3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   11
      Top             =   5220
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.CommandButton cmdAuthenticate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Authenticate"
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
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6540
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   8400
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Login"
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
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   1455
   End
   Begin VB.ComboBox cmbProject 
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
      Height          =   330
      ItemData        =   "frmLogin.frx":D1214
      Left            =   8760
      List            =   "frmLogin.frx":D121E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   7500
      Width           =   3135
   End
   Begin VB.ComboBox cmbDomain 
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
      Height          =   330
      ItemData        =   "frmLogin.frx":D1237
      Left            =   8760
      List            =   "frmLogin.frx":D123E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   7080
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   8760
      PasswordChar    =   "�"
      TabIndex        =   1
      Top             =   6120
      Width           =   3075
   End
   Begin VB.TextBox txtUserName 
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
      Left            =   8760
      TabIndex        =   0
      Top             =   5520
      Width           =   3075
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   -120
      Top             =   5040
      Width           =   12495
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   7830
      TabIndex        =   7
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domain:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   7800
      TabIndex        =   6
      Top             =   7140
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   7620
      TabIndex        =   5
      Top             =   6240
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   7440
      TabIndex        =   4
      Top             =   5640
      Width           =   1185
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cuser As String
Public cpassword As String
Public cdomain As String
Public cproject As String
Public cLogPath As String
Public cServerURL As String
Dim filename As String
Public stamp
Dim WindowFunct As New clsWindow
Dim Override As Boolean

Private Sub cmbDomain_Change()
Dim Project_name
cmbProject.Clear
cmbProject.Enabled = True
For Each Project_name In QCConnection.VisibleProjects(cmbDomain.Text)
    cmbProject.AddItem Project_name
Next
cmbProject.ListIndex = 0
End Sub

Private Sub cmbDomain_Click()
cmbDomain_Change
End Sub

Private Sub cmbDomain_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmbDomain.ListIndex <> -1 Then cmbProject.SetFocus
End If
End Sub

Private Sub cmbDomain_LostFocus()
cmbDomain_Change
End Sub

Private Sub cmbProject_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmbProject.ListIndex <> -1 Then cmdLogin_Click
End If
End Sub

Private Sub cmdAuthenticate_Click()
On Error Resume Next
    Dim domain_name
    If optR3.Value = True Then
        QCConnection.InitConnectionEx ("https://qcurl.saas.hp.com/qcbin")
        curQCInstance = "https://qcurl.saas.hp.com/qcbin"
    Else
        QCConnection.InitConnectionEx ("http://qcurl/qcbin")
        curQCInstance = "http://qcurl/qcbin"
    End If
    If txtUserName.Text = "" Then
        MsgBox "Invalid: Username is empty."
        Exit Sub
    End If
    If txtPassword.Text = "" Then
        MsgBox "Invalid: Password is empty."
        Exit Sub
    End If
    
    'Authenticate the user
    If QCConnection.Connected = True Then
        QCConnection.Login CStr(txtUserName.Text), CStr(txtPassword.Text)
        If QCConnection.LoggedIn = False Then
            MsgBox "Invalid: Failed to authenticate."
            txtPassword.Text = ""
            txtPassword.SetFocus
        Else
            cmdAuthenticate.Enabled = False
            cmbDomain.Clear
            cmbDomain.Enabled = True
            cmdLogin.Enabled = True
            txtUserName.Enabled = False
            txtPassword.Enabled = False
            For Each domain_name In QCConnection.VisibleDomains
                cmbDomain.AddItem domain_name
            Next
            cmbDomain.ListIndex = 0
        End If
    End If
End Sub

Private Sub cmdLogin_Click()
Dim FileFunct As New clsFiles
Dim tmpF
If cmbDomain.Text = "" Or cmbProject.Text = "" Then Exit Sub
On Error GoTo Exp
    'Storing the Template project connection information
    frmLogin.cuser = txtUserName.Text
    frmLogin.cpassword = txtPassword.Text
    frmLogin.cdomain = cmbDomain.Text
    frmLogin.cproject = cmbProject.Text
    
    'Connecting to the Template project
    On Error Resume Next
    QCConnection.Logout
    QCConnection.Login ADMIN_ID, ADMIN_PASS
    ProgressBar1.Value = 50
    QCConnection.CONNECT cmbDomain.Text, cmbProject.Text
    curDomain = cmbDomain.Text
    curProject = cmbProject.Text
    curUser = Trim(txtUserName.Text)
    ProgressBar1.Value = 100
    If Err.Number <> 0 Then
            MsgBox "This version of SuperQC is now disabled. Please download the latest SuperQC tools." & vbCrLf & "For more information contact Roland Ross Hadi."
            ProgressBar1.Value = 0
    Else
        'Retrieving the group name
        FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "<USERID>", curUser
        If optR4.Value = True Then
            FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "<QCINSTANCE>", "R4"
            curInstance = "Release 4"
        Else
            FileFunct.WriteKeyToFile App.path & "\SQC DAT" & "\" & "myReports01.hxh", "<QCINSTANCE>", "R3"
            curInstance = "Release 3"
        End If
        If Override = False Then
          If LatestVersionCheck = False And InStr(1, App.EXEName, "prjSuper QualityCenterExplorerUltimate", vbTextCompare) = 0 Then    ' Change to False
              Patch
          End If
        End If
        SetHeader
        Unload Me
        frmTray.Show
        WindowFunct.WndHide frmTray.hWnd
        mdiMain.Show
    End If
Exit Sub
Exp:
MsgBox Err.Description
End Sub
 
Private Sub Form_Activate()
If Trim(txtUserName.Text) <> "" Then txtPassword.SetFocus
End Sub

Private Sub Form_DblClick()
Dim tmp
On Error Resume Next
tmp = (InputBox("Enter key to override auto update", "SQC.Debug"))

If tmp = 1 Then
  Override = True
Else
  Override = False
End If

If Override = True Then
  txtUserName.ForeColor = vbRed
  txtPassword.ForeColor = vbRed
Else
  txtUserName.ForeColor = vbBlack
  txtPassword.ForeColor = vbBlack
End If
End Sub

Private Sub Form_Load()
Dim FileFunct As New clsFiles
If App.PrevInstance = True Then MsgBox "Another instance of SuperQC is running on the background." & vbCrLf & "Activate SuperQc in the taskbar.": End
txtUserName.Text = FileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "<USERID>")
If FileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "<QCINSTANCE>") = "R4" Then
    optR4.Value = True
Else
    optR3.Value = True
End If
ACCESS_ = FileFunct.ReadFromFile(App.path & "\SuperQC.ico")
Select Case Replace(UCase(Trim(ACCESS_)), vbCrLf, "")
Case "ULTIMATE"
    ACCESS_ = "Ultimate"
Case "TEAM"
    ACCESS_ = "Team"
Case "REPORTER"
    ACCESS_ = "Reporter"
Case "USER"
    ACCESS_ = "User"
Case "AUTO"
    ACCESS_ = "Auto"
Case Else
    ACCESS_ = "Team"
End Select
CurVersion = "RR-R1 Testing Team - SuperQC " & "" & " - " & "" & "-" & "" & " ver." & App.Major & "." & App.Minor & "." & App.Revision
Me.Caption = CurVersion
Clipboard.Clear
Clipboard.SetText Me.Caption
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtUserName.Text <> "" And txtPassword.Text <> "" Then cmdAuthenticate_Click
End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtUserName.Text <> "" Then txtPassword.SetFocus
End If
End Sub


Private Sub AttachedPic(d_path As String)
Dim Treemgr, subjnode, tfact, tfilter, tsname, tList, Test, AttachFact, Attachment
    Set Treemgr = QCConnection.TreeManager
    Set subjnode = Treemgr.NodeByPath("Subject\BPT Resources\SuperQC\")
    Set tfact = subjnode.TestFactory
    Set tfilter = tfact.Filter
    tsname = "_AUTO_UPDATE_SQC_"
    tsname = Replace(tsname, " ", "*")
    tsname = Replace(tsname, "(", "*")
    tsname = Replace(tsname, ")", "*")
    tfilter.Filter("TS_NAME") = tsname
    Set tList = tfact.NewList(tfilter.Text)
    Set Test = tList.Item(1)
    Set AttachFact = Test.Attachments
    Set Attachment = AttachFact.AddItem(Null)
    Attachment.filename = d_path
    Attachment.Type = 1
    Attachment.Post
End Sub

 

