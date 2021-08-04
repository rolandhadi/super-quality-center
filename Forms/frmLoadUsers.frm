VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLoadUsers 
   Caption         =   "Upload Users Module"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   Icon            =   "frmLoadUsers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   10125
   Tag             =   "Upload Users Module"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
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
            Style           =   4
            Object.Width           =   2350
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   2650
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdOutput"
            Object.ToolTipText     =   "Export to Excel"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin VB.CheckBox chkSendEmail 
         Caption         =   "Send Email to Users"
         Height          =   315
         Left            =   6060
         TabIndex        =   6
         Top             =   120
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CommandButton cmdUpload 
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
         Left            =   2880
         Picture         =   "frmLoadUsers.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Import step description and expected results from an excel file"
         Top             =   60
         Width           =   2505
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
         Left            =   540
         Picture         =   "frmLoadUsers.frx":30EE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Import step description and expected results from an excel file"
         Top             =   60
         Width           =   2205
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
            Picture         =   "frmLoadUsers.frx":3894
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadUsers.frx":3B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadUsers.frx":3DB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   4995
      Left            =   60
      TabIndex        =   1
      Top             =   960
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   8811
      _Version        =   393216
      Cols            =   11
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
            Picture         =   "frmLoadUsers.frx":4046
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadUsers.frx":4758
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadUsers.frx":4E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadUsers.frx":557C
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
      TabIndex        =   4
      Top             =   6000
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
            Picture         =   "frmLoadUsers.frx":5C8E
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
            Picture         =   "frmLoadUsers.frx":61DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadUsers.frx":64C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadUsers.frx":6A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadUsers.frx":6F63
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "*Double click cells to update manually"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2835
   End
End
Attribute VB_Name = "frmLoadUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Type US_BPT
    UserID As String
    FullName As String
    Password As String
    Email As String
    Phone As String
    Group1 As String
    Group2 As String
    Project_1 As String
    Project_2 As String
    Description As String
    Log As String
End Type

Private All_US() As US_BPT
Private HasIssue As Boolean
Private HasUploadIssue  As Integer

Private custom As Customization
Private Users As CustomizationUsers
Private UGroups As CustomizationUsersGroups
Private UGrp As CustomizationUsersGroup
Private CustUsersGroups As CustomizationUsersGroups
Private CustGroup As CustomizationUsersGroup
Private user As CustomizationUser

Private Function LoadToArray()
Dim lastrow, i, EndArr
lastrow = flxImport.Rows - 1
ReDim All_US(0)
EndArr = -1
For i = 1 To lastrow
    If Trim(flxImport.TextMatrix(i, 0)) = "" Or Trim(flxImport.TextMatrix(i, 1)) = "" Then
        All_US(EndArr).Log = All_US(EndArr).Log & vbCrLf & "Line " & i & " is blank"
    Else
        EndArr = EndArr + 1
        ReDim Preserve All_US(EndArr)
        All_US(EndArr).UserID = LCase(flxImport.TextMatrix(i, 0))
        All_US(EndArr).FullName = flxImport.TextMatrix(i, 1)
        All_US(EndArr).Password = flxImport.TextMatrix(i, 2)
        All_US(EndArr).Email = flxImport.TextMatrix(i, 3)
        All_US(EndArr).Phone = flxImport.TextMatrix(i, 4)
        All_US(EndArr).Group1 = flxImport.TextMatrix(i, 5)
        All_US(EndArr).Group2 = flxImport.TextMatrix(i, 6)
        All_US(EndArr).Project_1 = flxImport.TextMatrix(i, 7)
        All_US(EndArr).Project_2 = flxImport.TextMatrix(i, 8)
        All_US(EndArr).Description = flxImport.TextMatrix(i, 9)
        If UCase(Trim(All_US(EndArr).Group1)) = "TDADMIN" Then All_US(EndArr).Group1 = "company_TesterFixer"
        If UCase(Trim(All_US(EndArr).Group2)) = "TDADMIN" Then All_US(EndArr).Group2 = "company_TesterFixer"
        If Trim(All_US(EndArr).Group2) = "" Then All_US(EndArr).Group2 = All_US(EndArr).Group1
        If Trim(All_US(EndArr).Project_2) = "" Then All_US(EndArr).Project_2 = All_US(EndArr).Project_1
    End If
Next
End Function

Function LoadToQC()
Dim i, j, TimeStart
Dim tmpComp, CreateAction
stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = ""
ReDim All_Folders(0)
TimeStart = Now
mdiMain.pBar.Max = UBound(All_US) + 1
For i = LBound(All_US) To UBound(All_US)
    On Error Resume Next
    CreateAction = Create_New_User(All_US(i))
    If Err.Number = -2147220181 Then Err.Clear
    If CreateAction = "Created" Then
        If Err.Number = 0 Or Err.Number = -2147220183 Then SendNotification All_US(i)
    ElseIf CreateAction = "Updated" Then
        If Err.Number = 0 Or Err.Number = -2147220183 Then SendUpdateNotification All_US(i)
    End If
    If Err.Number <> 0 And Err.Number <> -2147220183 Then
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[CREATE USER: (FAILED) " & Now & " " & All_US(i).UserID & "-" & All_US(i).FullName & "] " & Err.Description
        HasUploadIssue = HasUploadIssue + 1
    Else
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[CREATE USER: (PASSED) " & Now & " " & All_US(i).UserID & "-" & All_US(i).FullName & "]"
    End If
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Loading Users " & i + 1 & " out of " & UBound(All_US) + 1 & " (" & All_US(i).UserID & ")"
    Err.Clear
    On Error GoTo 0
    mdiMain.pBar.Value = i + 1
        If mdiMain.pBar.Max > 10 Then
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
    QCConnection.Logout
    QCConnection.Login ADMIN_ID, ADMIN_PASS
    QCConnection.CONNECT curDomain, curProject
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = UBound(All_US) + 1 & " Users(s) loaded successfully. Email was sent to the user(s). (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log (Start: " & TimeStart & ") (End: " & Now & ")"
    If HasUploadIssue <> 0 Then
      Dim tmpFile As New clsFiles
      frmLogs.Caption = App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log"
      frmLogs.txtLogs.Text = tmpFile.ReadFromFile_FAILED(App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log")
      frmLogs.Show 1
    End If
End Function

'Private Function Create_New_User(tmpUser As US_BPT) As String
'Dim lastrowofsheet As Integer
'Dim custom As Customization
'Dim Users As CustomizationUsers
'Dim UGroups As CustomizationUsersGroups
'Dim UGrp As CustomizationUsersGroup
'Dim CustUsersGroups As CustomizationUsersGroups
'Dim CustGroup As CustomizationUsersGroup
'Dim user As CustomizationUser
'Dim UserRemoved As Boolean
'
'    Pause 1
'    ' ***** ID must be TDAmin *****
'    QCConnection.Logout
'    Pause 1
'    QCConnection.Login ADMIN_ID, ADMIN_PASS
'    QCConnection.Connect curDomain, tmpUser.Project_1
'    Pause 1
'    Set custom = QCConnection.Customization
'    Set Users = custom.Users
'    Set UGroups = custom.UsersGroups
'    custom.Load
'
''    'Delete the user
''    If chkDelete.Value = Checked Then
''        On Error Resume Next
''        If Users.UserExistsInSite(tmpUser.UserID) = True Then
''            Set user = Users.user(tmpUser.UserID)
''            user.RemoveFromGroup tmpUser.Group1
''            If tmpUser.Group1 <> tmpUser.Group2 Then user.RemoveFromGroup tmpUser.Group2
''            Users.RemoveUser tmpUser.UserID
''            custom.Commit
''            UserRemoved = True
''        End If
''        On Error GoTo 0
''    End If
''
'    If Users.UserExistsInSite(tmpUser.UserID) = True Then 'And UserRemoved = False Then
'
'       '***** ADD NEW ROLES *****
'        custom.Load
'        Set user = Users.user(tmpUser.UserID)
'        If user.InGroup(tmpUser.Group1) = False Then
'            user.AddToGroup (tmpUser.Group1)
'        End If
'        '***** ADD 2ND ROLE *****
'        If tmpUser.Group2 = "" Or tmpUser.Group2 = tmpUser.Group1 Then
'        Else
'            If user.InGroup(tmpUser.Group2) = False Then
'                user.AddToGroup (tmpUser.Group2)
'            End If
'        End If
'        custom.Commit
'
'        If tmpUser.Project_1 <> "" Then
'            Pause 1
'            QCConnection.Logout
'            Pause 1
'            ' ***** ID must be TDAmin *****
'            QCConnection.Login ADMIN_ID, ADMIN_PASS
'            QCConnection.Connect curDomain, tmpUser.Project_2
'            Pause 1
'            Set custom = QCConnection.Customization
'            Set Users = custom.Users
'            Set UGroups = custom.UsersGroups
'
'           '***** ADD EXISTING USER TO CURRENT PROJECT *****
'            On Error Resume Next
'            Users.AddUser (tmpUser.UserID)
'            If Err.Number = "-2147220183" Then
'            Else
'                custom.Commit
'            End If
'            On Error GoTo 0
'
'           '***** ADD NEW ROLE *****
'            custom.Load
'            Set user = Users.user(tmpUser.UserID)
'            '***** ADD NEW ROLES *****
'             Set user = Users.user(tmpUser.UserID)
'             If user.InGroup(tmpUser.Group1) = False Then
'                 user.AddToGroup (tmpUser.Group1)
'             End If
'             '***** ADD 2ND ROLE *****
'             If tmpUser.Group2 = "" Or tmpUser.Group2 = tmpUser.Group1 Then
'             Else
'                 If user.InGroup(tmpUser.Group2) = False Then
'                     user.AddToGroup (tmpUser.Group2)
'                 End If
'             End If
'             custom.Commit
'        End If
'
'        If tmpUser.Project_2 <> "" And tmpUser.Project_2 <> tmpUser.Project_1 Then
'            Pause 1
'            QCConnection.Logout
'            Pause 1
'            ' ***** ID must be TDAmin *****
'            QCConnection.Login ADMIN_ID, ADMIN_PASS
'            QCConnection.Connect curDomain, tmpUser.Project_2
'            Pause 1
'            Set custom = QCConnection.Customization
'            Set Users = custom.Users
'            Set UGroups = custom.UsersGroups
'
'           '***** ADD EXISTING USER TO CURRENT PROJECT *****
'            On Error Resume Next
'            Users.AddUser (tmpUser.UserID)
'            If Err.Number = "-2147220183" Then
'            Else
'                custom.Commit
'            End If
'            On Error GoTo 0
'
'           '***** ADD NEW ROLE *****
'            custom.Load
'            Set user = Users.user(tmpUser.UserID)
'            '***** ADD NEW ROLES *****
'             Set user = Users.user(tmpUser.UserID)
'             If user.InGroup(tmpUser.Group1) = False Then
'                 user.AddToGroup (tmpUser.Group1)
'             End If
'             '***** ADD 2ND ROLE *****
'             If tmpUser.Group2 = "" Or tmpUser.Group2 = tmpUser.Group1 Then
'             Else
'                 If user.InGroup(tmpUser.Group2) = False Then
'                     user.AddToGroup (tmpUser.Group2)
'                 End If
'             End If
'             custom.Commit
'        End If
'
'        Create_New_User = "Updated"
'    Else
'        '***** USER CREATION *****
'        custom.Load
'        Set UGrp = UGroups.Group(tmpUser.Group1)
'        Users.AddSiteUser tmpUser.UserID, _
'            tmpUser.FullName, tmpUser.Email, _
'            tmpUser.Description, _
'             tmpUser.Phone, UGrp
'        custom.Commit
'
'        If tmpUser.Group1 <> "" Then
'            custom.Load
'            custom.Users.AddUser tmpUser.UserID
'            Set CustUsersGroups = custom.UsersGroups
'            Set CustGroup = CustUsersGroups.Group(tmpUser.Group1)
'            CustGroup.AddUser tmpUser.UserID
'            custom.Commit
'        End If
'        If tmpUser.Group2 <> "" And tmpUser.Group2 <> tmpUser.Group1 Then
'            custom.Load
'            custom.Users.AddUser tmpUser.UserID
'            Set CustUsersGroups = custom.UsersGroups
'            Set CustGroup = CustUsersGroups.Group(tmpUser.Group2)
'            CustGroup.AddUser tmpUser.UserID
'            custom.Commit
'        End If
'
'        '***** CREATE PASSWORD *****
'        Pause 1
'        QCConnection.Logout
'        Pause 1
'        QCConnection.Login tmpUser.UserID, ""
'        Pause 1
'        QCConnection.Connect curDomain, tmpUser.Project_1
'        Pause 1
'        Set custom = QCConnection.Customization
'        Set Users = custom.Users
'        Set UGroups = custom.UsersGroups
'
'        Set user = Users.user(tmpUser.UserID)
'        '***** MUST USE RANDOM PASSWORD *****
'        user.Password = tmpUser.Password
'        custom.Commit
'
'        If tmpUser.Project_1 <> "" Then
'            QCConnection.Logout
'            Pause 1
'            ' ***** ID must be TDAmin *****
'            QCConnection.Login ADMIN_ID, ADMIN_PASS
'            QCConnection.Connect curDomain, tmpUser.Project_1
'            Pause 1
'            Set custom = QCConnection.Customization
'            Set Users = custom.Users
'            Set UGroups = custom.UsersGroups
'
'           '***** ADD EXISTING USER TO CURRENT PROJECT *****
'            On Error Resume Next
'            Users.AddUser (tmpUser.UserID)
'            If Err.Number = "-2147220183" Then
'            Else
'                custom.Commit
'            End If
'            On Error GoTo 0
'
'           '***** ADD NEW ROLE *****
'            custom.Load
'            Set user = Users.user(tmpUser.UserID)
'            '***** ADD NEW ROLES *****
'             Set user = Users.user(tmpUser.UserID)
'             If user.InGroup(tmpUser.Group1) = False Then
'                 user.AddToGroup (tmpUser.Group1)
'             End If
'             '***** ADD 2ND ROLE *****
'             If tmpUser.Group2 = "" Or tmpUser.Group2 = tmpUser.Group1 Then
'             Else
'                 If user.InGroup(tmpUser.Group2) = False Then
'                     user.AddToGroup (tmpUser.Group2)
'                 End If
'             End If
'             custom.Commit
'        End If
'
'        If tmpUser.Project_2 <> "" And tmpUser.Project_2 <> tmpUser.Project_1 Then
'            QCConnection.Logout
'            Pause 1
'            ' ***** ID must be TDAmin *****
'            QCConnection.Login ADMIN_ID, ADMIN_PASS
'            QCConnection.Connect curDomain, tmpUser.Project_2
'            Pause 1
'            Set custom = QCConnection.Customization
'            Set Users = custom.Users
'            Set UGroups = custom.UsersGroups
'
'           '***** ADD EXISTING USER TO CURRENT PROJECT *****
'            On Error Resume Next
'            Users.AddUser (tmpUser.UserID)
'            If Err.Number = "-2147220183" Then
'            Else
'                custom.Commit
'            End If
'            On Error GoTo 0
'
'           '***** ADD NEW ROLE *****
'            custom.Load
'            Set user = Users.user(tmpUser.UserID)
'            '***** ADD NEW ROLES *****
'             Set user = Users.user(tmpUser.UserID)
'             If user.InGroup(tmpUser.Group1) = False Then
'                 user.AddToGroup (tmpUser.Group1)
'             End If
'             '***** ADD 2ND ROLE *****
'             If tmpUser.Group2 = "" Or tmpUser.Group2 = tmpUser.Group1 Then
'             Else
'                 If user.InGroup(tmpUser.Group2) = False Then
'                     user.AddToGroup (tmpUser.Group2)
'                 End If
'             End If
'             custom.Commit
'        End If
'        Create_New_User = "Created"
'    End If
'End Function

Private Function Create_New_User(tmpUser As US_BPT) As String
Dim lastrowofsheet As Integer
Dim UserRemoved As Boolean
    ' ***** ID must be TDAmin *****
    QCConnection.Logout
    QCConnection.Login ADMIN_ID, ADMIN_PASS
    QCConnection.CONNECT curDomain, tmpUser.Project_1
    Pause 1
    Set custom = QCConnection.Customization
    Set Users = custom.Users
    Set UGroups = custom.UsersGroups
    custom.Load
    Pause 1
    If Users.UserExistsInSite(tmpUser.UserID) = True Then
        AddProjectUser tmpUser
        Pause 1
        AddRoles tmpUser
        Pause 1
        If tmpUser.Project_2 <> "" And (Trim(tmpUser.Project_1) <> Trim(tmpUser.Project_2)) And Trim(tmpUser.Group2) <> "" Then
            QCConnection.Logout
            QCConnection.Login ADMIN_ID, ADMIN_PASS
            QCConnection.CONNECT curDomain, tmpUser.Project_2
            Set custom = QCConnection.Customization
            Set Users = custom.Users
            AddProjectUser tmpUser
            AddRoles tmpUser
        End If
        Create_New_User = "Updated"
    Else
        UserCreation tmpUser
        Pause 1
        AddProjectUser tmpUser
        CreatePassword tmpUser
        Pause 1
        If tmpUser.Project_2 <> "" And (Trim(tmpUser.Project_1) <> Trim(tmpUser.Project_2)) And Trim(tmpUser.Group2) <> "" Then
            QCConnection.Logout
            QCConnection.Login ADMIN_ID, ADMIN_PASS
            QCConnection.CONNECT curDomain, tmpUser.Project_2
            Set custom = QCConnection.Customization
            Set Users = custom.Users
            AddProjectUser tmpUser
            AddRoles tmpUser
        End If
        Create_New_User = "Created"
    End If
    Pause 1
    QCConnection.Logout
    QCConnection.Login ADMIN_ID, ADMIN_PASS
    QCConnection.CONNECT curDomain, curProject
End Function

Private Sub UserCreation(tmpUser As US_BPT)
Set UGrp = UGroups.Group(tmpUser.Group1)
Users.AddSiteUser tmpUser.UserID, _
    tmpUser.FullName, tmpUser.Email, _
    tmpUser.Description, _
    "", UGrp
custom.Commit
Set UGrp = Nothing
Set custom = Nothing
End Sub
Private Sub CreatePassword(tmpUser As US_BPT)
QCConnection.Logout
QCConnection.Login tmpUser.UserID, ""
QCConnection.CONNECT curDomain, tmpUser.Project_1
Set custom = QCConnection.Customization
Set Users = custom.Users
Set UGroups = custom.UsersGroups
Set user = Users.user(tmpUser.UserID)
'***** MUST USE RANDOM PASSWORD *****
user.Password = tmpUser.Password
custom.Commit
QCConnection.Logout
QCConnection.Login ADMIN_ID, ADMIN_PASS
End Sub

Private Sub AddProjectUser(tmpUser As US_BPT)
Set custom = QCConnection.Customization
custom.Load
On Error Resume Next
custom.Users.AddUser tmpUser.UserID
If Err.Number = "-2147220183" Then
    Exit Sub
Else
    Set CustUsersGroups = custom.UsersGroups
    Set CustGroup = CustUsersGroups.Group(tmpUser.Group1)
    CustGroup.AddUser tmpUser.UserID
End If
custom.Commit

Set custom = Nothing
Set CustUsersGroups = Nothing
Set CustGroup = Nothing
End Sub
Private Sub AddRoles(tmpUser As US_BPT)
'***** ADD NEW ROLES *****
Set custom = QCConnection.Customization
Set Users = custom.Users
custom.Load
Set user = Users.user(tmpUser.UserID)
If user.InGroup(tmpUser.Group1) = False Then
    user.AddToGroup (tmpUser.Group1)
End If
'***** ADD 2ND ROLE *****
If tmpUser.Group2 = "" Then
Else
    If user.InGroup(tmpUser.Group2) = False Then
        user.AddToGroup (tmpUser.Group2)
    End If
End If
custom.Commit
Set custom = Nothing
Set user = Nothing
End Sub

Private Sub SendNotification(tmpUser As US_BPT)
Dim tmp
tmp = "Your HPQC Account is now created. <br><br>"
tmp = tmp & "User ID: <b>" & tmpUser.UserID & "</b><br>"
tmp = tmp & "Password: <b>" & tmpUser.Password & "</b><br>"
If Trim(tmpUser.Description) <> "" Then
  tmp = tmp & "Assigned Group: <b>" & tmpUser.Description & "</b><br>"
End If
tmp = tmp & "User Role: <b>" & tmpUser.Group1 & "</b><br>"
tmp = tmp & "HPQC Link: <b>" & "http://qctesting.companyemail.net/qcbin/" & "</b><br><br>"
tmp = tmp & "To change your password, open <a href=""http://qctesting.companyemail.net/qcbin/"">HPQC</a> and access this link --> <b>https://eroom2.companyemail.com/eRoomReq/Files/Facility38/1CompanyWorld26-TestingitemsfromWorld8/0_b2e2b/PasswordReset.pdf</b> for guidelines on how to change password." & "<br><br>QC Support Contact: <b>1CompanyTesting-R4@companyemail.com</b><br>QC Support Hotline: <b>+65 8599 8076</b><br><br>Quick Documentation: <b>https://eroom2.companyemail.com/eRoom/Facility38/1CompanyWorld26-TestingitemsfromWorld8/0_ac7a8</b><br>" & "<br>" & "<b>***This is a HPQC automatic email notification. Do not reply.***</b>"
'tmp = "Your HPQC Account is now created. Your User ID <b>" & tmpUser.UserID & "</b> has now <b>" & tmpUser.Group1 & "</b> role(s)" & "</b> with the default password of <b>" & tmpUser.Password & "</b><br><br>" & "To change your password follow this link --> <b>http://qctesting.companyemail.net/qcbin/start_a.htm</b> and access this link --> <b>https://eroom2.companyemail.com/eRoomReq/Files/Facility38/1CompanyWorld26-TestingitemsfromWorld8/0_b2e2b/PasswordReset.pdf</b> for guidelines on how to change password." & "<br><br>QC Support Contact: <b>1CompanyTesting-R4@companyemail.com</b><br>QC Support Hotline: <b>+65 8599 8076</b><br><br>Quick Documentation: <b>https://eroom2.companyemail.com/eRoom/Facility38/1CompanyWorld26-TestingitemsfromWorld8/0_ac7a8</b><br>" & "<br>" & "<b>***This is a HPQC automatic email notification. Do not reply.***</b>"
If chkSendEmail.Value = Checked Then
  QCConnection.SendMail tmpUser.Email, "", "[HPQC ACCOUNTS] Your user account was successfully created", tmp, "", "HTML"
End If
QCConnection.SendMail "1CompanyTesting-R4@companyemail.com", "", "[HPQC ACCOUNTS] User account (" & tmpUser.UserID & ") was successfully created by " & curUser & " in " & curDomain & "-" & curProject, tmp, "", "HTML"
End Sub

Private Sub SendUpdateNotification(tmpUser As US_BPT)
Dim tmp
tmp = "Your HPQC Account is now updated. <br><br>"
tmp = tmp & "User ID: <b>" & tmpUser.UserID & "</b><br>"
tmp = tmp & "User Role: <b>" & tmpUser.Group1 & "</b><br>"
tmp = tmp & "HPQC Link: <b>" & "http://qctesting.companyemail.net/qcbin/" & "</b><br><br>"
tmp = tmp & "To change your password, open <a href=""http://qctesting.companyemail.net/qcbin/"">HPQC</a> and access this link --> <b>https://eroom2.companyemail.com/eRoomReq/Files/Facility38/1CompanyWorld26-TestingitemsfromWorld8/0_b2e2b/PasswordReset.pdf</b> for guidelines on how to change password." & "<br><br>QC Support Contact: <b>1CompanyTesting-R4@companyemail.com</b><br>QC Support Hotline: <b>+65 8599 8076</b><br><br>Quick Documentation: <b>https://eroom2.companyemail.com/eRoom/Facility38/1CompanyWorld26-TestingitemsfromWorld8/0_ac7a8</b><br>" & "<br>" & "<b>***This is a HPQC automatic email notification. Do not reply.***</b>"
'tmp = "Your HPQC Account is now updated. Your User ID <b>" & tmpUser.UserID & "</b> has now <b>" & tmpUser.Group1 & "</b> role(s)" & "</b><br><br>" & "To change your password follow this link --> <b>http://qctesting.companyemail.net/qcbin/start_a.htm</b> and access this link --> <b>https://eroom2.companyemail.com/eRoomReq/Files/Facility38/1CompanyWorld26-TestingitemsfromWorld8/0_b2e2b/PasswordReset.pdf</b> for guidelines on how to change password." & "<br><br>QC Support Contact: <b>1CompanyTesting-R4@companyemail.com</b><br>QC Support Hotline: <b>+65 8599 8076</b><br><br>Quick Documentation: <b>https://eroom2.companyemail.com/eRoom/Facility38/1CompanyWorld26-TestingitemsfromWorld8/0_ac7a8</b><br>" & "<br>" & "<b>***This is a HPQC automatic email notification. Do not reply.***</b>"
If chkSendEmail.Value = Checked Then
  QCConnection.SendMail tmpUser.Email, "", "[HPQC ACCOUNTS] Your user account was successfully created", tmp, "", "HTML"
End If
QCConnection.SendMail "1CompanyTesting-R4@companyemail.com", "", "[HPQC ACCOUNTS] User account (" & tmpUser.UserID & ") was successfully updated by " & curUser & " in " & curDomain & "-" & curProject, tmp, "", "HTML"
End Sub

Sub Start()
Debug.Print "New Session: " & Now
LoadToArray
LoadToQC
Debug.Print "New Finished: " & Now
End Sub

Private Function CleanHTML_BC(strText As String) As String
        Dim tmp, i
        tmp = Replace(tmp, "<html><body>", "", 1, , vbTextCompare)
        tmp = Replace(tmp, "</body></html>", "", 1, , vbTextCompare)
        tmp = Replace(strText, "&", "&amp;", 1, , vbTextCompare)
        tmp = Replace(tmp, "'", "''", 1, , vbTextCompare)
        tmp = Replace(tmp, "<", "&lt;", 1, , vbTextCompare)
        tmp = Replace(tmp, ">", "&gt;", 1, , vbTextCompare)
        tmp = Replace(tmp, """", "&quot;", 1, , vbTextCompare)
        For i = 1 To 100
            tmp = Replace(tmp, "<br>", vbCrLf, 1, , vbTextCompare)
        Next
        For i = 1 To 100
            tmp = Replace(tmp, vbCrLf, "<br>", 1, , vbTextCompare)
            tmp = Replace(tmp, vbNewLine, "<br>", 1, , vbTextCompare)
            tmp = Replace(tmp, Chr(10) & Chr(13), "<br>", 1, , vbTextCompare)
            tmp = Replace(tmp, Chr(13), "<br>", 1, , vbTextCompare)
            tmp = Replace(tmp, vbCr, "<br>", 1, , vbTextCompare)
            tmp = Replace(tmp, vbLf, "<br>", 1, , vbTextCompare)
        Next
        CleanHTML_BC = tmp
End Function

Private Sub chkSendEmail_Click()
If chkSendEmail.Value = Unchecked Then
    If MsgBox("Are you sure you want to disable Email Notifications?", vbYesNo) = vbYes Then
        chkSendEmail.Value = Unchecked
    Else
        chkSendEmail.Value = Checked
    End If
End If
End Sub

Private Sub cmdLoadExcel_Click()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim fname As String
Dim lastrow
Dim i, j, tmpParam
Dim tmpSts
Dim strFunct As New clsFiles
Dim stringFunct As New clsStrings
Dim intFunct As New clsInternet

Dim AllEmails, cnt As Integer

HasIssue = False

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

    With xlObject.ActiveWorkbook.ActiveSheet
         If UCase(Trim(.Range("A1").Value)) <> UCase(Trim("User ID")) Then
            MsgBox "Import file is invalid. Please use only sheets generated by the SuperQC"
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
         lastrow = .Range("A" & .Rows.Count).End(xlUp).row
        '.Range("A3:M" & LastRow).Copy 'Set selection to Copy
        
        ClearTable
        flxImport.Redraw = False     'Dont draw until the end, so we avoid that flash
        flxImport.row = 0            'Paste from first cell
        flxImport.col = 0
        flxImport.Rows = lastrow
        flxImport.Cols = 11
        flxImport.Redraw = False
        
        'A - Load HPQC Folder Path
        'Should not be blank
        mdiMain.pBar.Max = lastrow + 2
        For i = 2 To lastrow
            
            
            flxImport.TextMatrix(i - 1, 0) = CleanTheString_PARAMS((Trim((.Range("A" & i).Value))))        'Change number and letter
            If Trim(.Range("A" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[User ID=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 1) = Trim((.Range("B" & i).Value))        'Change number and letter
            If Trim(.Range("B" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Full Name=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 2) = "welcome_" & LCase(flxImport.TextMatrix(i - 1, 0)) & "_" & Format(stringFunct.RandomNumber(10, 1), "00")
            'flxImport.TextMatrix(i - 1, 2) = stringFunct.Scramble((Trim((.Range("A" & i).Value))), 1) & stringFunct.RandomNumber(9999, 1000)         'Change number and letter
            
            flxImport.TextMatrix(i - 1, 3) = Trim((.Range("D" & i).Value))        'Change number and letter
            If Trim(.Range("D" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Email=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            AllEmails = Split(Trim(.Range("D" & i).Value), ",")
            For cnt = LBound(AllEmails) To UBound(AllEmails)
              If intFunct.ValidateEmail(AllEmails(cnt)) = False Then
                  flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Email=Invalid Format]"
                  tmpSts = tmpSts + 1
              End If
            Next
            
            flxImport.TextMatrix(i - 1, 4) = Trim((.Range("E" & i).Value))        'Change number and letter

            flxImport.TextMatrix(i - 1, 5) = Trim((.Range("F" & i).Value))        'Change number and letter
            If Trim(.Range("F" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Role 1=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 6) = Trim((.Range("G" & i).Value))        'Change number and letter

            flxImport.TextMatrix(i - 1, 7) = Trim((.Range("H" & i).Value))        'Change number and letter
            If Trim(.Range("H" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Project 1=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 8) = Trim((.Range("I" & i).Value))        'Change number and letter
            
            flxImport.TextMatrix(i - 1, 9) = Trim((.Range("J" & i).Value))        'Change number and letter
                        
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = i - 1 & " out of " & lastrow - 1 & " validated " & Format(i / lastrow, "0.0%") & " (" & tmpSts & ") errors found."
            mdiMain.pBar.Value = i
        Next
    End With
    mdiMain.pBar.Value = mdiMain.pBar.Max
    flxImport.Redraw = True
    If tmpSts > 0 Then HasIssue = True
    xlObject.DisplayAlerts = False 'To avoid "Save woorkbook" messagebox
    
    'Close Excel
    xlWB.Close
    xlObject.Application.Quit
    Set xlWB = Nothing
    Set xlObject = Nothing
Exit Sub
ErrLoad:
MsgBox "There was an error while importing the file. Please refresh and close all excel and try again" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub cmdUpload_Click()
Dim tmpR
If Trim(flxImport.TextMatrix(1, 1)) <> "" Then
    If IncorrectHeaderDetails = False And CheckQuickUpload = True Then
        If MsgBox("Are you sure you want to upload this to HPQC?", vbYesNo) = vbYes Then
            HasUploadIssue = 0
            If HasIssue = True Then
                If MsgBox("There are some issues found in the upload sheet. Do you want to proceed?", vbYesNo) = vbYes Then
                    Randomize: tmpR = CInt(Rnd(1000) * 10000)
                    If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
                        Start
                    Else
                        MsgBox "Invalid pass key", vbCritical
                    End If
                End If
            Else
                Randomize: tmpR = CInt(Rnd(1000) * 10000)
                If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
                    Start
                Else
                    MsgBox "Invalid pass key", vbCritical
                End If
            End If
        End If
    Else
        MsgBox "The template has an invalid/incorrect headers or invalid data"
    End If
Else
    MsgBox "No items to be uploaded."
End If
End Sub

Private Function CheckQuickUpload() As Boolean
Dim intFunct As New clsInternet
  With flxImport
    .TextMatrix(1, 10) = ""
    If Trim(.TextMatrix(1, 0)) = "" Then
      .TextMatrix(1, 10) = .TextMatrix(1, 10) & "[BLANK USER ID]"
      CheckQuickUpload = False
      Exit Function
    End If
    If Trim(.TextMatrix(1, 1)) = "" Then
      .TextMatrix(1, 10) = .TextMatrix(1, 10) & "[BLANK FULL NAME]"
      CheckQuickUpload = False
      Exit Function
    End If
    If Trim(.TextMatrix(1, 3)) = "" Then
      .TextMatrix(1, 10) = .TextMatrix(1, 10) & "[BLANK EMAIL]"
      CheckQuickUpload = False
      Exit Function
    End If
    If intFunct.ValidateEmail(Trim(.TextMatrix(1, 3))) = False Then
      .TextMatrix(1, 10) = .TextMatrix(1, 10) & "[INVALID EMAIL]"
      CheckQuickUpload = False
      Exit Function
    End If
    If Trim(.TextMatrix(1, 5)) = "" Then
      .TextMatrix(1, 10) = .TextMatrix(1, 10) & "[BLANK ROLE 1]"
      CheckQuickUpload = False
      Exit Function
    End If
    If Trim(.TextMatrix(1, 7)) = "" Then
      .TextMatrix(1, 10) = .TextMatrix(1, 10) & "[BLANK PROJECT 1]"
      CheckQuickUpload = False
      Exit Function
    End If
  End With
CheckQuickUpload = True
End Function

Private Sub flxImport_DblClick()
Dim tmp
tmp = InputBox("Enter " & flxImport.TextMatrix(0, flxImport.ColSel), "Add New User", flxImport.TextMatrix(flxImport.RowSel, flxImport.ColSel))
flxImport.TextMatrix(flxImport.RowSel, flxImport.ColSel) = Trim(tmp)
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
flxImport.height = stsBar.Top - flxImport.Top - 250
flxImport.width = Me.width - flxImport.Left - 350
End Sub

Private Sub stsBar_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
frmLogs.txtLogs.Text = stsBar.Panels(2).Text: frmLogs.Show 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "cmdRefresh"
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
    ClearForm
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
Case "cmdOutput"
    If flxImport.Rows <= 1 Then
        MsgBox "Nothing to output", vbInformation
    Else
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
            OutputTable
    End If
End Select
End Sub

Private Sub ClearForm()
ClearTable
 Me.Caption = Me.Tag
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.Cols = 11
flxImport.TextMatrix(0, 0) = "User ID"
flxImport.TextMatrix(0, 1) = "Full Name"
flxImport.TextMatrix(0, 2) = "Password (Auto)"
flxImport.TextMatrix(0, 3) = "Email"
flxImport.TextMatrix(0, 4) = "Phone"
flxImport.TextMatrix(0, 5) = "Role 1"
flxImport.TextMatrix(0, 6) = "Role 2"
flxImport.TextMatrix(0, 7) = "Project 1"
flxImport.TextMatrix(0, 8) = "Project 2"
flxImport.TextMatrix(0, 9) = "Description"
flxImport.TextMatrix(0, 10) = "Validation"
flxImport.Rows = 2
End Sub

Public Function IncorrectHeaderDetails() As Boolean
    If flxImport.TextMatrix(0, 0) <> "User ID" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 1) <> "Full Name" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 2) <> "Password (Auto)" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 3) <> "Email" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 4) <> "Phone" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 5) <> "Role 1" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 6) <> "Role 2" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 7) <> "Project 1" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 8) <> "Project 2" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 9) <> "Description" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 10) <> "Validation" Then IncorrectHeaderDetails = True
End Function

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
    'xlObject.Sheets("Sheet2").Range("A1").Value = "1 - Only edit values in the column(s) colored green"
    'xlObject.Sheets("Sheet2").Range("A2").Value = "2 - Do not Add, Delete or Modify Rows and Column's Position, Color or Order"
    xlObject.Sheets("Sheet2").Range("A3").Value = "3 - The same sheet will be uploaded using SuperQC tools"
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
  xlObject.Sheets("Sheet3").Range("A1").Value = "1Company"
  xlObject.Sheets("Sheet3").Range("A2").Value = "ISAP_TRAINING"
  xlObject.Sheets("Sheet3").Range("B1").Value = "company_Test Admin"
  xlObject.Sheets("Sheet3").Range("B2").Value = "company_Test Coordinator"
  xlObject.Sheets("Sheet3").Range("B3").Value = "company_Fix Lead"
  xlObject.Sheets("Sheet3").Range("B4").Value = "company_Tester"
  xlObject.Sheets("Sheet3").Range("B5").Value = "company_Fixer"
  xlObject.Sheets("Sheet3").Range("B6").Value = "company_TesterFixer"
  xlObject.Sheets("Sheet3").Range("B7").Value = "company_Reporter"
  xlObject.Sheets("Sheet3").Range("B8").Value = "company_End Tester"
  xlObject.Sheets("Sheet3").Name = "Source"
  xlObject.Sheets("Sheet1").Columns("F:G").Select
    With xlObject.Sheets("Sheet1").Columns("F:G").Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, "company_Test Admin, company_Test Coordinator, company_Fix Lead, company_Tester, company_Fixer, company_TesterFixer, company_Reporter, company_End Tester"
    End With
  xlObject.Sheets("Sheet1").Columns("H:I").Select
    With xlObject.Sheets("Sheet1").Columns("H:I").Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, "1Company_RELEASE4, 1Company_RELEASE4_TRAINING"
    End With
  'xlObject.Visible = True
  
  curTab = "US_BPT-01"
  xlObject.Sheets("Sheet1").Name = curTab
  flxImport.FixedCols = 0
  flxImport.FixedRows = 0
  flxImport.row = 0
  flxImport.col = 0
  Pause 1
  flxImport.RowSel = flxImport.Rows - 1
  flxImport.ColSel = flxImport.Cols - 1
  Clipboard.Clear
  
  Clipboard.SetText flxImport.Clip
  flxImport.FixedCols = 0
  flxImport.FixedRows = 1

  xlObject.Sheets(curTab).Range("A1").Select
  xlObject.Sheets(curTab).Paste

'On Error Resume Next
    xlObject.Sheets(curTab).Range("A:K").Select

    xlObject.Sheets(curTab).Range("A:K").Borders(xlDiagonalDown).LineStyle = xlNone
    xlObject.Sheets(curTab).Range("A:K").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlObject.Sheets(curTab).Range("A:K").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:K").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:K").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:K").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:K").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:K").Borders(xlInsideHorizontal)
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
    xlObject.Sheets(curTab).Range("A:K").Select
    xlObject.Sheets(curTab).Range("A:K").EntireColumn.AutoFit
    xlObject.Sheets(curTab).Range("A1").Select

    xlObject.Sheets(curTab).Range("A1").AddComment
    xlObject.Sheets(curTab).Range("A1").Comment.Visible = False
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & mdiMain.Caption & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""
    
    xlObject.Sheets(curTab).Range("C:C").Interior.ColorIndex = 3
    xlObject.Sheets(curTab).Range("K:K").Interior.ColorIndex = 3
    'xlObject.Sheets(curTab).Protection.AllowEditRanges.Add Title:="Range1", Range:=xlObject.Sheets(curTab).Range("A:K")
    'xlObject.Sheets(curTab).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  xlObject.Workbooks(1).SaveAs "US_BPT-01" & "-" & Format(Now, "mmddyyyy HHMMSS AMPM")
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate

  Set xlWB = Nothing
  Set xlObject = Nothing
  FXGirl.EZPlay FXExportToExcel
  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export to MS Excel completed.": Exit Sub:
OutErr:     MsgBox Err.Description, vbCritical: xlObject.Visible = True: xlObject.ActiveWindow.Activate: Set xlWB = Nothing: Set xlObject = Nothing
On Error GoTo 0
End Sub

 Function GetCommentText(rCommentCell As Range)
     Dim strGotIt As String
         On Error Resume Next
         strGotIt = WorksheetFunction.Clean _
             (rCommentCell.Comment.Text)
         GetCommentText = strGotIt
         On Error GoTo 0
End Function

