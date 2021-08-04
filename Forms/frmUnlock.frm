VERSION 5.00
Begin VB.Form frmUnlock 
   Caption         =   "Unlock Items"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   Icon            =   "frmUnlock.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   5295
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   1500
      TabIndex        =   3
      Top             =   660
      Width           =   3555
      Begin VB.CommandButton cmdSeeAll 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "See All"
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
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Select User"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   660
         Width           =   1755
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All Users"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.TextBox txtUserName 
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
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   3075
      End
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
      ItemData        =   "frmUnlock.frx":08CA
      Left            =   1500
      List            =   "frmUnlock.frx":08E0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   3555
   End
   Begin VB.CommandButton cmdUnlock 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Unlock"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QC Module:"
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
      Left            =   270
      TabIndex        =   2
      Top             =   300
      Width           =   1095
   End
End
Attribute VB_Name = "frmUnlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSeeAll_Click()
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim LockType, i, tmp, txt As New clsStrings
LockType = GetLockType
Set objCommand = QCConnection.Command
objCommand.CommandText = "SELECT LK_OBJECT_KEY, LK_USER, LK_CLIENT_MACHINE_NAME FROM LOCKS WHERE LK_OBJECT_TYPE = '" & LockType & "' ORDER BY LK_OBJECT_KEY"
Set rs = objCommand.Execute
Select Case Trim(UCase(cmbModule.Text))
    Case "BUSINESS COMPONENTS"
        If rs.RecordCount <> 0 Then
            For i = 1 To rs.RecordCount
                tmp = tmp & Format(rs.FieldValue("LK_OBJECT_KEY"), "0000000000") & vbTab & txt.TextFill(rs.FieldValue("LK_USER"), 20, " ") & vbTab & txt.TextFill(rs.FieldValue("LK_CLIENT_MACHINE_NAME"), 30, "-") & vbTab & GetBusinessComponentFolderPath(rs.FieldValue("LK_OBJECT_KEY")) & "\" & GetFromTable(rs.FieldValue("LK_OBJECT_KEY"), "CO_ID", "CO_NAME", "COMPONENT") & vbCrLf
                rs.Next
            Next
            tmp = "UNIQUE ID " & vbTab & txt.TextFill("USER ID", 20, " ") & vbTab & txt.TextFill("MACHINE NAME", 50, " ") & vbTab & "LOCKED ITEM" & vbCrLf & tmp
            frmLogs.txtLogs.Text = tmp
            frmLogs.Show 1
        Else
            MsgBox "No locked items found"
        End If
    Case "TEST PLAN"
        If rs.RecordCount <> 0 Then
            For i = 1 To rs.RecordCount
                tmp = tmp & Format(rs.FieldValue("LK_OBJECT_KEY"), "0000000000") & vbTab & txt.TextFill(rs.FieldValue("LK_USER"), 20, " ") & vbTab & txt.TextFill(rs.FieldValue("LK_CLIENT_MACHINE_NAME"), 30, "-") & vbTab & GetTestFolderPath(GetFromTable(rs.FieldValue("LK_OBJECT_KEY"), "TS_TEST_ID", "TS_SUBJECT", "TEST")) & "\" & GetFromTable(rs.FieldValue("LK_OBJECT_KEY"), "TS_TEST_ID", "TS_NAME", "TEST") & vbCrLf
                rs.Next
            Next
            tmp = "UNIQUE ID " & vbTab & txt.TextFill("USER ID", 20, " ") & vbTab & txt.TextFill("MACHINE NAME", 50, " ") & vbTab & "LOCKED ITEM" & vbCrLf & tmp
            frmLogs.txtLogs.Text = tmp
            frmLogs.Show 1
        Else
            MsgBox "No locked items found"
        End If
    Case "TEST LAB (TEST INSTANCE)"
        If rs.RecordCount <> 0 Then
            For i = 1 To rs.RecordCount
                tmp = tmp & Format(rs.FieldValue("LK_OBJECT_KEY"), "0000000000") & vbTab & txt.TextFill(rs.FieldValue("LK_USER"), 20, " ") & vbTab & txt.TextFill(rs.FieldValue("LK_CLIENT_MACHINE_NAME"), 30, "-") & vbTab & GetTestInstanceFolderPath(rs.FieldValue("LK_OBJECT_KEY")) & "\" & GetFromTable(GetFromTable(rs.FieldValue("LK_OBJECT_KEY"), "TC_TESTCYCL_ID", "TC_TEST_ID", "TESTCYCL"), "TS_TEST_ID", "TS_NAME", "TEST") & vbCrLf
                rs.Next
            Next
            tmp = "UNIQUE ID " & vbTab & txt.TextFill("USER ID", 20, " ") & vbTab & txt.TextFill("MACHINE NAME", 50, " ") & vbTab & "LOCKED ITEM" & vbCrLf & tmp
            frmLogs.txtLogs.Text = tmp
            frmLogs.Show 1
        Else
            MsgBox "No locked items found"
        End If
    Case "TEST LAB (RUN)"
        If rs.RecordCount <> 0 Then
            For i = 1 To rs.RecordCount
                tmp = tmp & Format(rs.FieldValue("LK_OBJECT_KEY"), "0000000000") & vbTab & txt.TextFill(rs.FieldValue("LK_USER"), 20, " ") & vbTab & txt.TextFill(rs.FieldValue("LK_CLIENT_MACHINE_NAME"), 30, "-") & vbTab & GetFromTable(rs.FieldValue("LK_OBJECT_KEY"), "RN_RUN_ID", "RN_RUN_NAME", "RUN") & vbCrLf
                rs.Next
            Next
            tmp = "UNIQUE ID " & vbTab & txt.TextFill("USER ID", 20, " ") & vbTab & txt.TextFill("MACHINE NAME", 50, " ") & vbTab & "LOCKED ITEM" & vbCrLf & tmp
            frmLogs.txtLogs.Text = tmp
            frmLogs.Show 1
        Else
            MsgBox "No locked items found"
        End If
    Case "DEFECTS"
        If rs.RecordCount <> 0 Then
            For i = 1 To rs.RecordCount
                tmp = tmp & Format(rs.FieldValue("LK_OBJECT_KEY"), "0000000000") & vbTab & txt.TextFill(rs.FieldValue("LK_USER"), 20, " ") & vbTab & txt.TextFill(rs.FieldValue("LK_CLIENT_MACHINE_NAME"), 30, "-") & vbTab & GetFromTable(rs.FieldValue("LK_OBJECT_KEY"), "BG_BUG_ID", "BG_SUMMARY", "BUG") & vbCrLf
                rs.Next
            Next
            tmp = "UNIQUE ID " & vbTab & txt.TextFill("USER ID", 20, " ") & vbTab & txt.TextFill("MACHINE NAME", 50, " ") & vbTab & "LOCKED ITEM" & vbCrLf & tmp
            frmLogs.txtLogs.Text = tmp
            frmLogs.Show 1
        Else
            MsgBox "No locked items found"
        End If
    Case "CUSTOMIZATION"
        If rs.RecordCount <> 0 Then
            For i = 1 To rs.RecordCount
                tmp = tmp & Format(rs.FieldValue("LK_OBJECT_KEY"), "0000000000") & vbTab & txt.TextFill(rs.FieldValue("LK_USER"), 20, " ") & vbTab & txt.TextFill(rs.FieldValue("LK_CLIENT_MACHINE_NAME"), 30, "-") & vbTab & "Customization" & vbCrLf
                rs.Next
            Next
            tmp = "UNIQUE ID " & vbTab & txt.TextFill("USER ID", 20, " ") & vbTab & txt.TextFill("MACHINE NAME", 50, " ") & vbTab & "LOCKED ITEM" & vbCrLf & tmp
            frmLogs.txtLogs.Text = tmp
            frmLogs.Show 1
        Else
            MsgBox "No locked items found"
        End If
End Select
End Sub

Private Sub cmdUnlock_Click()
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim LockType
Dim LUser
Dim tmpR
    Set objCommand = QCConnection.Command
    LockType = GetLockType
    LUser = LCase(Trim(txtUserName.Text))
    If optAll.Value = True Then
        objCommand.CommandText = "SELECT * FROM LOCKS WHERE LK_OBJECT_TYPE = '" & LockType & "'"
        Set rs = objCommand.Execute
        If rs.RecordCount <> 0 Then
            If MsgBox("Are you sure you want to unlock " & rs.RecordCount & " locked item(s) in " & cmbModule.Text & "?", vbYesNo) = vbYes Then
                Randomize: tmpR = CInt(Rnd(1000) * 10000)
                If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
                    QCConnection.SendMail "user@companyemail.com"", "[UNLOCK] " & rs.RecordCount & " user(s) are kicked by " & curUser & " in " & curDomain & "-" & curProject & " in " & cmbModule.Text & " module", rs.RecordCount & " user(s) are kicked by " & curUser & " in " & curDomain & "-" & curProject & " in " & cmbModule.Text & " module", "", "HTML"
                    Set rs = Nothing
                    objCommand.CommandText = "DELETE FROM LOCKS WHERE LK_OBJECT_TYPE = '" & LockType & "'"
                    Set rs = objCommand.Execute
                    MsgBox "Unlocked successfully"
                Else
                    MsgBox "Invalid pass key", vbCritical
                End If
            End If
        Else
            MsgBox "No locked items found"
        End If
    Else
        If LUser = "" Then
            MsgBox "Please enter user name"
            Exit Sub
        End If
        objCommand.CommandText = "SELECT * FROM LOCKS WHERE LK_OBJECT_TYPE = '" & LockType & "' AND LK_USER = '" & LUser & "'"
        Set rs = objCommand.Execute
        If rs.RecordCount <> 0 Then
            If MsgBox("Are you sure you want to unlock " & rs.RecordCount & " locked item(s) in " & cmbModule.Text & "?", vbYesNo) = vbYes Then
                Randomize: tmpR = CInt(Rnd(1000) * 10000)
                If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
                    Set rs = Nothing
                    objCommand.CommandText = "DELETE FROM LOCKS WHERE LK_OBJECT_TYPE = '" & LockType & "' AND LK_USER = '" & LUser & "'"
                    Set rs = objCommand.Execute
                    QCConnection.SendMail "user@companyemail.com", "", "[UNLOCK] " & LUser & " user is kicked by " & curUser & " in " & curDomain & "-" & curProject & " in " & cmbModule.Text & " module", LUser & " user is kicked by " & curUser & " in " & curDomain & "-" & curProject & " in " & cmbModule.Text & " module", "", "HTML"
                    MsgBox "Unlocked successfully"
                Else
                    MsgBox "Invalid pass key", vbCritical
                End If
            End If
        Else
            MsgBox "No locked items found"
        End If
    End If
End Sub

Function GetLockType() As String
Select Case Trim(UCase(cmbModule.Text))
    Case "BUSINESS COMPONENTS"
        GetLockType = "COMPONENT"
    Case "TEST PLAN"
        GetLockType = "TEST"
    Case "TEST LAB (TEST INSTANCE)"
        GetLockType = "TESTCYCL"
    Case "TEST LAB (RUN)"
        GetLockType = "RUN"
    Case "DEFECTS"
        GetLockType = "BUG"
    Case "CUSTOMIZATION"
        GetLockType = "CUSTOMIZATION"
    Case Else
        GetLockType = "ROZS"
End Select
End Function

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
cmbModule.ListIndex = 0
End Sub

Private Sub optAll_Click()
txtUserName.Text = ""
txtUserName.Enabled = False
End Sub

Private Sub Option1_Click()
txtUserName.Text = ""
txtUserName.Enabled = True
End Sub
