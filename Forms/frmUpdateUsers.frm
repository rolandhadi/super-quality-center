VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUpdateUsers 
   Caption         =   "Upload Users Module"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   Icon            =   "frmUpdateUsers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   10530
   Tag             =   "Upload Users Module"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdQuickReset 
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
      Left            =   2100
      Picture         =   "frmUpdateUsers.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Import step description and expected results from an excel file"
      Top             =   540
      Width           =   1365
   End
   Begin VB.TextBox txtUserID 
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
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   1935
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
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
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdAllUsers"
            Object.ToolTipText     =   "Export All User Details to Excel"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin VB.CheckBox chkSendEmail 
         Caption         =   "Send Email to Users"
         Height          =   315
         Left            =   6480
         TabIndex        =   8
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
         Picture         =   "frmUpdateUsers.frx":114E
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
         Picture         =   "frmUpdateUsers.frx":3972
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
            Picture         =   "frmUpdateUsers.frx":4118
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateUsers.frx":43AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateUsers.frx":463C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   4755
      Left            =   60
      TabIndex        =   1
      Top             =   1200
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   8387
      _Version        =   393216
      Cols            =   5
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
            Picture         =   "frmUpdateUsers.frx":48CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateUsers.frx":4FDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateUsers.frx":56EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateUsers.frx":5E00
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
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmUpdateUsers.frx":6512
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
            Picture         =   "frmUpdateUsers.frx":6A63
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateUsers.frx":6D45
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateUsers.frx":7296
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateUsers.frx":77E7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "*Double click cells to update manually"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1020
      Width           =   2835
   End
End
Attribute VB_Name = "frmUpdateUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Type US_BPT
    UserID As String
    NewPassword As String
    NewEmailAddress As String
    NewDescription As String
    Log As String
End Type

Private All_US() As US_BPT
Private HasIssue As Boolean
Private HasUploadIssue  As Integer


Private Function LoadToArray()
Dim lastrow, i, EndArr
lastrow = flxImport.Rows - 1
ReDim All_US(0)
EndArr = -1
For i = 1 To lastrow
    If Trim(flxImport.TextMatrix(i, 0)) = "" Then
        All_US(EndArr).Log = All_US(EndArr).Log & vbCrLf & "Line " & i & " is blank"
    Else
        EndArr = EndArr + 1
        ReDim Preserve All_US(EndArr)
        All_US(EndArr).UserID = LCase(flxImport.TextMatrix(i, 0))
        All_US(EndArr).NewPassword = flxImport.TextMatrix(i, 1)
        All_US(EndArr).NewEmailAddress = flxImport.TextMatrix(i, 2)
        All_US(EndArr).NewDescription = flxImport.TextMatrix(i, 3)
    End If
Next
End Function

Function LoadToQC()
Dim i, j, TimeStart
Dim tmpComp, ChangeAction
stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = ""
ReDim All_Folders(0)
TimeStart = Now
mdiMain.pBar.Max = UBound(All_US) + 1
For i = LBound(All_US) To UBound(All_US)
    On Error Resume Next
    ChangeAction = Change_User(All_US(i))
    If Err.Number = -2147220181 Then Err.Clear
    If ChangeAction = "Updated" Then
        If Err.Number = 0 Or Err.Number = -2147220183 Then SendUpdateNotification All_US(i)
    End If
    If Err.Number <> 0 And Err.Number <> -2147220183 Then
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[UPDATED USER: (FAILED) " & Now & " " & All_US(i).UserID & "] " & Err.Description
        HasUploadIssue = HasUploadIssue + 1
    Else
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[UPDATED USER: (PASSED) " & Now & " " & All_US(i).UserID & "]"
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
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = UBound(All_US) + 1 & " Users(s) updated successfully. Email was sent to the user(s). (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log (Start: " & TimeStart & ") (End: " & Now & ")"
    If HasUploadIssue <> 0 Then
      Dim tmpFile As New clsFiles
      frmLogs.Caption = App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log"
      frmLogs.txtLogs.Text = tmpFile.ReadFromFile_FAILED(App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log")
      frmLogs.Show 1
    End If
End Function

Private Function Change_User(curUser As US_BPT)
Dim SA As New SAapi
SA.Login "http://qctesting.companyemail.net/qcbin", "qcadmin", "companyb2012"
If curUser.NewPassword <> "" Then
    SA.SetUserProperty curUser.UserID, 7, curUser.NewPassword
End If
If curUser.NewEmailAddress <> "" Then
    SA.SetUserProperty curUser.UserID, 4, curUser.NewEmailAddress
End If
If curUser.NewDescription <> "" Then
    SA.SetUserProperty curUser.UserID, 6, curUser.NewDescription
End If
If UCase(Trim(curUser.NewDescription)) = "[BLANK]" Or UCase(Trim(curUser.NewDescription)) = "[EMPTY]" Then
    SA.SetUserProperty curUser.UserID, 6, ""
End If
SA.Logout
Change_User = "Updated"
End Function

Private Function Dump_All_Users()
Dim SA As New SAapi
Dim tmp As String, tmpD As String
Dim FileFunct As New clsFiles
Dim testFile As Workbook
SA.Login "http://qctesting.companyemail.net/qcbin", "qcadmin", "companyb2012"
tmp = SA.GetAllUsers
SA.Logout
'tmp = Replace(tmp, vbCrLf, "")
tmp = Replace(tmp, "<?xml version=""1.0""?>", "")
tmp = Replace(tmp, "<GetAllUsers>", "")
tmp = Replace(tmp, "</GetAllUsers>", "")
tmp = Replace(tmp, "<TDXItem>", "")
tmp = Replace(tmp, "</TDXItem>", vbCrLf)
tmp = Replace(tmp, "<USER_ID>", """")
tmp = Replace(tmp, "<USER_NAME>", """")
tmp = Replace(tmp, "<ACC_IS_ACTIVE>", """")
tmp = Replace(tmp, "<FULL_NAME>", """")
tmp = Replace(tmp, "<EMAIL>", """")
tmp = Replace(tmp, "<USER_PASSWORD>", """")
tmp = Replace(tmp, "<DESCRIPTION>", """")
tmp = Replace(tmp, "<PHONE_NUMBER>", """")
tmp = Replace(tmp, "<LAST_UPDATE>", """")
tmp = Replace(tmp, "<US_DOM_AUTH>", """")
tmp = Replace(tmp, "<US_REPORT_ROLE>", """")
tmp = Replace(tmp, "</USER_ID>", """,")
tmp = Replace(tmp, "</USER_NAME>", """,")
tmp = Replace(tmp, "</ACC_IS_ACTIVE>", """,")
tmp = Replace(tmp, "</FULL_NAME>", """,")
tmp = Replace(tmp, "</EMAIL>", """,")
tmp = Replace(tmp, "</USER_PASSWORD>", """,")
tmp = Replace(tmp, "</DESCRIPTION>", """,")
tmp = Replace(tmp, "</PHONE_NUMBER>", """,")
tmp = Replace(tmp, "</LAST_UPDATE>", """,")
tmp = Replace(tmp, "</US_DOM_AUTH>", """,")
tmp = Replace(tmp, "</US_REPORT_ROLE>", """,")
tmp = "USER ID" & "," & "USER NAME" & "," & "ACCOUNT IS ACTIVE?" & "," & "FULL NAME" & "," & "EMAIL ADDRESS" & "," & "USER PASSWORD" & "," & "DESCRIPTION" & "," & "PHONE NUMBER" & "," & "LAST UPDATE" & "," & "USER DOM AUTHORIZATION" & "," & "USER REPORT ROLE" & tmp
tmpD = Format(Now, "mm-dd-yyyy hhmmAM/PM")
FileFunct.FileWrite App.path & "\SQC Logs" & "\Users - " & tmpD & ".csv", tmp
If MsgBox("Successfully exported to " & App.path & "\SQC Logs" & "\Users - " & tmpD & ".csv" & vbCrLf & "Do you want to launch the extracted file?", vbYesNo) = vbYes Then
  Shell "explorer.exe " & App.path & "\SQC Logs" & "\", vbNormalFocus
End If
End Function

Private Sub SendUpdateNotification(tmpUser As US_BPT)
Dim tmp
tmp = ""
  tmp = "User ID: " & "<b>" & tmpUser.UserID & "</b><br>"
  If tmpUser.NewPassword <> "" Then
    tmp = tmp & "New Password: " & "<b>" & tmpUser.NewPassword & "</b><br>"
  End If
  If tmpUser.NewEmailAddress <> "" Then
    tmp = tmp & "New Email Address: " & "<b>" & tmpUser.NewEmailAddress & "</b><br>"
  End If
  If tmpUser.NewDescription <> "" Then
    tmp = tmp & "New Assigned Team: " & "<b>" & tmpUser.NewDescription & "</b><br>"
  End If
  tmp = tmp & "HPQC Link: <b>" & "http://qctesting.companyemail.net/qcbin/" & "</b><br>"
  tmp = "Your HPQC Account is now updated. <br><br>" & tmp & "<br>To change your password, open <a href=""http://qctesting.companyemail.net/qcbin/"">HPQC</a> and access this link --> <b>https://eroom2.companyemail.com/eRoomReq/Files/Facility38/1CompanyWorld26-TestingitemsfromWorld8/0_b2e2b/PasswordReset.pdf</b> for guidelines on how to change password." & "<br><br>QC Support Contact: <b>1CompanyTesting-R4@companyemail.com</b><br>QC Support Hotline: <b>+65 8599 8076</b><br><br>Quick Documentation: <b>https://eroom2.companyemail.com/eRoom/Facility38/1CompanyWorld26-TestingitemsfromWorld8/0_ac7a8</b><br>" & "<br>" & "<b>***This is a HPQC automatic email notification. Do not reply.***</b>"
If chkSendEmail.Value = Checked Then
  QCConnection.SendMail tmpUser.UserID, "", "[HPQC ACCOUNTS] Your user account was successfully updated", tmp, "", "HTML"
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
         If UCase(Trim(.Range("B1").Value)) <> UCase(Trim("New Password")) Then
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
        flxImport.Cols = 5
        flxImport.Redraw = False
        
        'A - Load HPQC Folder Path
        'Should not be blank
        mdiMain.pBar.Max = lastrow + 2
        For i = 2 To lastrow
            
            
            flxImport.TextMatrix(i - 1, 0) = CleanTheString_PARAMS((Trim((.Range("A" & i).Value))))        'Change number and letter
            If Trim(.Range("A" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 4) = flxImport.TextMatrix(i - 1, 4) & "[User ID=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 1) = CleanTheString((Trim((.Range("B" & i).Value))))
            If UCase(flxImport.TextMatrix(i - 1, 1)) = "AUTO" Then
                flxImport.TextMatrix(i - 1, 1) = "welcome_" & LCase(flxImport.TextMatrix(i - 1, 0)) & "_" & Format(stringFunct.RandomNumber(10, 1), "00")
            End If
            
            flxImport.TextMatrix(i - 1, 2) = ((Trim((.Range("C" & i).Value))))
            If intFunct.ValidateEmail(Trim(.Range("C" & i).Value)) = False And flxImport.TextMatrix(i - 1, 2) <> "" Then
                flxImport.TextMatrix(i - 1, 4) = flxImport.TextMatrix(i - 1, 4) & "[Email=Invalid Format]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 3) = ((Trim((.Range("D" & i).Value))))
                        
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

Private Sub cmdQuickReset_Click()
Dim tmp As US_BPT, tmpR
Dim stringFunct As New clsStrings
On Error GoTo Err1
If Trim(txtUserID.Text) <> "" Then
  If MsgBox("Are you sure you want to reset the password of user " & Trim(txtUserID.Text) & "?", vbYesNo) = vbYes Then
    Randomize: tmpR = CInt(Rnd(1000) * 10000)
    If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
      tmp.UserID = LCase(Trim(txtUserID.Text))
      tmp.NewPassword = "welcome_" & LCase(tmp.UserID) & "_" & Format(stringFunct.RandomNumber(10, 1), "00")
      Change_User tmp
      SendUpdateNotification tmp
      MsgBox "Password reset for " & tmp.UserID & " completed!", vbInformation
      txtUserID.Text = ""
    End If
  End If
End If
Exit Sub
Err1:
MsgBox "The user " & tmp.UserID & " does not exist!", vbCritical
txtUserID.Text = ""
End Sub

Private Sub cmdUpload_Click()
Dim tmpR
If Trim(flxImport.TextMatrix(1, 0)) <> "" Then
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
    .TextMatrix(1, 4) = ""
    If Trim(.TextMatrix(1, 0)) = "" Then
      .TextMatrix(1, 4) = .TextMatrix(1, 4) & "[BLANK USER ID]"
      CheckQuickUpload = False
      Exit Function
    End If
    If intFunct.ValidateEmail(Trim(.TextMatrix(1, 2))) = False And Trim(.TextMatrix(1, 2)) <> "" Then
      .TextMatrix(1, 4) = .TextMatrix(1, 4) & "[INVALID EMAIL]"
      CheckQuickUpload = False
      Exit Function
    End If
  End With
CheckQuickUpload = True
End Function

Private Sub flxImport_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 67 And Shift = vbCtrlMask Then
    Clipboard.Clear
    Clipboard.SetText flxImport.Clip
End If
End Sub

Private Sub flxImport_DblClick()
Dim tmp
tmp = InputBox("Enter " & flxImport.TextMatrix(0, flxImport.ColSel), "Modify User", flxImport.TextMatrix(flxImport.RowSel, flxImport.ColSel))
flxImport.TextMatrix(flxImport.RowSel, flxImport.ColSel) = Trim(tmp)
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
Case "cmdAllUsers"
  If MsgBox("Are you sure you want to export all users?", vbYesNo) = vbYes Then
    Dump_All_Users
  End If
End Select
End Sub

Private Sub ClearForm()
ClearTable
Me.Caption = Me.Tag
txtUserID.Text = ""
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.Cols = 5
flxImport.TextMatrix(0, 0) = "User ID"
flxImport.TextMatrix(0, 1) = "New Password"
flxImport.TextMatrix(0, 2) = "New Email Address"
flxImport.TextMatrix(0, 3) = "New Description"
flxImport.TextMatrix(0, 4) = "Validation"
flxImport.Rows = 2
End Sub

Public Function IncorrectHeaderDetails() As Boolean
    If flxImport.TextMatrix(0, 0) <> "User ID" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 1) <> "New Password" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 2) <> "New Email Address" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 3) <> "New Description" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 4) <> "Validation" Then IncorrectHeaderDetails = True
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

    xlObject.Sheets(curTab).Range("A:A").Interior.ColorIndex = 3
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

