VERSION 5.00
Begin VB.Form frmTray 
   BorderStyle     =   0  'None
   ClientHeight    =   1410
   ClientLeft      =   150
   ClientTop       =   810
   ClientWidth     =   5625
   Icon            =   "frmTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer tmrCheck 
      Interval        =   5000
      Left            =   60
      Top             =   120
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Menu"
      Begin VB.Menu mnuActivate 
         Caption         =   "Activate"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu mnuTerminate 
         Caption         =   "Terminate"
      End
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit
'Declare a user-defined variable to pass to the Shell_NotifyIcon
'function.
Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

Private Type AllReports
    ReportKey As String
    ReportName As String
    Run As Boolean
    Time As Date
    Target As String
    Days As String
    SQL As String
End Type

'Declare the constants for the API function. These constants can be
'found in the header file Shellapi.h.

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the
'taskbar status area.
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Private Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.

'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA
Dim FileFunct As New clsFiles
Dim stringFunct As New clsStrings
Dim myReports() As AllReports
Dim LastRunAt As Date

'**********************************************************
'**********************************************************
Private Sub Form_Load()
    'Set the individual values of the NOTIFYICONDATA data type.
    nid.cbSize = Len(nid)
    nid.hwnd = Me.hwnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = Me.Icon
    nid.szTip = "RR-R1 Testing Team - SuperQC " & curDomain & "-" & curProject & vbNullChar
    
    'Call the Shell_NotifyIcon function to add the icon to the taskbar
    'status area.
    Shell_NotifyIcon NIM_ADD, nid
End Sub

'**********************************************************
'**********************************************************
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    DoEvents

    Dim msg As Long

    If Me.ScaleMode = vbPixels Then
     msg = x
    Else
     msg = x / Screen.TwipsPerPixelX
    End If

    Select Case msg
       Case WM_LBUTTONDOWN
            mnuPopup.Visible = False
       Case WM_LBUTTONUP
            mnuPopup.Visible = False
       Case WM_LBUTTONDBLCLK
            mnuPopup.Visible = False
       Case WM_RBUTTONDOWN
            mnuPopup.Visible = False
       Case WM_RBUTTONUP
            PopupMenu mnuPopup
       Case WM_RBUTTONDBLCLK
            mnuPopup.Visible = False
    End Select
End Sub


Private Sub Form_Terminate()
On Error Resume Next
QCConnection.Disconnect
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
.uCallBackMessage = vbNull
.hIcon = vbNull
.szTip = "" & vbNullChar
End With
Shell_NotifyIcon NIM_DELETE, nid
QCConnection.DisconnectProject
QCConnection.Disconnect
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
QCConnection.Disconnect
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
.uCallBackMessage = vbNull
.hIcon = vbNull
.szTip = "" & vbNullChar
End With
Shell_NotifyIcon NIM_DELETE, nid
QCConnection.DisconnectProject
QCConnection.Disconnect
End Sub

Private Sub mnuActivate_Click()
QCConnection.RefreshConnectionState
If QCConnection.Connected = False Then
    On Error Resume Next
    QCConnection.InitConnectionEx (curQCInstance)
    QCConnection.Logout
    QCConnection.Login ADMIN_ID, ADMIN_PASS
    QCConnection.Connect curDomain, curProject
    If QCConnection.Connected = False Then MsgBox "Session was disconnected. Please re-load the SuperQC tool.": If MsgBox("Do you want to close this session?", vbYesNo) = vbYes Then End
End If
mdiMain.Show
End Sub

Private Sub mnuLogOff_Click()
If MsgBox("Are you sure you want to Log Off?", vbYesNo) = vbYes Then
    Unload mdiMain
    frmLogin.Show
    Unload Me
End If
End Sub

Private Sub mnuTerminate_Click()
If MsgBox("Are you sure you want to terminate this application?", vbYesNo) = vbYes Then
    On Error Resume Next
    QCConnection.Disconnect
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uCallBackMessage = vbNull
        .hIcon = vbNull
        .szTip = "" & vbNullChar
    End With
    Shell_NotifyIcon NIM_DELETE, nid
    End
End If
End Sub

'Private Sub tmrCheck_Timer()
'QCConnection.RefreshConnectionState
'If QCConnection.Connected = False Then
'    On Error Resume Next
'    QCConnection.InitConnectionEx ("https://qcurl.saas.hp.com/qcbin")
'    QCConnection.Logout
'    QCConnection.Login ADMIN_ID, ADMIN_PASS
'    QCConnection.Connect curDomain, curProject
'    If QCConnection.Connected = False Then MsgBox "Session was disconnected. Please re-load the SuperQC tool.": If MsgBox("Do you want to close this session?", vbYesNo) = vbYes Then End
'End If
'DoEvents
'On Error Resume Next
'If Format(LastRunAt, "hh:mm") = Format(Now, "hh:mm") Then Exit Sub
'LoadReports
'ExecuteReport
'LastRunAt = Format(Now, "hh:mm")
'On Error GoTo 0
'End Sub

Private Sub LoadReports()
Dim tmpContent1, tmpContent2, i, tmp1, tmp2, cnt
Dim tmpName, tmpParent
ReDim myReports(0)
tmpContent1 = FileFunct.ReadFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh")
tmp1 = Split(tmpContent1, "~")
For i = LBound(tmp1) To UBound(tmp1)
    If Left(tmp1(i), 1) = "F" And stringFunct.StrIn(tmp1(i), "R") Then
        ReDim Preserve myReports(cnt)
        myReports(cnt).ReportKey = tmp1(i)
        tmpName = stringFunct.GetValueFromKey(CStr(tmp1(i + 1)), "NAME")
        myReports(cnt).ReportName = tmpName
        tmpParent = stringFunct.GetValueFromKey(CStr(tmp1(i + 1)), "PARENTID")
         tmpContent2 = FileFunct.ReadKeyFromFile(App.path & "\SQC DAT" & "\" & "myReports01.hxh", "~" & tmp1(i) & "~")
         If Trim(tmpContent2) = "" Then Exit Sub
         tmp2 = stringFunct.GetValueFromKey(CStr(tmpContent2), "RUN")
         If UCase(tmp2) = "YES" Then
            myReports(cnt).Run = True
         Else
            myReports(cnt).Run = True
         End If
         tmp2 = stringFunct.GetValueFromKey(CStr(tmpContent2), "TIME")
         If Left(tmp2, 2) = 24 Then tmp2 = "00" & Right(tmp2, Len(tmp2) - 2)
         myReports(cnt).Time = Format(CDate(tmp2), "hh:mm")
         tmp2 = stringFunct.GetValueFromKey(CStr(tmpContent2), "TARGET")
         myReports(cnt).Target = tmp2
         tmp2 = stringFunct.GetValueFromKey(CStr(tmpContent2), "DAYS")
         myReports(cnt).Days = tmp2
         tmp2 = stringFunct.GetValueFromKey(CStr(tmpContent2), "SQL")
         myReports(cnt).SQL = tmp2
         cnt = cnt + 1
    End If
Next
End Sub

Private Sub ExecuteReport()
Dim i, reportStart, reportEnd
For i = LBound(myReports) To UBound(myReports)
    If myReports(i).Run = True Then
        If stringFunct.StrIn(GetRunDays(myReports(i).Days), Format(Now, "dddd")) Then
            If Format(Now, "hh:mm") = Format(myReports(i).Time, "hh:mm") Then
                reportStart = Now
                GenerateReport myReports(i)
                reportEnd = Now
                FileFunct.FileAppend App.path & "\SQC Logs" & "\" & myReports(i).ReportKey & ".txt", "[" & Format(Now, "mm/dd/yyyy") & "] " & myReports(i).ReportName & " report creation started at " & reportStart & " and finished at " & reportEnd & "..."
            End If
        End If
    End If
Next
End Sub

Private Function GetRunDays(x) As String
Dim tmp
If Mid(x, 1, 1) = 1 Then
    tmp = tmp & " Monday "
End If
If Mid(x, 2, 1) = 1 Then
    tmp = tmp & " Tuesday "
End If
If Mid(x, 3, 1) = 1 Then
    tmp = tmp & " Wednesday "
End If
If Mid(x, 4, 1) = 1 Then
    tmp = tmp & " Thursday "
End If
If Mid(x, 5, 1) = 1 Then
    tmp = tmp & " Friday "
End If
If Mid(x, 6, 1) = 1 Then
    tmp = tmp & " Saturday "
End If
If Mid(x, 7, 1) = 1 Then
    tmp = tmp & " Sunday "
End If
GetRunDays = tmp
End Function

Private Sub GenerateReport(tmpReport As AllReports)
Dim rs As TDAPIOLELib.Recordset
Dim AllScripts, AllColumns()
Dim objCommand, LastFolderID As String, LastFolder As String
Dim i, k, p, tmp
Dim strPath
Dim tmpF

ReDim AllColumns(0)
    strPath = tmpReport.SQL
    Set objCommand = QCConnection.Command
    objCommand.CommandText = strPath
    Set rs = objCommand.Execute
    For i = 0 To rs.ColCount - 1
        ReDim Preserve AllColumns(i)
        AllColumns(i) = rs.ColName(i)
        AllScripts = AllScripts & Replace(rs.ColName(i), ";", ":") & vbTab
        rs.Next
    Next
    AllScripts = AllScripts & vbCrLf
    rs.First
    k = 0
    For i = 1 To rs.RecordCount
                k = k + 1
                For p = 0 To rs.ColCount - 1
                    Select Case UCase(Trim(AllColumns(p)))
                        Case UCase("RequirementFolderPath")
                                If rs.FieldValue("RQ_REQ_ID") <> "" Then
                                    If LastFolderID <> rs.FieldValue("RQ_REQ_ID") Then
                                        tmpF = GetRequirementFolderPath(rs.FieldValue("RQ_REQ_ID"))
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    Else
                                        tmpF = LastFolder
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    End If
                                    LastFolderID = rs.FieldValue("RQ_REQ_ID")
                                Else
                                    If LastFolderID <> rs.FieldValue("Requirement ID") Then
                                        tmpF = GetRequirementFolderPath(rs.FieldValue("Requirement ID"))
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    Else
                                        tmpF = LastFolder
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    End If
                                    LastFolderID = rs.FieldValue("Requirement ID")
                                End If
                        Case UCase("TestSetFolderPath")
                                If rs.FieldValue("CY_CYCLE_ID") <> "" Then
                                    If LastFolderID <> rs.FieldValue("CY_CYCLE_ID") Then
                                        tmpF = GetTestSetFolderPath(rs.FieldValue("CY_CYCLE_ID"))
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    Else
                                        tmpF = LastFolder
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    End If
                                    LastFolderID = rs.FieldValue("CY_CYCLE_ID")
                                Else
                                    If LastFolderID <> rs.FieldValue("Test Set ID") Then
                                        tmpF = GetTestSetFolderPath(rs.FieldValue("Test Set ID"))
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    Else
                                        tmpF = LastFolder
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    End If
                                    LastFolderID = rs.FieldValue("Test Set ID")
                                End If
                        Case UCase("BusinessComponentFolderPath")
                                If rs.FieldValue("CO_ID") <> "" Then
                                    If LastFolderID <> rs.FieldValue("CO_ID") Then
                                        tmpF = GetBusinessComponentFolderPath(rs.FieldValue("CO_ID"))
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    Else
                                        tmpF = LastFolder
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    End If
                                    LastFolderID = rs.FieldValue("CO_ID")
                                Else
                                    If LastFolderID <> rs.FieldValue("Component ID") Then
                                        tmpF = GetBusinessComponentFolderPath(rs.FieldValue("Component ID"))
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    Else
                                        tmpF = LastFolder
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    End If
                                    LastFolderID = rs.FieldValue("Component ID")
                                End If
                        Case UCase("TestFolderPath")
                                If rs.FieldValue("TS_SUBJECT") <> "" Then
                                    If LastFolderID <> rs.FieldValue("TS_SUBJECT") Then
                                        tmpF = GetTestFolderPath(rs.FieldValue("TS_SUBJECT"))
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    Else
                                        tmpF = LastFolder
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    End If
                                    LastFolderID = rs.FieldValue("TS_SUBJECT")
                                Else
                                    If LastFolderID <> rs.FieldValue("Subject") Then
                                        tmpF = GetTestFolderPath(rs.FieldValue("Subject"))
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    Else
                                        tmpF = LastFolder
                                        AllScripts = AllScripts & (tmpF) & vbTab
                                        LastFolder = tmpF
                                    End If
                                    LastFolderID = rs.FieldValue("Subject")
                                End If
                        Case Else
                            tmpF = ReverseCleanHTML(rs.FieldValue(Trim(AllColumns(p))))
                            If stringFunct.StrIn(rs.FieldValue(Trim(AllColumns(p))), vbCrLf) = True Or stringFunct.StrIn(rs.FieldValue(Trim(AllColumns(p))), Chr(10) + Chr(13)) = True Or stringFunct.StrIn(rs.FieldValue(Trim(AllColumns(p))), Chr(10)) = True Or stringFunct.StrIn(rs.FieldValue(Trim(AllColumns(p))), Chr(13)) = True Or stringFunct.StrIn(rs.FieldValue(Trim(AllColumns(p))), vbTab) = True Then
                                    AllScripts = AllScripts & Replace(ReverseCleanHTML(rs.FieldValue(Trim(AllColumns(p)))), vbTab, " ") & vbTab
                            Else
                                    AllScripts = AllScripts & ReverseCleanHTML(rs.FieldValue(Trim(AllColumns(p)))) & vbTab
                            End If
                        End Select
                Next
                AllScripts = AllScripts & vbCrLf
                rs.Next
    Next
    Set rs = Nothing
    tmp = OutputTable(ColumnLetter(UBound(AllColumns) + 1), tmpReport, AllScripts)
    QCConnection.SendMail curUser, "", "[HPQC AUTO REPORTS] " & tmp, "Auto Report " & tmp & " was successfully generated", "", "HTML"
End Sub

Private Function OutputTable(ColLetter As String, tmpReport As AllReports, AllScripts)
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim i, Protections
Dim curTab
Dim w, tmp

FileWrite App.path & "\SQC DAT" & "\" & "REPORT01" & ".xls", AllScripts

Set xlObject = New Excel.Application

On Error Resume Next
For Each w In xlObject.Workbooks
   w.Close savechanges:=False
Next w
On Error GoTo 0

Set xlWB = xlObject.Workbooks.Open(App.path & "\SQC DAT" & "\" & "REPORT01" & ".xls")
    xlObject.Sheets.Add
    xlObject.Sheets(1).Range("A1").Value = "Report Name: " & tmpReport.ReportName
    xlObject.Sheets(1).Range("A2").Value = "Code: " & tmpReport.SQL
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

    'xlObject.Sheets(curTab).Range(GetEditableFields).Interior.ColorIndex = 35

  'xlObject.Sheets(curTab).Protection.AllowEditRanges.Add Title:="Range1", Range:=xlObject.Sheets(curTab).Range(GetEditableFields)
  xlObject.Sheets(2).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  tmp = Trim(FileFunct.AddBackslash(tmpReport.Target) & "(AUTO) " & tmpReport.ReportName & "-" & Format(Now, "mm-dd-yyyy HH-MM AMPM") & ".xls")
  xlObject.Workbooks(1).SaveAs tmp
  xlObject.Workbooks.Close
  Set xlWB = Nothing
  Set xlObject = Nothing
  OutputTable = tmp
On Error GoTo 0
End Function
