VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLoadRequirement 
   Caption         =   "Upload Requirements Module"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   Icon            =   "frmLoadRequirement.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   10275
   Tag             =   "Upload Requirements Module"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
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
         Picture         =   "frmLoadRequirement.frx":08CA
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
         Picture         =   "frmLoadRequirement.frx":30EE
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
            Picture         =   "frmLoadRequirement.frx":3894
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadRequirement.frx":3B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadRequirement.frx":3DB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   5355
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   9446
      _Version        =   393216
      Cols            =   10
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
            Picture         =   "frmLoadRequirement.frx":4046
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadRequirement.frx":4758
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadRequirement.frx":4E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadRequirement.frx":557C
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
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmLoadRequirement.frx":5C8E
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
            Picture         =   "frmLoadRequirement.frx":61DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadRequirement.frx":64C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadRequirement.frx":6A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadRequirement.frx":6F63
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLoadRequirement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Type RQ_BPT
    Path_Location As String
    Requirement_Type  As String
    Requirement_Name  As String
    Additional_Info As String
    Author As String
    Peer_Review  As String
    Status As String
    ExistR1 As String
    ExistR2 As String
    WRIEF_Info As String
    Folder_Created As Boolean
    Log As String
End Type

Private All_RQ() As RQ_BPT
Private HasIssue As Boolean
Private HasUploadIssue  As Integer
Private newR, reqF, strID

Private Function LoadToArray()
Dim lastrow, i, EndArr
lastrow = flxImport.Rows - 1
ReDim All_RQ(0)
EndArr = -1
For i = 1 To lastrow
    If Trim(flxImport.TextMatrix(i, 0)) = "" Or Trim(flxImport.TextMatrix(i, 1)) = "" Then
        All_RQ(EndArr).Log = All_RQ(EndArr).Log & vbCrLf & "Line " & i & " is blank"
    Else
        EndArr = EndArr + 1
        ReDim Preserve All_RQ(EndArr)
        All_RQ(EndArr).Path_Location = flxImport.TextMatrix(i, 0)
        All_RQ(EndArr).Requirement_Type = flxImport.TextMatrix(i, 1)
        All_RQ(EndArr).Requirement_Name = flxImport.TextMatrix(i, 2)
        All_RQ(EndArr).Additional_Info = flxImport.TextMatrix(i, 3)
        All_RQ(EndArr).Author = flxImport.TextMatrix(i, 4)
        All_RQ(EndArr).Peer_Review = flxImport.TextMatrix(i, 5)
        All_RQ(EndArr).Status = flxImport.TextMatrix(i, 6)
        All_RQ(EndArr).ExistR1 = flxImport.TextMatrix(i, 7)
        All_RQ(EndArr).ExistR2 = flxImport.TextMatrix(i, 8)
        All_RQ(EndArr).WRIEF_Info = flxImport.TextMatrix(i, 9)
    End If
Next
End Function

Function LoadToQC()
Dim i, j, TimeStart
Dim tmpComp
stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = ""
ReDim All_Folders(0)
strID = ""
Set newR = Nothing
Set reqF = Nothing
TimeStart = Now
mdiMain.pBar.Max = UBound(All_RQ) + 1
For i = LBound(All_RQ) To UBound(All_RQ)
    On Error Resume Next
    If Create_Requirement(All_RQ(i)) = True Then Folder_Update (All_RQ(i).Path_Location)
    If Err.Number <> 0 Then
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[CREATE REQUIREMENT: (FAILED) " & Now & " " & All_RQ(i).Path_Location & "-" & All_RQ(i).Requirement_Name & "] " & Err.Description
        HasUploadIssue = HasUploadIssue + 1
    Else
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[CREATE REQUIREMENT: (PASSED) " & Now & " " & All_RQ(i).Path_Location & "-" & All_RQ(i).Requirement_Name & "]"
    End If
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(3).Picture: stsBar.Panels(2).Text = "Loading Requirements " & i + 1 & " out of " & UBound(All_RQ) + 1 & " (" & All_RQ(i).Requirement_Name & ")"
    Err.Clear
    On Error GoTo 0
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
FXGirl.EZPlay FXDataUploadCompleted
stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = UBound(All_RQ) + 1 & " Requirement(s) loaded successfully. (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log (Start: " & TimeStart & ") (End: " & Now & ")"
QCConnection.SendMail "user@companyemail.com", "", "[HPQC UPDATES] Requirement(s) loaded successfully by " & curUser & " in " & curDomain & "-" & curProject, UBound(All_RQ) + 1 & " Requirement(s) loaded successfully. (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log (Start: " & TimeStart & ") (End: " & Now & ")" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
QCConnection.SendMail curUser, "", "[HPQC UPDATES] Requirement(s) loaded successfully by " & curUser & " in " & curDomain & "-" & curProject, UBound(All_RQ) + 1 & " Requirement(s) loaded successfully. (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log (Start: " & TimeStart & ") (End: " & Now & ")" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
    If HasUploadIssue <> 0 Then
      Dim tmpFile As New clsFiles
      frmLogs.Caption = App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log"
      frmLogs.txtLogs.Text = tmpFile.ReadFromFile_FAILED(App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log")
      frmLogs.Show 1
    End If
End Function

Sub Start()
Debug.Print "New Session: " & Now
LoadToArray
LoadToQC
Debug.Print "New Finished: " & Now
End Sub

Function Folder_Update(X As String)
Dim i
For i = LBound(All_RQ) To UBound(All_RQ)
    If UCase(Trim(All_RQ(i).Path_Location)) = UCase(Trim(X)) Then
        All_RQ(i).Folder_Created = True
    End If
Next
End Function

'########################### Create New Requirements ###########################
Private Function Create_Requirement(tmpReq As RQ_BPT)
Dim ReqFilter, reqfol, FullPath
Dim NodeArray, strreqfol, stru
Dim WorkingDepth, ReqList, ParentReq
Dim NewRequirement, strName
Dim theReq, PathArray, strReqID
Dim i
Set reqF = QCConnection.ReqFactory
Set ReqFilter = reqF.Filter

    If tmpReq.Folder_Created = True Then
    Else
       reqfol = tmpReq.Path_Location
       FullPath = "\" & reqfol
       FullPath = Trim(FullPath)
       Dim pos%, ln%
       pos = InStr(1, FullPath, "\")
       If pos = 1 Then
           FullPath = Mid(FullPath, 2)
       End If
    
       ln = Len(FullPath)
       pos = InStr(ln - 1, FullPath, "\")
       If pos > 0 Then
           FullPath = Mid(FullPath, 1, ln - 1)
       End If
    
       NodeArray = Split(FullPath, "\")
       strreqfol = Split(reqfol, "\")
       stru = UBound(strreqfol)
    
       For WorkingDepth = LBound(NodeArray) To UBound(NodeArray)
           If WorkingDepth = LBound(NodeArray) Then
               'Set reqF = QCConnection.ReqFactory
               'Set ReqFilter = reqF.Filter
               Set ReqList = reqF.find(-1, "RQ_REQ_NAME", _
               NodeArray(WorkingDepth), TDREQMODE_FIND_EXACT)
           Else
               Set ReqList = reqF.find(ParentReq.ID, "RQ_REQ_NAME", _
               NodeArray(WorkingDepth), TDREQMODE_FIND_EXACT)
           End If
       
           Set ParentReq = Nothing
           Dim strItem$, reqID&, thePath$
    
           On Error Resume Next
           strItem = ReqList(1)
           If Err.Number = "-2147023483" Then
               Set newR = reqF.AddItem(Null)
               With newR
                   .ParentId = strID
                   .Name = strreqfol(WorkingDepth)
                   .Comment = ""
                   .Priority = ""
                   .Author = tmpReq.Author
                   .Reviewed = tmpReq.Status
                   .Field("RQ_USER_TEMPLATE_01") = tmpReq.Peer_Review
                   .TypeId = "Folder"
                   .Post
                   On Error GoTo 0
                   strID = newR.ID
               End With
               Set NewRequirement = newR
               Set NewRequirement = Nothing
           Else
               pos = InStr(strItem, ",")
               strID = Mid(strItem, 1, pos - 1)
               strName = Mid(strItem, pos + 1)
           End If
           
           ' Convert the ID to a long, and get the object
           reqID = CLng(strID)
    
           Set theReq = reqF.Item(reqID)
               'MsgBox theReq.Name
           thePath = theReq.path
    
           PathArray = Split(thePath, "\")
    
           Set ParentReq = theReq
           If UBound(PathArray) = WorkingDepth Then
           End If
           
           If ParentReq Is Nothing Then Exit For
       Next WorkingDepth
    End If
    Create_Requirement = True
           '*****CREATING REQUIREMENTS*****
    
    Set NewRequirement = newR
    Set NewRequirement = Nothing
       
       Set newR = reqF.AddItem(Null)
       strReqID = strID
       With newR
           .ParentId = strReqID
           .Name = tmpReq.Requirement_Name
           .Comment = ""
           .Priority = ""
           .Author = tmpReq.Author
           .Reviewed = tmpReq.Status
           .Field("RQ_USER_TEMPLATE_01") = tmpReq.Peer_Review
           .TypeId = tmpReq.Requirement_Type
           .Field("RQ_USER_TEMPLATE_02") = tmpReq.Additional_Info
           .Field("RQ_USER_01") = tmpReq.ExistR1
           .Field("RQ_USER_02") = tmpReq.ExistR2
           .Field("RQ_REQ_COMMENT") = tmpReq.WRIEF_Info
           On Error Resume Next
           .Post
           If Err.Number = "-2147219448" Then
               On Error GoTo 0
           End If
       End With
    
       Set NewRequirement = newR
       Set NewRequirement = Nothing
End Function
'########################### End Of Create New Test Plan BPT ###########################

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

Private Sub cmdLoadExcel_Click()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim fname As String
Dim lastrow
Dim i, j, tmpParam
Dim tmpSts
Dim strFunct As New clsFiles

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
         If UCase(Trim(.Range("A1").Value)) <> UCase(Trim("Folder Structure")) Then
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
            
            
            flxImport.TextMatrix(i - 1, 0) = strFunct.RemoveBackslash(Trim((.Range("A" & i).Value)))        'Change number and letter
            If Trim(.Range("A" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Folder Structure=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 1) = Trim((.Range("B" & i).Value))        'Change number and letter
            If Trim(.Range("B" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Requirement Type=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 2) = Trim((.Range("C" & i).Value))        'Change number and letter
            If Trim(.Range("C" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Requirement Name=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 3) = Trim((.Range("D" & i).Value))        'Change number and letter
'            If Trim(.Range("D" & i).Value) = "" Then
'                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Additional Info=BLANK]"
'                tmpSts = tmpSts + 1
'            End If
            
            flxImport.TextMatrix(i - 1, 4) = Trim((.Range("E" & i).Value))        'Change number and letter
            If Trim(.Range("E" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Author=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 5) = Trim((.Range("F" & i).Value))        'Change number and letter
            If Trim(.Range("F" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Peer Review=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 6) = Trim((.Range("G" & i).Value))        'Change number and letter
            If Trim(.Range("G" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Status=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 7) = Trim((.Range("H" & i).Value))        'Change number and letter
            If Trim(UCase(.Range("H" & i).Value)) <> "YES" And Trim(UCase(.Range("H" & i).Value)) <> "NO" And Trim(UCase(.Range("H" & i).Value)) <> "" Then
                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Exist in Release 1=Invalid Value]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 8) = Trim((.Range("I" & i).Value))        'Change number and letter
            If Trim(UCase(.Range("I" & i).Value)) <> "YES" And Trim(UCase(.Range("I" & i).Value)) <> "NO" And Trim(UCase(.Range("I" & i).Value)) <> "" Then
                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[Exist in Release 2=Invalid Value]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 9) = Trim((.Range("J" & i).Value))        'Change number and letter
'            If Trim(.Range("J" & i).Value) = "" Then
'                flxImport.TextMatrix(i - 1, 10) = flxImport.TextMatrix(i - 1, 10) & "[WRIEF Info=BLANK]"
'                tmpSts = tmpSts + 1
'            End If
            
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = i - 1 & " out of " & lastrow - 1 & " validated " & Format(i / lastrow, "0.0%") & " (" & tmpSts & ") errors found."
            mdiMain.pBar.Value = i
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
    End With
       
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
    If IncorrectHeaderDetails = False Then
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
        MsgBox "The template has an invalid/incorrect headers"
    End If
Else
    MsgBox "No items to be uploaded."
End If
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
flxImport.TextMatrix(0, 0) = "Folder Structure"
flxImport.TextMatrix(0, 1) = "Requirement Type"
flxImport.TextMatrix(0, 2) = "Requirement Name"
flxImport.TextMatrix(0, 3) = "Additional Info"
flxImport.TextMatrix(0, 4) = "Author"
flxImport.TextMatrix(0, 5) = "Peer Review"
flxImport.TextMatrix(0, 6) = "Status"
flxImport.TextMatrix(0, 7) = "Exist in Release 1"
flxImport.TextMatrix(0, 8) = "Exist in Release 2"
flxImport.TextMatrix(0, 9) = "WRIEF Info"
flxImport.TextMatrix(0, 10) = "Validation"
flxImport.Rows = 2
End Sub

Public Function IncorrectHeaderDetails() As Boolean
    If flxImport.TextMatrix(0, 0) <> "Folder Structure" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 1) <> "Requirement Type" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 2) <> "Requirement Name" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 3) <> "Additional Info" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 4) <> "Author" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 5) <> "Peer Review" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 6) <> "Status" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 7) <> "Exist in Release 1" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 8) <> "Exist in Release 2" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 9) <> "WRIEF Info" Then IncorrectHeaderDetails = True
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
  'xlObject.Visible = True
  curTab = "RQ_BPT-01"
  xlObject.Sheets("Sheet1").Name = curTab
  flxImport.FixedCols = 0
  flxImport.FixedRows = 0
  Pause 1
  flxImport.RowSel = flxImport.Rows - 1
  flxImport.ColSel = flxImport.Cols - 1
  Clipboard.Clear
  
  Clipboard.SetText flxImport.Clip
  flxImport.FixedCols = 1
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
    
    xlObject.Sheets(curTab).Range("K:K").Interior.ColorIndex = 3
    'xlObject.Sheets(curTab).Protection.AllowEditRanges.Add Title:="Range1", Range:=xlObject.Sheets(curTab).Range("A:J")
    'xlObject.Sheets(curTab).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  xlObject.Workbooks(1).SaveAs "RQ_BPT-01" & "-" & Format(Now, "mmddyyyy HHMMSS AMPM")
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

