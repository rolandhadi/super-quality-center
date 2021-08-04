VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLinkTestLabBPT 
   Caption         =   "Link Test Plan to Test Lab Module"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12645
   Icon            =   "frmLinkTestLabBPT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   12645
   Tag             =   "Link Test Plan to Test Lab Module"
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkOnlyID 
      Caption         =   "I already have all the ID and I don't need validations"
      Height          =   315
      Left            =   4920
      TabIndex        =   12
      Top             =   600
      Width           =   4035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12645
      _ExtentX        =   22304
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
               Picture         =   "frmLinkTestLabBPT.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLinkTestLabBPT.frx":0BAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLinkTestLabBPT.frx":10FD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLinkTestLabBPT.frx":164E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.StatusBar stsBar 
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   5520
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
               Picture         =   "frmLinkTestLabBPT.frx":1B9F
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Bevel           =   2
               Object.Width           =   21661
               MinWidth        =   17639
            EndProperty
         EndProperty
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
         Picture         =   "frmLinkTestLabBPT.frx":20F0
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
         Picture         =   "frmLinkTestLabBPT.frx":4914
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
            Picture         =   "frmLinkTestLabBPT.frx":50BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkTestLabBPT.frx":534C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkTestLabBPT.frx":55DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   7215
      Left            =   4920
      TabIndex        =   1
      Top             =   900
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   12726
      _Version        =   393216
      Cols            =   5
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
            Picture         =   "frmLinkTestLabBPT.frx":586C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkTestLabBPT.frx":5F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkTestLabBPT.frx":6690
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkTestLabBPT.frx":6DA2
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
   Begin MSComctlLib.TreeView QCTree_TL 
      Height          =   3375
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   5953
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
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
   Begin MSComctlLib.TreeView QCTree_TP 
      Height          =   3255
      Left            =   60
      TabIndex        =   5
      Top             =   4800
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   5741
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   8220
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmLinkTestLabBPT.frx":74B4
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   21079
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
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
            Picture         =   "frmLinkTestLabBPT.frx":7A05
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkTestLabBPT.frx":7CE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkTestLabBPT.frx":8238
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkTestLabBPT.frx":8789
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Test Plan Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Test Lab Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1875
   End
   Begin VB.Label lbl_TP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2340
      TabIndex        =   7
      Top             =   4380
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lbl_TL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2340
      TabIndex        =   6
      Top             =   540
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "frmLinkTestLabBPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Type TL_BPT
    Test_Plan_ID As String
    Test_Lab_ID As String
    Log As String
End Type

Private All_BPT() As TL_BPT
Private HasIssue As Boolean
Private HasUploadIssue  As Integer

Private Function LoadToArray()
Dim lastrow, i, EndArr
lastrow = flxImport.Rows - 1
ReDim All_BPT(0)
EndArr = -1
For i = 1 To lastrow
    If Trim(flxImport.TextMatrix(i, 0)) = "" Or Trim(flxImport.TextMatrix(i, 1)) = "" Then
        All_BPT(EndArr).Log = All_BPT(EndArr).Log & vbCrLf & "Line " & i & " is blank"
    Else
        EndArr = EndArr + 1
        ReDim Preserve All_BPT(EndArr)
        All_BPT(EndArr).Test_Lab_ID = flxImport.TextMatrix(i, 0)
        All_BPT(EndArr).Test_Plan_ID = flxImport.TextMatrix(i, 1)
        If Trim(flxImport.TextMatrix(i, flxImport.Cols - 1)) <> "" Then All_BPT(EndArr).Log = "ISSUE"
    End If
Next
End Function

Sub Start()
Debug.Print "New Session: " & Now
LoadToArray
LoadToQC
Debug.Print "New Finished: " & Now
End Sub

Function LoadToQC()
Dim i, j
Dim tmpComp
stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = ""
mdiMain.pBar.Max = UBound(All_BPT) + 3
For i = LBound(All_BPT) To UBound(All_BPT)
    On Error Resume Next
    If All_BPT(i).Log = "" Then
        Call LinkBPTTest(All_BPT(i))
        stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Linking Test Plan - Component " & i + 1 & " out of " & UBound(All_BPT) + 1 & " (" & All_BPT(i).Test_Lab_ID & "-" & All_BPT(i).Test_Plan_ID & ")"
        If Err.Number <> 0 Then
            FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[LINK BC: (FAILED) " & Now & " " & All_BPT(i).Test_Lab_ID & "-" & All_BPT(i).Test_Plan_ID & "] " & Err.Description
            HasUploadIssue = HasUploadIssue + 1
        Else
            FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[LINK BC: (PASSED) " & Now & " " & All_BPT(i).Test_Lab_ID & "-" & All_BPT(i).Test_Plan_ID & "]"
        End If
    Else
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[LINK BC: (SKIPPED) " & Now & " " & All_BPT(i).Test_Lab_ID & "-" & All_BPT(i).Test_Plan_ID & "]"
    End If
    Err.Clear
    On Error GoTo 0
    mdiMain.pBar.Value = i + 1
        If mdiMain.pBar.Max > 50 Then
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
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = UBound(All_BPT) + 1 & " Test Lab - Test Plan linked successfully. (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log"
    QCConnection.SendMail "user@companyemail.com", "", "[HPQC UPDATES] Test Lab - Test Plan linked successfully by " & curUser & " in " & curDomain & "-" & curProject, UBound(All_BPT) + 1 & " Test Lab - Test Plan linked successfully. (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
    QCConnection.SendMail curUser, "", "[HPQC UPDATES] Test Lab - Test Plan linked successfully by " & curUser & " in " & curDomain & "-" & curProject, UBound(All_BPT) + 1 & " Test Lab - Test Plan linked successfully. (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
    If HasUploadIssue <> 0 Then
      Dim tmpFile As New clsFiles
      frmLogs.Caption = App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log"
      frmLogs.txtLogs.Text = tmpFile.ReadFromFile_FAILED(App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log")
      frmLogs.Show 1
    End If
End Function

'########################### Link New Test Plan Folder ###########################
Private Sub LinkBPTTest(tmp_BPT As TL_BPT)
Dim tfact As TestFactory
Dim mytest As Test
Dim TestSetFact As TestSetFactory
Dim mytestset As TestSet
Dim tsttestsetFact As TSTestFactory


'Get the Test Factory
Set TestSetFact = QCConnection.TestSetFactory
Set mytestset = TestSetFact.Item(tmp_BPT.Test_Lab_ID)

Set tfact = QCConnection.TestFactory
Set mytest = tfact.Item(tmp_BPT.Test_Plan_ID)
 
Set tsttestsetFact = mytestset.TSTestFactory

tsttestsetFact.AddItem (mytest.ID)

Set tfact = Nothing
Set mytest = Nothing
End Sub
'########################### End Of Link New Test Plan Folder ###########################

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

Private Sub chkOnlyID_Click()
If chkOnlyID.Value = Checked Then
    If MsgBox("Are you sure you want to disable validation?", vbYesNo) = vbYes Then
        chkOnlyID.Value = Checked
    Else
        chkOnlyID.Value = Unchecked
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
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim strPath

HasIssue = False

On Error Resume Next
strPath = QCTree_TL.SelectedItem.Key
strPath = QCTree_TP.SelectedItem.Key
strPath = ""
If Err.Number = 91 Then
    MsgBox "Please select a folder in Test Plan and in the Test Lab tree folder"
    Exit Sub
End If
On Error GoTo 0
On Error GoTo ErrLoad
On Error Resume Next
    xlWB.Close
    xlObject.Application.Quit
On Error GoTo 0
    dlgOpenExcel.filename = "": dlgOpenExcel.ShowOpen
    fname = dlgOpenExcel.filename
    If fname = "" Then Exit Sub Else Me.Caption = Me.Caption & " (" & dlgOpenExcel.FileTitle & ")"
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Open(fname) 'Open your book here
                
    Clipboard.Clear

    With xlObject.ActiveWorkbook.ActiveSheet
         If UCase(Trim(.Range("A1").Value)) <> UCase(Trim("Test Set ID")) Then
            MsgBox "Import file is invalid. Please use only sheets generated by the SuperQC"
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
         lastrow = .Range("C" & .Rows.Count).End(xlUp).row
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
        If chkOnlyID.Value = Checked Then
            For i = 2 To lastrow
                flxImport.TextMatrix(i - 1, 0) = CleanTheString(Trim((.Range("A" & i).Value)))         'Change number and letter
                If Trim(.Range("A" & i).Value) = "" Then
                    flxImport.TextMatrix(i - 1, 4) = flxImport.TextMatrix(i - 1, 4) & "[Test Set ID=BLANK]"
                    tmpSts = tmpSts + 1
                End If
                flxImport.TextMatrix(i - 1, 1) = CleanTheString(Trim((.Range("B" & i).Value)))         'Change number and letter
                If Trim(.Range("B" & i).Value) = "" Then
                    flxImport.TextMatrix(i - 1, 4) = flxImport.TextMatrix(i - 1, 4) & "[Test ID=BLANK]"
                    tmpSts = tmpSts + 1
                End If
            Next
        Else
            For i = 2 To lastrow
                flxImport.TextMatrix(i - 1, 2) = CleanTheString(Trim((.Range("C" & i).Value)))        'Change number and letter
                If Trim(.Range("C" & i).Value) = "" Then
                    flxImport.TextMatrix(i - 1, 4) = flxImport.TextMatrix(i - 1, 4) & "[Test Set Name=BLANK]"
                    tmpSts = tmpSts + 1
                Else
                            ReDim CheckedItems(0): strPath = ""
                            GetAllCheckedItems QCTree_TL.Nodes(1)
                            For j = LBound(CheckedItems) To UBound(CheckedItems) - 1
                                strPath = strPath & "CF_ITEM_PATH LIKE '" & GetFromTable(Right(CheckedItems(j), Len(CheckedItems(j)) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%'" & " OR "
                            Next
                            If Trim(strPath) <> "" Then
                                strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
                            Else
                                MsgBox "Please select and check source(s) in the HPQC folder tree"
                                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                                Exit Sub
                            End If
                        'strPath = "'" & GetFromTable(Right(QCTree_TL.SelectedItem.Key, Len(QCTree_TL.SelectedItem.Key) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%'"
                        Set objCommand = QCConnection.Command
                        objCommand.CommandText = "SELECT CY_CYCLE_ID FROM CYCLE, CYCL_FOLD WHERE CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " AND CY_CYCLE = '" & Trim(.Range("C" & i).Value) & "'"
                        Debug.Print Me.Caption & "-" & objCommand.CommandText
                    Set rs = objCommand.Execute
                    If rs.RecordCount = 1 Then
                        flxImport.TextMatrix(i - 1, 0) = rs.FieldValue("CY_CYCLE_ID")
                    ElseIf rs.RecordCount > 1 Then
                        flxImport.TextMatrix(i - 1, 0) = rs.FieldValue("CY_CYCLE_ID")
                        flxImport.TextMatrix(i - 1, 4) = flxImport.TextMatrix(i - 1, 4) & "[MULTIPLE TEST (" & rs.RecordCount & "]"
                        tmpSts = tmpSts + 1
                    Else
                        flxImport.TextMatrix(i - 1, 4) = flxImport.TextMatrix(i - 1, 4) & "[TEST N/A]"
                        tmpSts = tmpSts + 1
                    End If
                End If
                
                flxImport.TextMatrix(i - 1, 3) = CleanTheString(Trim((.Range("D" & i).Value)))        'Change number and letter
                If Trim(.Range("D" & i).Value) = "" Then
                    flxImport.TextMatrix(i - 1, 4) = flxImport.TextMatrix(i - 1, 4) & "[Test Name=BLANK]"
                    tmpSts = tmpSts + 1
                Else
                            ReDim CheckedItems(0): strPath = ""
                            GetAllCheckedItems QCTree_TP.Nodes(1)
                            For j = LBound(CheckedItems) To UBound(CheckedItems) - 1
                                strPath = strPath & "AL_ABSOLUTE_PATH LIKE '" & GetFromTable(Right(CheckedItems(j), Len(CheckedItems(j)) - 1), "AL_ITEM_ID", "AL_ABSOLUTE_PATH", "ALL_LISTS") & "%'" & " OR "
                            Next
                            If Trim(strPath) <> "" Then
                                strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
                            Else
                                MsgBox "Please select and check source(s) in the HPQC folder tree"
                                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                                Exit Sub
                            End If
                    'strPath = "'" & GetFromTable(Right(QCTree_TP.SelectedItem.Key, Len(QCTree_TP.SelectedItem.Key) - 1), "AL_ITEM_ID", "AL_ABSOLUTE_PATH", "ALL_LISTS") & "%'"
                    Set objCommand = QCConnection.Command
                    objCommand.CommandText = "SELECT TS_TEST_ID FROM TEST, ALL_LISTS WHERE TS_SUBJECT = AL_ITEM_ID AND " & strPath & " AND TS_NAME = '" & Trim(.Range("D" & i).Value) & "'"
                    Debug.Print Me.Caption & "-" & objCommand.CommandText
                    Set rs = objCommand.Execute
                    If rs.RecordCount = 1 Then
                        flxImport.TextMatrix(i - 1, 1) = rs.FieldValue("TS_TEST_ID")
                    ElseIf rs.RecordCount > 1 Then
                        flxImport.TextMatrix(i - 1, 1) = rs.FieldValue("TS_TEST_ID")
                        flxImport.TextMatrix(i - 1, 4) = flxImport.TextMatrix(i - 1, 4) & "[MULTIPLE TEST]"
                        tmpSts = tmpSts + 1
                    Else
                        flxImport.TextMatrix(i - 1, 4) = flxImport.TextMatrix(i - 1, 4) & "[TEST N/A]"
                        tmpSts = tmpSts + 1
                    End If
                End If
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = i - 1 & " out of " & lastrow - 1 & " validated " & Format(i / lastrow, "0.0%") & " (" & tmpSts & ") errors found."
                mdiMain.pBar.Value = i
                If mdiMain.pBar.Max > 50 Then
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
        End If
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

Private Sub Label1_Click()
Dim tmpPath, tmpID: On Error Resume Next
tmpID = Right(QCTree_TL.SelectedItem.Key, Len(QCTree_TL.SelectedItem.Key) - 1)
tmpPath = GetFromTable(Right(QCTree_TL.SelectedItem.Key, Len(QCTree_TL.SelectedItem.Key) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%"
frmLogs.txtLogs.Text = "Test Set ID: " & tmpID & vbCrLf & "CF_ITEM_PATH: " & tmpPath & vbCrLf & "Folder Path: " & QCTree_TL.SelectedItem.FullPath
frmLogs.Show 1
End Sub

Private Sub Label2_Click()
Dim tmpPath, tmpID: On Error Resume Next
tmpID = Right(QCTree_TP.SelectedItem.Key, Len(QCTree_TP.SelectedItem.Key) - 1)
tmpPath = GetFromTable(Right(QCTree_TP.SelectedItem.Key, Len(QCTree_TP.SelectedItem.Key) - 1), "AL_ITEM_ID", "AL_ABSOLUTE_PATH", "ALL_LISTS") & "%"
frmLogs.txtLogs.Text = "Test ID: " & tmpID & vbCrLf & "AL_ABSOLUTE_PATH: " & tmpPath & vbCrLf & "Folder Path: " & QCTree_TP.SelectedItem.FullPath
frmLogs.Show 1
End Sub

Private Sub QCTree_TL_DblClick()
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
Dim nodx As Node

    If QCTree_TL.SelectedItem.Children <> 0 Then Exit Sub
    
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = " & Right(QCTree_TL.SelectedItem.Key, Len(QCTree_TL.SelectedItem.Key) - 1) & " ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree_TL.Nodes.Add QCTree_TL.SelectedItem.Key, tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    lbl_TL.Caption = QCTree_TL.SelectedItem.Text
    ClearTable
End Sub

Private Sub QCTree_TL_NodeCheck(ByVal Node As MSComctlLib.Node)
Node.Selected = True
End Sub

Private Sub QCTree_TP_DblClick()
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
Dim nodx As Node

    If QCTree_TP.SelectedItem.Children <> 0 Then Exit Sub
    
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT AL_ITEM_ID, AL_DESCRIPTION FROM ALL_LISTS WHERE AL_FATHER_ID = " & Right(QCTree_TP.SelectedItem.Key, Len(QCTree_TP.SelectedItem.Key) - 1) & " ORDER BY AL_DESCRIPTION"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree_TP.Nodes.Add QCTree_TP.SelectedItem.Key, tvwChild, CStr("F" & rs.FieldValue("AL_ITEM_ID")), rs.FieldValue("AL_DESCRIPTION"), 1
        rs.Next
    Next
    lbl_TP.Caption = QCTree_TP.SelectedItem.Text
    ClearTable
End Sub

Private Sub QCTree_TP_NodeCheck(ByVal Node As MSComctlLib.Node)
Node.Selected = True
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
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
QCTree_TL.Nodes.Clear
QCTree_TP.Nodes.Clear

    QCTree_TL.Nodes.Add , , "Root", "Root", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = 0 ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree_TL.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next

    QCTree_TP.Nodes.Add , , "Root", "Subject", 1
    objCommand.CommandText = "SELECT AL_ITEM_ID, AL_DESCRIPTION FROM ALL_LISTS WHERE AL_FATHER_ID = 2 ORDER BY AL_DESCRIPTION"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree_TP.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("AL_ITEM_ID")), rs.FieldValue("AL_DESCRIPTION"), 1
        rs.Next
    Next
    
    lbl_TL.Caption = ""
    lbl_TP.Caption = ""
     Me.Caption = Me.Tag
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.Cols = 5
flxImport.TextMatrix(0, 0) = "Test Set ID"
flxImport.TextMatrix(0, 1) = "Test ID"
flxImport.TextMatrix(0, 2) = "Test Set Name (TEST LAB)"
flxImport.TextMatrix(0, 3) = "Test Name (TEST PLAN)"
flxImport.TextMatrix(0, 4) = "Validation"
flxImport.Rows = 2
End Sub

Public Function IncorrectHeaderDetails() As Boolean
    If flxImport.TextMatrix(0, 0) <> "Test Set ID" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 1) <> "Test ID" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 2) <> "Test Set Name (TEST LAB)" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 3) <> "Test Name (TEST PLAN)" Then IncorrectHeaderDetails = True
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
  'xlObject.Visible = True
  curTab = "TL_LINKTP-01"
  xlObject.Sheets("Sheet1").Name = curTab
  flxImport.FixedCols = 0
  flxImport.FixedRows = 0
  flxImport.row = 0
  flxImport.col = 0
  Pause 1
  flxImport.RowSel = flxImport.Rows - 1
  flxImport.ColSel = flxImport.Cols - 1
  Clipboard.Clear
'  For i = 1 To 5
'    flxImport.Clip = Replace(flxImport.Clip, vbCrLf, "<br>", 1, , vbTextCompare)
'    flxImport.Clip = Replace(flxImport.Clip, vbNewLine, "<br>", 1, , vbTextCompare)
'    flxImport.Clip = Replace(flxImport.Clip, Chr(10) & Chr(13), "<br>", 1, , vbTextCompare)
'    flxImport.Clip = Replace(flxImport.Clip, Chr(10), "<br>", 1, , vbTextCompare)
'    flxImport.Clip = Replace(flxImport.Clip, Chr(13), "<br>", 1, , vbTextCompare)
'    flxImport.Clip = Replace(flxImport.Clip, vbCr, "<br>", 1, , vbTextCompare)
'    flxImport.Clip = Replace(flxImport.Clip, vbLf, "<br>", 1, , vbTextCompare)
'  Next
  Clipboard.SetText flxImport.Clip
  flxImport.FixedCols = 1
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

    xlObject.Sheets(curTab).Range("A:B").Interior.ColorIndex = 3
    xlObject.Sheets(curTab).Range("E:E").Interior.ColorIndex = 3
    'xlObject.Sheets(curTab).Protection.AllowEditRanges.Add Title:="Range1", Range:=xlObject.Sheets(curTab).Range("C:D")
    'xlObject.Sheets(curTab).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
  xlObject.Workbooks(1).SaveAs "TL_LINKTP-01" & "-" & Format(Now, "mmddyyyy HHMMSS AMPM")
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

