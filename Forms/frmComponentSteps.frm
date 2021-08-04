VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComponentSteps 
   Caption         =   "Update Component Steps"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12630
   Icon            =   "frmComponentSteps.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   12630
   WindowState     =   2  'Maximized
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
      Left            =   4560
      Picture         =   "frmComponentSteps.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Import step description and expected results from an excel file"
      Top             =   600
      Width           =   2205
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12630
      _ExtentX        =   22278
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
            Key             =   "cmdUpload"
            Object.ToolTipText     =   "Upload to HPQC"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdOutput"
            Object.ToolTipText     =   "Export to Excel"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   30
         Left            =   1080
         TabIndex        =   4
         Top             =   300
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   53
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6060
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
            Picture         =   "frmComponentSteps.frx":1070
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponentSteps.frx":1302
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponentSteps.frx":1594
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView QCTree 
      Height          =   5355
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9446
      _Version        =   393217
      HideSelection   =   0   'False
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
   End
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   4935
      Left            =   4560
      TabIndex        =   3
      Top             =   1020
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   8
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
            Picture         =   "frmComponentSteps.frx":1822
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponentSteps.frx":1F34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponentSteps.frx":2646
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponentSteps.frx":2D58
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
End
Attribute VB_Name = "frmComponentSteps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BC_Component
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
    Step_Order() As Integer
    Step_Name() As String
    Step_Description() As String
    Step_ExpectedResult() As String
    Log As String
    StepGroup As Integer
End Type

Private All_BC() As BC_Component
Private HasIssue As Boolean
Private HasUploadIssue  As Integer

Private Function LoadToArray()
Dim LastRow, i, LastVal, EndArr
LastRow = flxImport.Rows - 1
ReDim All_BC(0)
EndArr = -1
LastVal = 0
For i = 1 To LastRow
    If Trim(flxImport.TextMatrix(i, 0)) = "" Or Trim(flxImport.TextMatrix(i, 1)) = "" Then
        All_BC(EndArr).Log = All_BC(EndArr).Log & vbCrLf & "Line " & i & " is blank"
    ElseIf (Trim(UCase(flxImport.TextMatrix(i, 0))) <> Trim(UCase(flxImport.TextMatrix(i - 1, 0)))) Or (Trim(UCase(flxImport.TextMatrix(i, 1))) <> Trim(UCase(flxImport.TextMatrix(i - 1, 1)))) Then
        EndArr = EndArr + 1
        ReDim Preserve All_BC(EndArr)
        LastVal = LastVal + 1
        All_BC(EndArr).Group_No = LastVal
        All_BC(EndArr).Component_Path = flxImport.TextMatrix(i, 0)
        All_BC(EndArr).Component_Name = flxImport.TextMatrix(i, 1)
        All_BC(EndArr).Component_Description = "" 'Range("E" & i).Value
        All_BC(EndArr).Scripter = flxImport.TextMatrix(i, 8)
        All_BC(EndArr).Peer_Reviewer = flxImport.TextMatrix(i, 9)
        All_BC(EndArr).QA_Reviewer = flxImport.TextMatrix(i, 3)
        All_BC(EndArr).Planned_Start_Date = flxImport.TextMatrix(i, 5)
        All_BC(EndArr).Planned_End_Date = flxImport.TextMatrix(i, 6)
        All_BC(EndArr).Status = flxImport.TextMatrix(i, 2)
        
        ReDim All_BC(EndArr).Step_Order(0)
        ReDim All_BC(EndArr).Step_Name(0)
        ReDim All_BC(EndArr).Step_Description(0)
        ReDim All_BC(EndArr).Step_ExpectedResult(0)
        
        All_BC(EndArr).Step_Order(0) = 1
        All_BC(EndArr).Step_Name(0) = "Step " & UBound(All_BC(EndArr).Step_Name) + 1
        All_BC(EndArr).Step_Description(0) = flxImport.TextMatrix(i, 11)
        All_BC(EndArr).Step_ExpectedResult(0) = flxImport.TextMatrix(i, 12)
    Else
        ReDim Preserve All_BC(EndArr).Step_Order(UBound(All_BC(EndArr).Step_Order) + 1)
        ReDim Preserve All_BC(EndArr).Step_Name(UBound(All_BC(EndArr).Step_Name) + 1)
        ReDim Preserve All_BC(EndArr).Step_Description(UBound(All_BC(EndArr).Step_Description) + 1)
        ReDim Preserve All_BC(EndArr).Step_ExpectedResult(UBound(All_BC(EndArr).Step_ExpectedResult) + 1)
        
        All_BC(EndArr).Step_Order(UBound(All_BC(EndArr).Step_Order)) = UBound(All_BC(EndArr).Step_Order) + 1
        All_BC(EndArr).Step_Name(UBound(All_BC(EndArr).Step_Name)) = "Step " & UBound(All_BC(EndArr).Step_Name) + 1
        All_BC(EndArr).Step_Description(UBound(All_BC(EndArr).Step_Description)) = flxImport.TextMatrix(i, 11)
        All_BC(EndArr).Step_ExpectedResult(UBound(All_BC(EndArr).Step_ExpectedResult)) = flxImport.TextMatrix(i, 12)
    End If
Next
End Function

Function LoadToQC()
Dim i, j
Dim tmpComp
stsBar.SimpleText = ""
For i = LBound(All_BC) To UBound(All_BC)
    On Error Resume Next
        'Set tmpComp = Create_New_Component(All_BC(i))
    If Err.Number = 0 Then
        Err.Clear
        On Error GoTo 0
        On Error Resume Next
        'Add_Params_And_Steps tmpComp, All_BC(i).Step_Order, All_BC(i).Step_Name, All_BC(i).Step_Description, All_BC(i).Step_ExpectedResult
    Else
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[CREATE BC: " & Now & " " & All_BC(i).Component_Path & "-" & All_BC(i).Component_Name & "] " & Err.Description
        HasUploadIssue = HasUploadIssue + 1
        Err.Clear
        On Error GoTo 0
    End If
    stsBar.SimpleText = "Loading Business Component " & i + 1 & " out of " & UBound(All_BC) + 1 & " (" & All_BC(i).Component_Name & ")"
    Err.Clear
    On Error GoTo 0
Next
stsBar.SimpleText = UBound(All_BC) + 1 & " Business Component(s) loaded successfully. (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log"
End Function

Sub Start()
Debug.Print "New Session: " & Now
LoadToArray
LoadToQC
Debug.Print "New Finished: " & Now
End Sub

Private Sub GenerateOutput()
Dim rs As TDAPIOLELib.Recordset
Dim AllScript
Dim objCommand
Dim i
Dim strPath
Dim iterd
Dim NewVal
Dim iter
Dim TimeSt
Dim AllF
Dim Decompile
Dim p
Dim k
Dim z
Dim mySTEP_ID
Dim myFOLDER_NAME
Dim myCOMPONENT_NAME
Dim mySTEP_ORDER
Dim mySTEP_NAME
Dim myDESCRIPTION
Dim myEXPECTED

    TimeSt = Format(Now, "mmddyyyy HHMM AMPM") & "-"
    AllF = "Business Component Steps"

    If Left(QCTree.SelectedItem.Key, 1) = "F" Then
        strPath = "FC_PATH LIKE '" & GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "FC_ID", "FC_PATH", "COMPONENT_FOLDER") & "%'"
    ElseIf Left(QCTree.SelectedItem.Key, 1) = "C" Then
        strPath = "CO_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
    End If

    Set objCommand = QCConnection.Command

    objCommand.CommandText = "SELECT CS_STEP_ID, FC_NAME, CO_NAME, CS_STEP_ORDER, CS_STEP_NAME, CS_DESCRIPTION, CS_EXPECTED FROM COMPONENT_STEP, COMPONENT, COMPONENT_FOLDER WHERE CO_ID = CS_COMPONENT_ID AND CO_FOLDER_ID = FC_ID AND " & strPath & " ORDER BY FC_NAME, CO_ID, CS_STEP_ORDER"

    Set rs = objCommand.Execute                                                                                                                                                                                                                                                 'HERE!!!!!! <<<-------------
    AllScript = "Step ID" & vbTab & "Test Set Folder" & vbTab & "Component Name" & vbTab & "Step Order" & vbTab & "Step Name" & vbTab & "Description" & vbTab & "Expected Result"
    FileWrite App.path & "\SQC DAT" & "\" & TimeSt & AllF & ".xls", AllScript
    AllScript = ""
    ClearTable
    flxImport.Rows = rs.RecordCount + 1
    k = 0

    For i = 1 To rs.RecordCount
        If i = 1 Then FileWrite App.path & "\SQC DAT" & "\" & QCTree.SelectedItem.Key & "-" & TimeSt & AllF & ".xls", AllScript
                stsBar.SimpleText = "Processing " & i & " out of " & rs.RecordCount
                
                mySTEP_ID = rs.FieldValue("CS_STEP_ID")
        
                myFOLDER_NAME = rs.FieldValue("FC_NAME")
                
                myCOMPONENT_NAME = rs.FieldValue("CO_NAME")
                
                mySTEP_ORDER = rs.FieldValue("CS_STEP_ORDER")
                
                mySTEP_NAME = rs.FieldValue("CS_STEP_NAME")
                mySTEP_NAME = Replace(mySTEP_NAME, "&amp;", "&", , , vbTextCompare)
                mySTEP_NAME = Replace(mySTEP_NAME, "&apos;", "'", , , vbTextCompare)
                mySTEP_NAME = Replace(mySTEP_NAME, "&lt;", "<", , , vbTextCompare)
                mySTEP_NAME = Replace(mySTEP_NAME, "&gt;", ">", , , vbTextCompare)
                mySTEP_NAME = Replace(mySTEP_NAME, "&quot;", """", , , vbTextCompare)
                mySTEP_NAME = Replace(mySTEP_NAME, "<html><body>", "", , , vbTextCompare)
                mySTEP_NAME = Replace(mySTEP_NAME, "</body></html>", "", , , vbTextCompare)
                
                myDESCRIPTION = rs.FieldValue("CS_DESCRIPTION")
                myDESCRIPTION = Replace(myDESCRIPTION, "&amp;", "&", , , vbTextCompare)
                myDESCRIPTION = Replace(myDESCRIPTION, "&apos;", "'", , , vbTextCompare)
                myDESCRIPTION = Replace(myDESCRIPTION, "&lt;", "<", , , vbTextCompare)
                myDESCRIPTION = Replace(myDESCRIPTION, "&gt;", ">", , , vbTextCompare)
                myDESCRIPTION = Replace(myDESCRIPTION, "&quot;", """", , , vbTextCompare)
                myDESCRIPTION = Replace(myDESCRIPTION, "<html><body>", "", , , vbTextCompare)
                myDESCRIPTION = Replace(myDESCRIPTION, "</body></html>", "", , , vbTextCompare)
                myDESCRIPTION = Replace(myDESCRIPTION, "apos;", "'", , , vbTextCompare)
                
                For z = 1 To 100
                    myDESCRIPTION = Replace(myDESCRIPTION, vbCrLf, "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 100
                    myDESCRIPTION = Replace(myDESCRIPTION, Chr(10) & Chr(13), "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 100
                    myDESCRIPTION = Replace(myDESCRIPTION, Chr(10), "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 100
                    myDESCRIPTION = Replace(myDESCRIPTION, Chr(13), "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 100
                    myDESCRIPTION = Replace(myDESCRIPTION, vbTab, "     ", , , vbTextCompare)
                Next
                
                myEXPECTED = rs.FieldValue("CS_EXPECTED")
                myEXPECTED = Replace(myEXPECTED, "&amp;", "&", , , vbTextCompare)
                myEXPECTED = Replace(myEXPECTED, "&apos;", "'", , , vbTextCompare)
                myEXPECTED = Replace(myEXPECTED, "&lt;", "<", , , vbTextCompare)
                myEXPECTED = Replace(myEXPECTED, "&gt;", ">", , , vbTextCompare)
                myEXPECTED = Replace(myEXPECTED, "&quot;", """", , , vbTextCompare)
                myEXPECTED = Replace(myEXPECTED, "<html><body>", "", , , vbTextCompare)
                myEXPECTED = Replace(myEXPECTED, "</body></html>", "", , , vbTextCompare)
                myEXPECTED = Replace(myEXPECTED, "apos;", "'", , , vbTextCompare)
                
                For z = 1 To 100
                    myEXPECTED = Replace(myEXPECTED, vbCrLf, "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 100
                    myEXPECTED = Replace(myEXPECTED, Chr(10) & Chr(13), "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 100
                    myEXPECTED = Replace(myEXPECTED, Chr(10), "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 100
                    myEXPECTED = Replace(myEXPECTED, Chr(13), "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 100
                    myEXPECTED = Replace(myEXPECTED, vbTab, "     ", , , vbTextCompare)
                Next
                        
                AllScript = AllScript & vbCrLf & mySTEP_ID & vbTab & myFOLDER_NAME & vbTab & myCOMPONENT_NAME & vbTab & mySTEP_ORDER & vbTab & mySTEP_NAME & vbTab & myDESCRIPTION & vbTab & myEXPECTED
                k = k + 1
                flxImport.Rows = k + 1
                flxImport.TextMatrix(k, 0) = mySTEP_ID
                flxImport.TextMatrix(k, 1) = myFOLDER_NAME
                flxImport.TextMatrix(k, 2) = myCOMPONENT_NAME
                flxImport.TextMatrix(k, 3) = mySTEP_ORDER
                flxImport.TextMatrix(k, 4) = mySTEP_NAME
                flxImport.TextMatrix(k, 5) = myDESCRIPTION
                flxImport.TextMatrix(k, 6) = myEXPECTED
        If i Mod 2500 = 0 Then
            FileAppend App.path & "\SQC DAT" & "\" & QCTree.SelectedItem.Key & "-" & TimeSt & AllF & ".xls", AllScript
            AllScript = ""
        End If
        rs.Next
    Next
FileAppend App.path & "\SQC DAT" & "\" & TimeSt & AllF & ".xls", AllScript
stsBar.SimpleText = "Ready"
End Sub


Private Sub cmdLoadExcel_Click()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim fname As String
Dim LastRow
Dim i, j, tmpParam
Dim tmpSts
Dim strFunct As New clsFiles
HasIssue = False

On Error Resume Next
    xlWB.Close
    xlObject.Application.Quit
On Error GoTo 0
    dlgOpenExcel.ShowOpen
    fname = dlgOpenExcel.filename
    If fname = "" Then Exit Sub
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Open(fname) 'Open your book here
                
    Clipboard.Clear

    With xlObject.ActiveWorkbook.ActiveSheet
         If UCase(Trim(.Range("A1").Value)) <> UCase(Trim("Step ID")) Then
            MsgBox "Import file is invalid. Please use only sheets generated by the Super QC"
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
         LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
        '.Range("A3:M" & LastRow).Copy 'Set selection to Copy
        
        ClearTable
        flxImport.Redraw = False     'Dont draw until the end, so we avoid that flash
        flxImport.Row = 0            'Paste from first cell
        flxImport.Col = 0
        flxImport.Rows = LastRow
        flxImport.Cols = 8
        flxImport.Redraw = False
        
        'A - Load HPQC Folder Path
        'Should not be blank
        For i = 2 To LastRow
            
            flxImport.TextMatrix(i - 1, 0) = strFunct.RemoveBackslash(((.Range("A" & i).Value)))          'Change number and letter
            If Trim(.Range("A" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step ID=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 1) = strFunct.RemoveBackslash(((.Range("B" & i).Value)))          'Change number and letter
            If Trim(.Range("B" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Test Set Folder=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 2) = strFunct.RemoveBackslash(((.Range("C" & i).Value)))          'Change number and letter
            If Trim(.Range("C" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Component Name=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 3) = strFunct.RemoveBackslash(((.Range("D" & i).Value)))          'Change number and letter
            flxImport.TextMatrix(i - 1, 4) = strFunct.RemoveBackslash(((.Range("E" & i).Value)))          'Change number and letter
            
            flxImport.TextMatrix(i - 1, 5) = Trim(.Range("F" & i).Value)    'Change number and letter
            If Trim(.Range("F" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step Description=BLANK]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, .Range("F" & i).Value, "<<<<", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step Description=<<<< FOUND]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, .Range("F" & i).Value, ">>>>", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step Description=>>>> FOUND]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, UCase(.Range("F" & i).Value), " <<P_", vbTextCompare) <> 0 Or InStr(1, UCase(.Range("F" & i).Value), vbCrLf & "<<P_", vbTextCompare) <> 0 Or InStr(1, UCase(.Range("F" & i).Value), " <<O_", vbTextCompare) <> 0 Or InStr(1, UCase(.Range("F" & i).Value), vbCrLf & "<<O_", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step Description=<< FOUND]"
                tmpSts = tmpSts + 1
            End If
            For j = 1 To 26
                If InStr(1, .Range("F" & i).Value, "<<<" & Chr(j + 64), vbTextCompare) <> 0 And (LCase(Chr(j + 64)) <> "p" And LCase(Chr(j + 64)) <> "o") Then
                    flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step Description=PARAMETER FORMAT FAIL]"
                    tmpSts = tmpSts + 1
                End If
            Next
            If HasParameters(.Range("F" & i).Value) = True Then
                tmpParam = ExtractParameters(.Range("F" & i).Value)
                For j = LBound(tmpParam) To UBound(tmpParam)
                    If InvalidParameterCheck(tmpParam(j)) = True Then
                        flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Parameter=INVALID FORMAT/CHAR]"
                        tmpSts = tmpSts + 1
                    End If
                Next
            End If
            If InStr(1, .Range("F" & i).Value, "<<< ", vbTextCompare) Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step Description=PARAMETER FORMAT FAIL]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, .Range("F" & i).Value, " >>>", vbTextCompare) Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step Description=PARAMETER FORMAT FAIL]"
                tmpSts = tmpSts + 1
            End If
            flxImport.TextMatrix(i - 1, 6) = Trim(.Range("G" & i).Value)    'Change number and letter
            If InStr(1, .Range("G" & i).Value, "<<<<", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step Description=<<<< FOUND]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, .Range("G" & i).Value, ">>>>", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step Description=>>>> FOUND]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, UCase(.Range("G" & i).Value), " <<P_", vbTextCompare) <> 0 Or InStr(1, UCase(.Range("G" & i).Value), vbCrLf & "<<P_", vbTextCompare) <> 0 Or InStr(1, UCase(.Range("G" & i).Value), " <<O_", vbTextCompare) <> 0 Or InStr(1, UCase(.Range("G" & i).Value), vbCrLf & "<<O_", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step Description=<< FOUND]"
                tmpSts = tmpSts + 1
            End If
            For j = 1 To 26
                If InStr(1, .Range("G" & i).Value, "<<<" & Chr(j + 64), vbTextCompare) <> 0 And (LCase(Chr(j + 64)) <> "p" And LCase(Chr(j + 64)) <> "o") Then
                    flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step Description=PARAMETER FORMAT FAIL]"
                    tmpSts = tmpSts + 1
                End If
            Next
            If HasParameters(.Range("G" & i).Value) = True Then
                tmpParam = ExtractParameters(.Range("G" & i).Value)
                For j = LBound(tmpParam) To UBound(tmpParam)
                    If InvalidParameterCheck(tmpParam(j)) = True Then
                        flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Parameter=INVALID FORMAT/CHAR]"
                        tmpSts = tmpSts + 1
                    End If
                Next
            End If
            If InStr(1, .Range("G" & i).Value, "<<< ", vbTextCompare) Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step Description=PARAMETER FORMAT FAIL]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, .Range("G" & i).Value, " >>>", vbTextCompare) Then
                flxImport.TextMatrix(i - 1, 7) = flxImport.TextMatrix(i - 1, 7) & "[Step Description=PARAMETER FORMAT FAIL]"
                tmpSts = tmpSts + 1
            End If
            stsBar.SimpleText = i - 1 & " out of " & LastRow - 1 & " validated " & Format(i / LastRow, "0.0%") & " (" & tmpSts & ") errors found."
        Next
    End With
    
    If Trim(flxImport.TextMatrix(flxImport.Rows - 1, 0)) = "" Then
        flxImport.Rows = flxImport.Rows - 1
    End If
    flxImport.Redraw = True
    If tmpSts > 0 Then HasIssue = True
    xlObject.DisplayAlerts = False 'To avoid "Save woorkbook" messagebox
    
    'Close Excel
    xlWB.Close
    xlObject.Application.Quit
    Set xlWB = Nothing
    Set xlObject = Nothing
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
QCTree.height = stsBar.Top - 550
flxImport.height = stsBar.Top - flxImport.Top - 250
flxImport.width = Me.width - flxImport.Left - 350
End Sub

Private Sub QCTree_DblClick()
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
Dim nodx As Node

    If QCTree.SelectedItem.Children <> 0 Then Exit Sub

    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT FC_ID, FC_NAME FROM COMPONENT_FOLDER WHERE FC_FATHER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY FC_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("F" & rs.FieldValue("FC_ID")), rs.FieldValue("FC_NAME"), 1
        rs.Next
    Next
    
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CO_ID, CO_NAME FROM COMPONENT, COMPONENT_FOLDER WHERE CO_FOLDER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " AND CO_FOLDER_ID = FC_ID ORDER BY CO_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("C" & rs.FieldValue("CO_ID")), rs.FieldValue("CO_NAME"), 3
        rs.Next
    Next
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "cmdRefresh"
    stsBar.SimpleText = "Preparing the process..."
    ClearForm
    stsBar.SimpleText = "Ready"
Case "cmdGenerate"
    stsBar.SimpleText = "Preparing the process..."
    On Error GoTo OutputErr
    GenerateOutput
    Exit Sub
OutputErr:
    MsgBox "Data was been truncated because of an error." & vbCrLf & Err.Description
Case "cmdOutput"
    If flxImport.Rows <= 1 Then
        MsgBox "Nothing to output", vbInformation
    Else
            stsBar.SimpleText = "Preparing the process..."
            OutputTable
    End If
End Select
End Sub

Private Sub ClearForm()
ClearTable
QCTree.Nodes.Clear

Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
    QCTree.Nodes.Add , , "Root", "Components", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT FC_ID, FC_NAME FROM COMPONENT_FOLDER WHERE FC_FATHER_ID = 1 ORDER BY FC_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("FC_ID")), rs.FieldValue("FC_NAME"), 1
        rs.Next
    Next
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.TextMatrix(0, 0) = "Step ID"
flxImport.TextMatrix(0, 1) = "Test Set Folder"
flxImport.TextMatrix(0, 2) = "Component Name"
flxImport.TextMatrix(0, 3) = "Step Order"
flxImport.TextMatrix(0, 4) = "Step Name"
flxImport.TextMatrix(0, 5) = "Description"
flxImport.TextMatrix(0, 6) = "Expected"
flxImport.TextMatrix(0, 7) = "Validation"
flxImport.Rows = 2
End Sub

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

Set xlWB = xlObject.Workbooks.Add
    xlObject.Sheets("Sheet2").Range("A1").Value = "1 - Only edit values in the column(s) colored green"
    xlObject.Sheets("Sheet2").Range("A2").Value = "2 - Do not Add, Delete or Modify Rows and Column's Position, Color or Order"
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
  curTab = "BC_STEPS-01"
  xlObject.Sheets("Sheet1").Name = curTab
  flxImport.FixedCols = 0
  flxImport.FixedRows = 0
  flxImport.RowSel = flxImport.Rows - 1
  flxImport.ColSel = flxImport.Cols - 1
  Clipboard.Clear
  Clipboard.SetText flxImport.Clip
  flxImport.FixedCols = 1
  flxImport.FixedRows = 1

  xlObject.Sheets(curTab).Range("A1").Select
  xlObject.Sheets(curTab).Paste

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

  xlObject.Sheets(curTab).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  xlObject.Workbooks(1).SaveAs "BC_STEPS-01" & QCTree.SelectedItem.Key & "-" & Format(Now, "mmddyyyy HHMM AMPM")
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate

  Set xlWB = Nothing
  Set xlObject = Nothing

  stsBar.SimpleText = "Ready"
On Error GoTo 0
End Sub
