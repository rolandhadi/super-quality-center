VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLoadTAOBC 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SAP TAO Components"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12675
   Icon            =   "frmLoadTAOBC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "SAP TAO Components"
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12675
      _ExtentX        =   22357
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
            Object.ToolTipText     =   "Upload to Consolidation List"
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
            Picture         =   "frmLoadTAOBC.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadTAOBC.frx":0B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadTAOBC.frx":0DEE
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
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   5355
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   9446
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
            Picture         =   "frmLoadTAOBC.frx":107C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadTAOBC.frx":178E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadTAOBC.frx":1EA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadTAOBC.frx":25B2
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
      Top             =   6045
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmLoadTAOBC.frx":2CC4
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   21608
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
            Picture         =   "frmLoadTAOBC.frx":3215
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadTAOBC.frx":34F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadTAOBC.frx":3A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadTAOBC.frx":3F99
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLoadTAOBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GenerateOutput()
Dim rs As TDAPIOLELib.Recordset
Dim AllScript
Dim objCommand
Dim i, j
Dim strPath
Dim iterd
Dim NewVal
Dim iter
Dim TimeSt
Dim AllF
Dim Decompile
Dim P
Dim k
    
    TimeSt = Format(Now, "mmm-dd-yyyy hhmmss") & "-"
    AllF = "BC Order"
    
    ReDim CheckedItems(0): strPath = ""
    GetAllCheckedItems QCTree.Nodes(1)
    For j = LBound(CheckedItems) To UBound(CheckedItems) - 1
        If Left(CheckedItems(j), 1) = "F" Then
            strPath = strPath & "AL_ABSOLUTE_PATH LIKE '" & GetFromTable(Right(CheckedItems(j), Len(CheckedItems(j)) - 1), "AL_ITEM_ID", "AL_ABSOLUTE_PATH", "ALL_LISTS") & "%'" & " OR "
        ElseIf Left(CheckedItems(j), 1) = "C" Then
            strPath = strPath & "TS_TEST_ID = " & Right(CheckedItems(j), Len(CheckedItems(j)) - 1) & " OR "
        End If
    Next
    If Trim(strPath) <> "" Then
        strPath = "(" & Left(strPath, Len(strPath) - 4) & ")"
    Else
        MsgBox "Please select and check source(s) in the HPQC folder tree"
        stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
        Exit Sub
    End If
    
    Set objCommand = QCConnection.Command
    
    objCommand.CommandText = "SELECT BC_ORDER, CO_ID, CO_NAME, FP_NAME, BPIP_VALUE FROM BP_ITER_PARAM, BP_PARAM, FRAMEWORK_PARAM, COMPONENT, TEST, BPTEST_TO_COMPONENTS, ALL_LISTS WHERE  BPIP_BPP_ID = BPP_ID AND BPP_PARAM_ID = FP_ID AND FP_COMPONENT_ID = CO_ID AND    TS_TEST_ID = BC_BPT_ID AND    BC_CO_ID = CO_ID AND    BC_ID = BPP_BPC_ID AND    TS_SUBJECT = AL_ITEM_ID AND " & strPath & " ORDER BY BC_ORDER, FP_ORDER"
    Debug.Print Me.Caption & "-" & objCommand.CommandText
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[SQL] " & Now & " " & objCommand.CommandText
    Set rs = objCommand.Execute                                                                                                                                                                                                                                                 'HERE!!!!!! <<<-------------
   
    ClearTable
    flxImport.Rows = rs.RecordCount + 1
    k = 0
    mdiMain.pBar.Max = rs.RecordCount + 3
    For i = 1 To rs.RecordCount
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Processing " & i & " out of " & rs.RecordCount
                k = k + 1
                flxImport.Rows = k + 1
                flxImport.TextMatrix(k, 0) = rs.FieldValue("BC_ORDER")
                flxImport.TextMatrix(k, 1) = rs.FieldValue("CO_ID")
                flxImport.TextMatrix(k, 2) = rs.FieldValue("CO_NAME")
                flxImport.TextMatrix(k, 3) = rs.FieldValue("FP_NAME")
                flxImport.TextMatrix(k, 4) = rs.FieldValue("BPIP_VALUE")
        mdiMain.pBar.Value = k + 1
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
        rs.Next
    Next
    mdiMain.pBar.Value = mdiMain.pBar.Max
    FXGirl.EZPlay FXSQCExtractCompleted
stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = mdiMain.pBar.Max & " record(s) generated"
End Sub

Function GetCommentText(rCommentCell As Range)
     Dim strGotIt As String
         On Error Resume Next
         strGotIt = WorksheetFunction.Clean _
             (rCommentCell.Comment.Text)
         GetCommentText = strGotIt
         On Error GoTo 0
End Function

Private Sub flxImport_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 67 And Shift = vbCtrlMask Then
    Clipboard.Clear
    Clipboard.SetText flxImport.Clip
End If
End Sub

Private Sub Form_Load()
ClearForm
End Sub

Private Sub QCTree_DblClick()
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
Dim nodx As Node

    If QCTree.SelectedItem.Children <> 0 Then Exit Sub
    
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT AL_ITEM_ID, AL_DESCRIPTION FROM ALL_LISTS WHERE AL_FATHER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY AL_DESCRIPTION"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("F" & rs.FieldValue("AL_ITEM_ID")), rs.FieldValue("AL_DESCRIPTION"), 1
        rs.Next
    Next
        
    If Left(QCTree.SelectedItem.Key, 1) = "F" Then
        Set objCommand = QCConnection.Command
        objCommand.CommandText = "SELECT DISTINCT TS_NAME, TS_TEST_ID FROM TEST, ALL_LISTS WHERE TS_SUBJECT = AL_ITEM_ID AND AL_ITEM_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY TS_NAME"
        Set rs = objCommand.Execute
        For i = 1 To rs.RecordCount
            QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("C" & rs.FieldValue("TS_TEST_ID")), rs.FieldValue("TS_NAME"), 2
            rs.Next
        Next
    End If

End Sub

Private Sub QCTree_NodeCheck(ByVal Node As MSComctlLib.Node)
Node.Selected = True
End Sub


Private Sub stsBar_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
frmLogs.txtLogs.Text = stsBar.Panels(2).Text: frmLogs.Show 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpR
Select Case Button.Key
Case "cmdRefresh"
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
    ClearForm
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
Case "cmdGenerate"
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
    On Error GoTo OutputErr
    GenerateOutput
    Exit Sub
OutputErr:
    MsgBox "Data was been truncated because of an error." & vbCrLf & Err.Description
Case "cmdOutput"
    If flxImport.Rows <= 1 Then
        MsgBox "Nothing to output", vbInformation
    Else
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
            OutputTable
    End If
Case "cmdUpload"
    If flxImport.Rows <= 2 Then
        MsgBox "Nothing to output", vbInformation
    Else
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
            TransferToConso
    End If
End Select
End Sub

Private Sub TransferToConso()
Dim fname As String
Dim lastrow
Dim i, j, tmpParam, k, X, Node_BC, Node_Param, Y
Dim tmpSts
Dim strFunct As New clsFiles
Dim stringFunct As New clsStrings
Dim intFunct As New clsInternet
Dim oldNode, tmpName

On Error GoTo ErrLoad
    
         lastrow = flxImport.Rows
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
        frmConsolidate.QCTree.Visible = False
        frmConsolidate.QCTree.Nodes.Clear
        tmpName = QCTree.SelectedItem.Text
        frmConsolidate.QCTree.Nodes.Add , , "Root", tmpName
        frmConsolidate.QCTree.Nodes.Item(X).Selected = True
        For i = 1 To lastrow - 1
            If flxImport.TextMatrix(i, 0) <> flxImport.TextMatrix(i - 1, 0) Then
                k = k + 1
                j = -1
                Node_BC = X
                X = 1
                frmConsolidate.QCTree.Nodes.Item(X).Selected = True
                frmConsolidate.QCTree.Nodes.Add frmConsolidate.QCTree.SelectedItem.Key, tvwChild, CStr("C" & k), "[" & Format(Trim(flxImport.TextMatrix(i, 0)), "000") & "] " & Trim(flxImport.TextMatrix(i, 2)), 4
                X = Node_BC
                X = frmConsolidate.QCTree.Nodes.Count
                j = j + 1
                Y = Y + 1
                frmConsolidate.QCTree.Nodes.Item(X).Selected = True
                frmConsolidate.QCTree.Nodes.Item(X).Tag = Trim(flxImport.TextMatrix(i, 1))
                frmConsolidate.QCTree.Nodes.Add frmConsolidate.QCTree.SelectedItem.Key, tvwChild, CStr("N" & Y), Trim(flxImport.TextMatrix(i, 3)), 2
                Node_BC = X
                frmConsolidate.QCTree.Nodes.Item(X + 1).Selected = True
                frmConsolidate.QCTree.Nodes.Add frmConsolidate.QCTree.SelectedItem.Key, tvwChild, CStr("V" & Y), Trim(flxImport.TextMatrix(i, 4)), 3
                frmConsolidate.QCTree.Nodes.Item(Node_BC).Expanded = False
            Else
                j = j + 1
                X = Node_BC
                Y = Y + 1
                frmConsolidate.QCTree.Nodes.Item(X).Selected = True
                frmConsolidate.QCTree.Nodes.Add frmConsolidate.QCTree.SelectedItem.Key, tvwChild, CStr("N" & Y), Trim(flxImport.TextMatrix(i, 3)), 2
                X = frmConsolidate.QCTree.Nodes.Count
                frmConsolidate.QCTree.Nodes.Item(X).Selected = True
                frmConsolidate.QCTree.Nodes.Add frmConsolidate.QCTree.SelectedItem.Key, tvwChild, CStr("V" & Y), Trim(flxImport.TextMatrix(i, 4)), 3
                frmConsolidate.QCTree.Nodes.Item(Node_BC).Expanded = False
                X = Node_BC
            End If
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = i - 1 & " out of " & lastrow - 1 & " validated " & Format(i / lastrow, "0.0%") & " (" & tmpSts & ") errors found."
            mdiMain.pBar.Value = i
        Next
    
    frmConsolidate.QCTree.Nodes(1).Expanded = True
    frmConsolidate.QCTree.Nodes(1).Selected = True
    frmConsolidate.QCTree.Visible = True
    mdiMain.pBar.Value = mdiMain.pBar.Max
        frmConsolidate.cmdUp.Enabled = True
        frmConsolidate.cmdDown.Enabled = True
        frmConsolidate.cmdAdd.Enabled = True
        frmConsolidate.cmdRemove.Enabled = True
        FXGirl.EZPlay FXSQCExtractCompleted
        frmConsolidate.stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: frmConsolidate.stsBar.Panels(2).Text = i - 1 & " out of " & lastrow - 1 & " validated " & Format(i / lastrow, "0.0%") & " (" & tmpSts & ") errors found."
Exit Sub
ErrLoad:
MsgBox "There was an error while importing the file. Please refresh and close all excel and try again" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub ClearForm()
ClearTable
QCTree.Nodes.Clear

Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
    QCTree.Nodes.Add , , "Root", "Subject", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT AL_ITEM_ID, AL_DESCRIPTION FROM ALL_LISTS WHERE AL_FATHER_ID = 2 ORDER BY AL_DESCRIPTION"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("AL_ITEM_ID")), rs.FieldValue("AL_DESCRIPTION"), 1
        rs.Next
    Next
    Me.Caption = Me.Tag
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.TextMatrix(0, 0) = "Order"
flxImport.TextMatrix(0, 1) = "Component ID"
flxImport.TextMatrix(0, 2) = "Component Name"
flxImport.TextMatrix(0, 3) = "Parameter Name"
flxImport.TextMatrix(0, 4) = "Parameter Value"
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

On Error GoTo OutErr: Set xlWB = xlObject.Workbooks.Add
    xlObject.Sheets("Sheet2").Range("A1").Value = "1 - Only edit values in the column(s) colored green"
    xlObject.Sheets("Sheet2").Range("A2").Value = "2 - Do not Add, Delete or Modify Rows and Column's Position, Color or Order"
    xlObject.Sheets("Sheet2").Range("A3").Value = "3 - The same sheet will be uploaded using SuperQC tools"
    xlObject.Sheets("Sheet2").Range("A4").Value = flxImport.TextMatrix(0, 0) & " First Entry:"
    xlObject.Sheets("Sheet2").Range("A5").Value = flxImport.TextMatrix(0, 0) & " Last Entry:"
    xlObject.Sheets("Sheet2").Range("A6").Value = "Total Records:"
    xlObject.Sheets("Sheet2").Range("B4").Value = flxImport.TextMatrix(1, 0)
    xlObject.Sheets("Sheet2").Range("B5").Value = flxImport.TextMatrix(flxImport.Rows - 1, 0)
    xlObject.Sheets("Sheet2").Range("B6").Value = flxImport.Rows - 1
    
    xlObject.Sheets("Sheet2").Range("A7").Value = "Environment:"
    xlObject.Sheets("Sheet2").Range("B7").Value = curDomain & "-" & curProject
    
    xlObject.Sheets("Sheet2").Range("A8").Value = "Test Plan ID:"
    xlObject.Sheets("Sheet2").Range("B8").Value = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
    
    xlObject.Sheets("Sheet2").Range("A9").Value = "BC ID:"
    xlObject.Sheets("Sheet2").Range("B9").Value = "NONE"
    
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
  flxImport.FixedCols = 0
  flxImport.FixedRows = 0
  flxImport.col = 0
  flxImport.row = 0
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
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & "TAO Automation Module " & curDomain & "-" & curProject & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""
    
  xlObject.Workbooks(1).SaveAs "SAPTAO01-" & CleanTheString(QCTree.SelectedItem.Text) & "-" & Format(Now, "mmddyyyy HHMM AMPM")
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate
  
  Set xlWB = Nothing
  Set xlObject = Nothing
  FXGirl.EZPlay FXExportToExcel
  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export to MS Excel completed.": Exit Sub:
OutErr:     MsgBox Err.Description, vbCritical: xlObject.Visible = True: xlObject.ActiveWindow.Activate: Set xlWB = Nothing: Set xlObject = Nothing
On Error GoTo 0
End Sub
