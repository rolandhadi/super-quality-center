VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLoadManualBC 
   Caption         =   "Upload Business Components Module"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   Icon            =   "frmLoadManualBC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   12660
   Tag             =   "Upload Business Components Module"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Override Template"
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   540
      Width           =   12435
      Begin VB.CheckBox chkNeedProcessTeam 
         Caption         =   "Process Team Required?"
         Height          =   255
         Left            =   7140
         TabIndex        =   16
         Top             =   1200
         Value           =   1  'Checked
         Width           =   3675
      End
      Begin VB.CheckBox chkOverride 
         Caption         =   "Overrride Template"
         Height          =   255
         Left            =   60
         TabIndex        =   15
         Top             =   240
         Width           =   3675
      End
      Begin VB.TextBox txtEndDate 
         Height          =   285
         Left            =   8580
         TabIndex        =   14
         Top             =   840
         Width           =   3675
      End
      Begin VB.TextBox txtStartDate 
         Height          =   285
         Left            =   8580
         TabIndex        =   13
         Top             =   540
         Width           =   3675
      End
      Begin VB.TextBox txtQAReviewer 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Text            =   "user"
         Top             =   1140
         Width           =   5175
      End
      Begin VB.TextBox txtStatus 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Text            =   "040 Ready For QA Review"
         Top             =   840
         Width           =   5175
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Text            =   "Components\030 Release 3\300 Unit Test\"
         Top             =   540
         Width           =   5175
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Scripting End Date:"
         Height          =   255
         Left            =   7140
         TabIndex        =   12
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Scripting Start Date:"
         Height          =   255
         Left            =   7140
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "QA Reviewer"
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   255
         Left            =   60
         TabIndex        =   6
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "HPQC Folder Path:"
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12660
      _ExtentX        =   22331
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
            Key             =   "cmdConsolidate"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "cmdByColor"
                  Text            =   "Consolidate By Color"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "cmdByParams"
                  Text            =   "Consolidate By Parameters"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   2650
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         Left            =   3540
         Picture         =   "frmLoadManualBC.frx":08CA
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
         Picture         =   "frmLoadManualBC.frx":30EE
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
            Picture         =   "frmLoadManualBC.frx":3894
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadManualBC.frx":3B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadManualBC.frx":3DB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   3795
      Left            =   60
      TabIndex        =   1
      Top             =   2160
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   6694
      _Version        =   393216
      Cols            =   7
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
            Picture         =   "frmLoadManualBC.frx":4046
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadManualBC.frx":4758
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadManualBC.frx":4E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadManualBC.frx":557C
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
      TabIndex        =   17
      Top             =   6000
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   670
            MinWidth        =   670
            Picture         =   "frmLoadManualBC.frx":5C8E
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   21105
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
            Picture         =   "frmLoadManualBC.frx":61DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadManualBC.frx":64C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadManualBC.frx":6A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadManualBC.frx":6F63
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLoadManualBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

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
    STEP_DESCRIPTION() As String
    Step_ExpectedResult() As String
    Log As String
    StepGroup As Integer
End Type

Private All_BC() As BC_Component
Private HasIssue As Boolean
Private HasUploadIssue  As Integer

Private Function LoadToArray()
Dim lastrow, i, LastVal, EndArr
lastrow = flxImport.Rows - 1
ReDim All_BC(0)
EndArr = -1
LastVal = 0
For i = 1 To lastrow
    If Trim(flxImport.TextMatrix(i, 0)) = "" Or Trim(flxImport.TextMatrix(i, 1)) = "" Then
        All_BC(EndArr).Log = All_BC(EndArr).Log & vbCrLf & "Line " & i & " is blank"
    ElseIf (Trim(UCase(flxImport.TextMatrix(i, 0))) <> Trim(UCase(flxImport.TextMatrix(i - 1, 0)))) Or (Trim(UCase(flxImport.TextMatrix(i, 1))) <> Trim(UCase(flxImport.TextMatrix(i - 1, 1)))) Then
        EndArr = EndArr + 1
        ReDim Preserve All_BC(EndArr)
        LastVal = LastVal + 1
        All_BC(EndArr).Group_No = LastVal
        All_BC(EndArr).Component_Path = flxImport.TextMatrix(i, 0)
        All_BC(EndArr).Component_Name = ReplaceAllEnter(flxImport.TextMatrix(i, 1))
        All_BC(EndArr).Component_Description = "" 'Range("E" & i).Value
        All_BC(EndArr).Scripter = flxImport.TextMatrix(i, 8)
        All_BC(EndArr).Peer_Reviewer = flxImport.TextMatrix(i, 9)
        All_BC(EndArr).QA_Reviewer = flxImport.TextMatrix(i, 3)
        All_BC(EndArr).Planned_Start_Date = flxImport.TextMatrix(i, 5)
        All_BC(EndArr).Planned_End_Date = flxImport.TextMatrix(i, 6)
        All_BC(EndArr).Status = flxImport.TextMatrix(i, 2)
        
        ReDim All_BC(EndArr).Step_Order(0)
        ReDim All_BC(EndArr).Step_Name(0)
        ReDim All_BC(EndArr).STEP_DESCRIPTION(0)
        ReDim All_BC(EndArr).Step_ExpectedResult(0)
        
        All_BC(EndArr).Step_Order(0) = 1
        All_BC(EndArr).Step_Name(0) = "Step " & UBound(All_BC(EndArr).Step_Name) + 1
        All_BC(EndArr).STEP_DESCRIPTION(0) = flxImport.TextMatrix(i, 11)
        All_BC(EndArr).Step_ExpectedResult(0) = flxImport.TextMatrix(i, 12)
    Else
        ReDim Preserve All_BC(EndArr).Step_Order(UBound(All_BC(EndArr).Step_Order) + 1)
        ReDim Preserve All_BC(EndArr).Step_Name(UBound(All_BC(EndArr).Step_Name) + 1)
        ReDim Preserve All_BC(EndArr).STEP_DESCRIPTION(UBound(All_BC(EndArr).STEP_DESCRIPTION) + 1)
        ReDim Preserve All_BC(EndArr).Step_ExpectedResult(UBound(All_BC(EndArr).Step_ExpectedResult) + 1)
        
        All_BC(EndArr).Step_Order(UBound(All_BC(EndArr).Step_Order)) = UBound(All_BC(EndArr).Step_Order) + 1
        All_BC(EndArr).Step_Name(UBound(All_BC(EndArr).Step_Name)) = "Step " & UBound(All_BC(EndArr).Step_Name) + 1
        All_BC(EndArr).STEP_DESCRIPTION(UBound(All_BC(EndArr).STEP_DESCRIPTION)) = flxImport.TextMatrix(i, 11)
        All_BC(EndArr).Step_ExpectedResult(UBound(All_BC(EndArr).Step_ExpectedResult)) = flxImport.TextMatrix(i, 12)
    End If
Next
End Function

Function LoadToQC()
Dim i, j
Dim tmpComp
stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = ""
mdiMain.pBar.Max = UBound(All_BC) + 3
For i = LBound(All_BC) To UBound(All_BC)
    On Error Resume Next
        Set tmpComp = Create_New_Component(All_BC(i))
    If Err.Number = 0 Then
        Err.Clear
        On Error GoTo 0
        On Error Resume Next
        Add_Params_And_Steps tmpComp, All_BC(i).Step_Order, All_BC(i).Step_Name, All_BC(i).STEP_DESCRIPTION, All_BC(i).Step_ExpectedResult
    Else
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[CREATE BC: (FAILED) " & Now & " " & All_BC(i).Component_Path & "-" & All_BC(i).Component_Name & "] " & Err.Description
        HasUploadIssue = HasUploadIssue + 1
        Err.Clear
        On Error GoTo 0
    End If
    stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Loading Business Component " & i + 1 & " out of " & UBound(All_BC) + 1 & " (" & All_BC(i).Component_Name & ")"
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
Next: i = i - 1
mdiMain.pBar.Value = mdiMain.pBar.Max
FXGirl.EZPlay FXDataUploadCompleted
If HasUploadIssue <> 0 Then
      Dim tmpFile As New clsFiles
      frmLogs.Caption = App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log"
      frmLogs.txtLogs.Text = tmpFile.ReadFromFile_FAILED(App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log")
      frmLogs.Show 1
Else
      FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[CREATE BC: (PASSED) " & Now & " " & All_BC(i).Component_Path & "-" & All_BC(i).Component_Name & "]"
End If
stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = UBound(All_BC) + 1 & " Business Component(s) loaded successfully. (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log"
QCConnection.SendMail "user@companyemail.com", "", "[HPQC UPDATES] Business Component(s) loaded successfully by " & curUser & " in " & curDomain & "-" & curProject, UBound(All_BC) + 1 & " Business Component(s) loaded successfully. (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
QCConnection.SendMail curUser, "", "[HPQC UPDATES] Business Component(s) loaded successfully by " & curUser & " in " & curDomain & "-" & curProject, UBound(All_BC) + 1 & " Business Component(s) loaded successfully. (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
End Function

Sub Start()
Debug.Print "New Session: " & Now
LoadToArray
LoadToQC
Debug.Print "New Finished: " & Now
End Sub

'########################### Create New Component ###########################
Private Function Create_New_Component(tmpComp As BC_Component) As Component
    Dim myComp As Component
    Dim compFactory As ComponentFactory
    Dim generalComponentFolderFactory As ComponentFolderFactory
    Dim cFolder As ComponentFolder, rootCFolder As ComponentFolder
    On Error GoTo NewComponentErr
    ' Get Component Folder
    ' Get a ComponentFolderFactory from the QCConnectiononnection object
    Set generalComponentFolderFactory = QCConnection.ComponentFolderFactory
    CreateComponentFolder tmpComp.Component_Path
    Set rootCFolder = generalComponentFolderFactory.FolderByPath(tmpComp.Component_Path)
    If IsMissing(rootCFolder) Then
        ' Get the root folder
        Set rootCFolder = generalComponentFolderFactory.Root
    End If
    ' Get a ComponentFactory
    Set compFactory = rootCFolder.ComponentFactory
    ' Add the component
    Set myComp = compFactory.AddItem(Null)
    Dim errString As String
    If (compFactory.IsComponentNameValid(tmpComp.Component_Name, errString)) Then
        myComp.Name = tmpComp.Component_Name
    End If
    myComp.Field("CO_RESPONSIBLE") = tmpComp.Scripter  'Scripter
    myComp.Field("CO_DESC") = "" 'tmpComp.Component_Description   'Component Description
    myComp.Field("CO_USER_TEMPLATE_01") = tmpComp.Status  'Status
    myComp.Field("CO_USER_TEMPLATE_02") = tmpComp.Peer_Reviewer  'Peer Reviewer
    myComp.Field("CO_USER_TEMPLATE_03") = tmpComp.QA_Reviewer  'QA Reviewer
    myComp.Field("CO_USER_TEMPLATE_05") = tmpComp.Planned_End_Date  'Planned End
    myComp.Field("CO_USER_TEMPLATE_04") = tmpComp.Planned_Start_Date  'Planned Start
    myComp.Field("CO_USER_TEMPLATE_08") = "Manual"
    myComp.Post
    'Return the new component.
    Set Create_New_Component = myComp
Exit Function
NewComponentErr:
    Dim tmpFile As New clsFiles
    HasUploadIssue = HasUploadIssue + 1
    Debug.Print "Error in creating component in line " & tmpComp.Component_Name
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[BC CREATE: (FAILED) " & Now & " " & tmpComp.Component_Name & "] " & Err.Description
End Function
'########################### End Of Create New Component ##########################

'########################### Add Steps and Parameters ###########################
Private Function Add_Params_And_Steps(ByRef myComponent, StepOrder() As Integer, StepName() As String, StepDesc() As String, StepExp() As String)
    Dim compParamFactory As ComponentParamFactory
    Dim compParam() As ComponentParam
    Dim compStepFactory As ComponentStepFactory
    Dim compStep() As ComponentStep
    Dim i, j, tmpParam()
    Dim pString As String
    Dim AllParam()
    Dim tmpFile As New clsFiles
' This example adds a parameter and step to the Component.
    Set compParamFactory = myComponent.ComponentParamFactory
    Set compStepFactory = myComponent.ComponentStepFactory
    ReDim compParam(0)
    ReDim compStep(0)
    ReDim AllParam(0)
' Create a parameter and set the properties.
On Error GoTo Err1
    For i = LBound(StepDesc) To UBound(StepDesc)
        If HasParameters(StepDesc(i)) = True Then
        tmpParam = ExtractParameters(StepDesc(i))
            For j = LBound(tmpParam) To UBound(tmpParam)
                If IsParameterDeclared(AllParam, tmpParam(j)) = False Then
                    ReDim Preserve compParam(j + 1)
                    Set compParam(j) = compParamFactory.AddItem(Null)
                    'Output parameter. For ComponentParam.IsOut, 1 true and 0 is false
                    If UCase(Left(tmpParam(j), 1)) = "O" Then
                        compParam(j).IsOut = 1
                    Else
                        compParam(j).IsOut = 0
                    End If
                    compParam(j).Name = Replace(LCase(tmpParam(j)), "-", "_")
                    compParam(j).Desc = LCase(tmpParam(j))
                    compParam(j).ValueType = "String"
                    'compParam(j).Order = 1
                    compParam(j).Post
                    AllParam(UBound(AllParam)) = LCase(tmpParam(j))
                    ReDim Preserve AllParam(UBound(AllParam) + 1)
                End If
            Next
        End If
    Next
    ReDim tmpParam(0)
    For i = LBound(StepExp) To UBound(StepExp)
        If Trim(StepExp(i)) <> "" Then
            If HasParameters(StepExp(i)) = True Then
            tmpParam = ExtractParameters(StepExp(i))
                For j = LBound(tmpParam) To UBound(tmpParam)
                    If IsParameterDeclared(AllParam, tmpParam(j)) = False Then
                        ReDim Preserve compParam(j + 1)
                        Set compParam(j) = compParamFactory.AddItem(Null)
                        'Output parameter. For ComponentParam.IsOut, 1 true and 0 is false
                        If UCase(Left(tmpParam(j), 1)) = "O" Then
                            compParam(j).IsOut = 1
                        Else
                            compParam(j).IsOut = 0
                        End If
                        compParam(j).Name = Replace(LCase(tmpParam(j)), "-", "_")
                        compParam(j).Desc = LCase(tmpParam(j))
                        compParam(j).ValueType = "String"
                        'compParam(j).Order = 1
                        compParam(j).Post
                        AllParam(UBound(AllParam)) = LCase(tmpParam(j))
                        ReDim Preserve AllParam(UBound(AllParam) + 1)
                    End If
                Next
            End If
        End If
    Next
On Error GoTo 0

On Error GoTo Err2
' Create a step
    For i = LBound(StepDesc) To UBound(StepDesc)
        ReDim tmpParam(0)
        ReDim Preserve compStep(i + 1)
        Set compStep(i) = compStepFactory.AddItem(Null)
        compStep(i).StepName = StepName(i)
        ' Add description and expected results that contain parameters.
        'Input parameter. For this example, assume that InParam was previously created.
        ' pString ="Description can contain <<<InParam>>> that are verified by QC"
        If HasParameters(StepDesc(i)) = True Then
            tmpParam = ExtractParameters(StepDesc(i))
            For j = LBound(tmpParam) To UBound(tmpParam)
                StepDesc(i) = Replace(StepDesc(i), "<<<" & tmpParam(j) & ">>>", "<<<" & LCase(tmpParam(j)) & ">>>")
            Next
        End If
        StepDesc(i) = CleanHTML_BC(StepDesc(i))
        pString = "<html><body>" & StepDesc(i) & "</body></html>"
        compStep(i).StepDescription = pString
        'Output parameter. OutParam was previously created. badParamName was not.
        ' pString ="Can also contain <<<OutParam>>> and <<<badParamName>>>"
        'Given that badParamName is not the name of any parameter,
        ' <<<badParamName>>> will be replaced by <badParamName>, so that
        ' the user can see in the user interface that it is not recognized.
        If HasParameters(StepExp(i)) = True Then
            tmpParam = ExtractParameters(StepExp(i))
            For j = LBound(tmpParam) To UBound(tmpParam)
                StepExp(i) = Replace(StepExp(i), "<<<" & tmpParam(j) & ">>>", "<<<" & LCase(tmpParam(j)) & ">>>")
            Next
        End If
        StepExp(i) = CleanHTML_BC(StepExp(i))
        pString = "<html><body>" & StepExp(i) & "</body></html>"
        compStep(i).StepExpectedResult = pString
        'Set the execution order
        compStep(i).Order = CInt(StepOrder(i))
        'Finish the step
        'Validate will throw an exception because of the badParamName.
        ' In real production code, you may wish to catch the exception
        ' and take action rather than continuing as this example does.
        'On Error GoTo myInvalidStepHandler
        On Error Resume Next
        compStep(i).Validate
        compStep(i).Post
    Next

    ' Save changes and unlock the component.
    myComponent.Post
    myComponent.UnLockObject
Exit Function
Err1:
HasUploadIssue = HasUploadIssue + 1
Debug.Print "Error in Parameter in line " & i
FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[PARAM CREATE: (FAILED) " & Now & "] In line " & i & " (" & myComponent.Name & ") | " & Err.Description
Exit Function
Err2:
HasUploadIssue = HasUploadIssue + 1
FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[STEP CREATE: (FAILED) " & Now & "] In line " & i & " (" & myComponent.Name & ") | " & Err.Description
Debug.Print "Error in Step in line " & i
End Function
'########################### End Of Add Steps and Parameters ###########################


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
Dim i, j, tmpParam, tmpParamOrig
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
         If UCase(Trim(.Range("A2").Value)) <> UCase(Trim("HPQC Folder Path")) Or InStr(1, .Name, "Unit Test BC Template") <> 0 Then
            MsgBox "Import file is invalid. Please use only sheets generated by the SuperQC"
            xlObject.DisplayAlerts = False 'To avoid "Save woorkbook" messagebox
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
         lastrow = .Range("H" & .Rows.Count).End(xlUp).row
        '.Range("A3:M" & LastRow).Copy 'Set selection to Copy
        
        ClearTable
        flxImport.Redraw = False     'Dont draw until the end, so we avoid that flash
        flxImport.row = 0            'Paste from first cell
        flxImport.col = 0
        flxImport.Rows = lastrow
        flxImport.Cols = 14
        flxImport.Redraw = False
        
        'A - Load HPQC Folder Path
        'Should not be blank
        mdiMain.pBar.Max = lastrow + 2
        For i = 3 To lastrow
            
            If chkOverride.Value = Checked Then
                                           
                If Left(Trim(.Range("H" & i).Value), 1) = "_" Then
                    flxImport.TextMatrix(i - 2, 1) = CleanTheString(Right(Trim(.Range("H" & i).Value), Len(Trim(.Range("H" & i).Value)) - 1))  'Change number and letter
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Component Name=REMOVED'_']"
                Else
                    flxImport.TextMatrix(i - 2, 1) = Trim(CleanTheString(.Range("H" & i).Value))    'Change number and letter
                End If
                
                If Right(Trim(flxImport.TextMatrix(i - 2, 0)), 1) = "_" Then
                    flxImport.TextMatrix(i - 2, 0) = Left(flxImport.TextMatrix(i - 2, 0), Len(flxImport.TextMatrix(i - 2, 0)) - 1)
                    flxImport.TextMatrix(i - 2, 0) = flxImport.TextMatrix(i - 2, 0) & Left(flxImport.TextMatrix(i - 2, 1), 1)
                End If
                
                If Trim(.Range("H" & i).Value) = "" Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Component Name=BLANK]"
                    tmpSts = tmpSts + 1
                End If
                
                If chkNeedProcessTeam.Value = Checked Then
                    If InStr(1, flxImport.TextMatrix(i - 2, 1), "_WMT_", vbTextCompare) = 0 And InStr(1, flxImport.TextMatrix(i - 2, 1), "_SUP_", vbTextCompare) = 0 And InStr(1, flxImport.TextMatrix(i - 2, 1), "_FIN_", vbTextCompare) = 0 And InStr(1, flxImport.TextMatrix(i - 2, 1), "_MKT_", vbTextCompare) = 0 And InStr(1, flxImport.TextMatrix(i - 2, 1), "_PIT_", vbTextCompare) = 0 And InStr(1, flxImport.TextMatrix(i - 2, 1), "_HUR_", vbTextCompare) = 0 And InStr(1, flxImport.TextMatrix(i - 2, 1), "_HRM_", vbTextCompare) = 0 And InStr(1, flxImport.TextMatrix(i - 2, 1), "_PIT_", vbTextCompare) = 0 Then
                        flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Component Name=NO PROCESS TEAM]"
                        tmpSts = tmpSts + 1
                    End If
                End If
                
                flxImport.TextMatrix(i - 2, 0) = strFunct.RemoveBackslash(((txtPath.Text & Left(Trim(flxImport.TextMatrix(i - 2, 1)), 1))))           'Change number and letter
                If Trim(flxImport.TextMatrix(i - 2, 0)) = "" Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Component Path=BLANK]"
                    tmpSts = tmpSts + 1
                End If
                
                If InStr(1, Trim(flxImport.TextMatrix(i - 2, 0)), "Components\", vbTextCompare) = 0 Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Component Path Invalid]"
                    tmpSts = tmpSts + 1
                End If
                
                flxImport.TextMatrix(i - 2, 2) = Trim(txtStatus.Text)     'Change number and letter
                If Trim(flxImport.TextMatrix(i - 2, 2)) = "" Or Trim(flxImport.TextMatrix(i - 2, 2)) <> "040 Ready For QA Review" And Trim(flxImport.TextMatrix(i - 2, 2)) <> "060 QA Review Completed" Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Status=INCORRECT]"
                    tmpSts = tmpSts + 1
                End If
                
                flxImport.TextMatrix(i - 2, 3) = Trim(txtQAReviewer.Text)     'Change number and letter
                If Trim(flxImport.TextMatrix(i - 2, 3)) = "" Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[QA Reviewer=BLANK]"
                    tmpSts = tmpSts + 1
                End If
                
                flxImport.TextMatrix(i - 2, 4) = Trim("BUSINESS-PROCESS")    'Change number and letter
                
                flxImport.TextMatrix(i - 2, 5) = Format(Trim(txtStartDate.Text), "dd/mm/yyyy")    'Change number and letter
                If Trim(flxImport.TextMatrix(i - 2, 5)) = "" Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Planned Scripting Start Date]=BLANK"
                    tmpSts = tmpSts + 1
                End If
                If InStr(1, Trim(flxImport.TextMatrix(i - 2, 5)), ".") <> 0 Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Planned Scripting Start Date=INVALID FORMAT]"
                    tmpSts = tmpSts + 1
                End If
                
                flxImport.TextMatrix(i - 2, 6) = Format(Trim(txtEndDate.Text), "dd/mm/yyyy")     'Change number and letter
                If Trim(flxImport.TextMatrix(i - 2, 6)) = "" Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Planned Scripting End Date=BLANK]"
                    tmpSts = tmpSts + 1
                End If
                If InStr(1, Trim(flxImport.TextMatrix(i - 2, 6)), ".") <> 0 Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Planned Scripting End Date=INVALID FORMAT]"
                    tmpSts = tmpSts + 1
                End If
            Else ' ORIGINAL >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                flxImport.TextMatrix(i - 2, 0) = strFunct.RemoveBackslash(((.Range("A" & i).Value)))          'Change number and letter
                If Trim(.Range("A" & i).Value) = "" Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Component Path=BLANK]"
                    tmpSts = tmpSts + 1
                End If
                
                If InStr(1, Trim(.Range("A" & i).Value), "Components\", vbTextCompare) = 0 Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Component Path Invalid]"
                    tmpSts = tmpSts + 1
                End If
                            
                If Left(Trim(.Range("B" & i).Value), 1) = "_" Then
                    flxImport.TextMatrix(i - 2, 1) = CleanTheString(Right(Trim(.Range("B" & i).Value), Len(Trim(.Range("B" & i).Value)) - 1))  'Change number and letter
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Component Name=REMOVED'_']"
                Else
                    flxImport.TextMatrix(i - 2, 1) = CleanTheString(Trim(.Range("B" & i).Value))    'Change number and letter
                End If
                
                If Right(Trim(flxImport.TextMatrix(i - 2, 0)), 1) = "_" Then
                    flxImport.TextMatrix(i - 2, 0) = Left(flxImport.TextMatrix(i - 2, 0), Len(flxImport.TextMatrix(i - 2, 0)) - 1)
                    flxImport.TextMatrix(i - 2, 0) = flxImport.TextMatrix(i - 2, 0) & Left(flxImport.TextMatrix(i - 2, 1), 1)
                End If
                
                If Trim(.Range("B" & i).Value) = "" Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Component Name=BLANK]"
                    tmpSts = tmpSts + 1
                End If
                
                If chkNeedProcessTeam.Value = Checked Then
                    If InStr(1, flxImport.TextMatrix(i - 2, 1), "_WMT_", vbTextCompare) = 0 And InStr(1, flxImport.TextMatrix(i - 2, 1), "_SUP_", vbTextCompare) = 0 And InStr(1, flxImport.TextMatrix(i - 2, 1), "_FIN_", vbTextCompare) = 0 And InStr(1, flxImport.TextMatrix(i - 2, 1), "_MKT_", vbTextCompare) = 0 And InStr(1, flxImport.TextMatrix(i - 2, 1), "_PIT_", vbTextCompare) = 0 And InStr(1, flxImport.TextMatrix(i - 2, 1), "_HUR_", vbTextCompare) = 0 And InStr(1, flxImport.TextMatrix(i - 2, 1), "_PIT_", vbTextCompare) = 0 Then
                        flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Component Name=NO PROCESS TEAM]"
                        tmpSts = tmpSts + 1
                    End If
                End If
                
                flxImport.TextMatrix(i - 2, 2) = Trim(.Range("C" & i).Value)    'Change number and letter
                If Trim(.Range("C" & i).Value) = "" Or (Trim(.Range("C" & i).Value) = "040 Ready For QA Review" And Trim(.Range("C" & i).Value) = "060 QA Review Completed") Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Status=INCORRECT]"
                    tmpSts = tmpSts + 1
                End If
                
                flxImport.TextMatrix(i - 2, 3) = Trim(.Range("D" & i).Value)    'Change number and letter
                If Trim(.Range("D" & i).Value) = "" Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[QA Reviewer=BLANK]"
                    tmpSts = tmpSts + 1
                End If
                
                flxImport.TextMatrix(i - 2, 4) = Trim("BUSINESS-PROCESS")    'Change number and letter
                
                flxImport.TextMatrix(i - 2, 5) = Format(Trim(.Range("F" & i).Value), "dd/mm/yyyy")    'Change number and letter
                If Trim(.Range("F" & i).Value) = "" Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Planned Scripting Start Date]=BLANK"
                    tmpSts = tmpSts + 1
                End If
                
                flxImport.TextMatrix(i - 2, 6) = Format(Trim(.Range("G" & i).Value), "dd/mm/yyyy")    'Change number and letter
                If Trim(.Range("G" & i).Value) = "" Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Planned Scripting End Date=BLANK]"
                    tmpSts = tmpSts + 1
                End If
            End If
            
           If Left(Trim(.Range("H" & i).Value), 1) = "_" Then
                flxImport.TextMatrix(i - 2, 7) = Right(Trim(.Range("H" & i).Value), Len(Trim(.Range("H" & i).Value)) - 1)  'Change number and letter
            Else
                flxImport.TextMatrix(i - 2, 7) = Trim(.Range("H" & i).Value)    'Change number and letter
            End If
            
            If Trim(.Range("H" & i).Value) = "" Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Component Name=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 2, 8) = Trim(.Range("I" & i).Value)    'Change number and letter
            If Trim(.Range("I" & i).Value) = "" Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Scripter=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 2, 9) = Trim(.Range("J" & i).Value)    'Change number and letter
            If Trim(.Range("J" & i).Value) = "" Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Peer Reviewer=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 2, 10) = Trim(.Range("K" & i).Value)
            
            flxImport.TextMatrix(i - 2, 11) = Trim(.Range("L" & i).Value)    'Change number and letter
            If Trim(flxImport.TextMatrix(i - 2, 11)) = "" Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Step Description=BLANK]"
                tmpSts = tmpSts + 1
            End If
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), "<<<<", "<<<")
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), "<<<<", "<<<")
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), "<<<<", "<<<")
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), "<<<<", "<<<")
            If InStr(1, flxImport.TextMatrix(i - 2, 11), "<<<<", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Step Description=<<<< FOUND]"
                tmpSts = tmpSts + 1
            End If
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), ">>>>", ">>>")
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), ">>>>", ">>>")
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), ">>>>", ">>>")
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), ">>>>", ">>>")
            If InStr(1, flxImport.TextMatrix(i - 2, 11), ">>>>", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Step Description=>>>> FOUND]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, UCase(flxImport.TextMatrix(i - 2, 11)), " <<P_", vbTextCompare) <> 0 Or InStr(1, UCase(flxImport.TextMatrix(i - 2, 11)), vbCrLf & "<<P_", vbTextCompare) <> 0 Or InStr(1, UCase(flxImport.TextMatrix(i - 2, 11)), " <<O_", vbTextCompare) <> 0 Or InStr(1, UCase(flxImport.TextMatrix(i - 2, 11)), vbCrLf & "<<O_", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Step Description=<< FOUND]"
                tmpSts = tmpSts + 1
            End If
            If HasParameters(flxImport.TextMatrix(i - 2, 11)) = True Then
                tmpParamOrig = ExtractParameters(flxImport.TextMatrix(i - 2, 11))
                tmpParam = ExtractParametersWithFix(flxImport.TextMatrix(i - 2, 11))
                For j = LBound(tmpParam) To UBound(tmpParam)
                    flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), tmpParamOrig(j), tmpParam(j))
                Next
                For j = LBound(tmpParam) To UBound(tmpParam)
                    If InvalidParameterCheck(tmpParam(j)) = True Then
                        flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Parameter=INVALID FORMAT/CHAR]"
                        tmpSts = tmpSts + 1
                    End If
                Next
            End If
            For j = 1 To 26
                If InStr(1, flxImport.TextMatrix(i - 2, 11), "<<<" & Chr(j + 64), vbTextCompare) <> 0 And (LCase(Chr(j + 64)) <> "p" And LCase(Chr(j + 64)) <> "o") Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Step Description=PARAMETER FORMAT FAIL]"
                    tmpSts = tmpSts + 1
                End If
            Next
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), "<<< ", "<<<")
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), "<<< ", "<<<")
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), "<<< ", "<<<")
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), "<<< ", "<<<")
            If InStr(1, flxImport.TextMatrix(i - 2, 11), "<<< ", vbTextCompare) Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Step Description=PARAMETER FORMAT FAIL]"
                tmpSts = tmpSts + 1
            End If
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), " >>>", ">>>")
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), " >>>", ">>>")
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), " >>>", ">>>")
            flxImport.TextMatrix(i - 2, 11) = Replace(flxImport.TextMatrix(i - 2, 11), " >>>", ">>>")
            If InStr(1, flxImport.TextMatrix(i - 2, 11), " >>>", vbTextCompare) Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Step Description=PARAMETER FORMAT FAIL]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 2, 12) = Trim(.Range("M" & i).Value)    'Change number and letter
            
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), "<<<<", "<<<")
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), "<<<<", "<<<")
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), "<<<<", "<<<")
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), "<<<<", "<<<")
            If InStr(1, flxImport.TextMatrix(i - 2, 12), "<<<<", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Step Description=<<<< FOUND]"
                tmpSts = tmpSts + 1
            End If
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), ">>>>", ">>>")
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), ">>>>", ">>>")
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), ">>>>", ">>>")
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), ">>>>", ">>>")
            If InStr(1, flxImport.TextMatrix(i - 2, 12), ">>>>", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Step Description=>>>> FOUND]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, UCase(flxImport.TextMatrix(i - 2, 12)), " <<P_", vbTextCompare) <> 0 Or InStr(1, UCase(flxImport.TextMatrix(i - 2, 12)), vbCrLf & "<<P_", vbTextCompare) <> 0 Or InStr(1, UCase(flxImport.TextMatrix(i - 2, 12)), " <<O_", vbTextCompare) <> 0 Or InStr(1, UCase(flxImport.TextMatrix(i - 2, 12)), vbCrLf & "<<O_", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Step Description=<< FOUND]"
                tmpSts = tmpSts + 1
            End If
            If HasParameters(flxImport.TextMatrix(i - 2, 12)) = True Then
                tmpParamOrig = ExtractParameters(flxImport.TextMatrix(i - 2, 12))
                tmpParam = ExtractParametersWithFix(flxImport.TextMatrix(i - 2, 12))
                For j = LBound(tmpParam) To UBound(tmpParam)
                    flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), tmpParamOrig(j), tmpParam(j))
                Next
                For j = LBound(tmpParam) To UBound(tmpParam)
                    If InvalidParameterCheck(tmpParam(j)) = True Then
                        flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Parameter=INVALID FORMAT/CHAR]"
                        tmpSts = tmpSts + 1
                    End If
                Next
            End If
            For j = 1 To 26
                If InStr(1, flxImport.TextMatrix(i - 2, 12), "<<<" & Chr(j + 64), vbTextCompare) <> 0 And (LCase(Chr(j + 64)) <> "p" And LCase(Chr(j + 64)) <> "o") Then
                    flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Step Description=PARAMETER FORMAT FAIL]"
                    tmpSts = tmpSts + 1
                End If
            Next
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), "<<< ", "<<<")
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), "<<< ", "<<<")
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), "<<< ", "<<<")
            If InStr(1, flxImport.TextMatrix(i - 2, 12), "<<< ", vbTextCompare) Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Step Description=PARAMETER FORMAT FAIL]"
                tmpSts = tmpSts + 1
            End If
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), " >>>", ">>>")
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), " >>>", ">>>")
            flxImport.TextMatrix(i - 2, 12) = Replace(flxImport.TextMatrix(i - 2, 12), " >>>", ">>>")
            If InStr(1, flxImport.TextMatrix(i - 2, 12), " >>>", vbTextCompare) Then
                flxImport.TextMatrix(i - 2, 13) = flxImport.TextMatrix(i - 2, 13) & "[Step Description=PARAMETER FORMAT FAIL]"
                tmpSts = tmpSts + 1
            End If
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
Exit Sub
ErrLoad:
MsgBox "There was an error while importing the file. Please refresh and close all excel and try again" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub cmdUpload_Click()
Dim tmpR
If Trim(flxImport.TextMatrix(1, 1)) <> "" Then
    If MsgBox("Are you sure you want to upload this to HPQC?", vbYesNo) = vbYes Then
        HasUploadIssue = 0
        If HasIssue = True Then
            If MsgBox("There are some issues found in the upload sheet. Do you want to proceed?", vbYesNo) = vbYes Then
                If MsgBox("Are you really sure that you want to upload this sheet? There are some issues found in the upload sheet. Do you want to proceed?", vbYesNo) = vbYes Then
                    Randomize: tmpR = CInt(Rnd(1000) * 10000)
                    If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
                        Start
                    Else
                        MsgBox "Invalid pass key", vbCritical
                    End If
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
With Me
    .txtPath = "Components\1Company Manual Testing\300 Unit Test\"
    .txtStartDate = Format(Now, "dd/mm/yyyy")
    .txtEndDate = Format(Now + 7, "dd/mm/yyyy")
    .txtQAReviewer = LCase(curUser)
    .txtStatus = "040 Ready For QA Review"
End With
 Me.Caption = Me.Tag
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.Cols = 14
flxImport.TextMatrix(0, 0) = "HPQC Folder Path"
flxImport.TextMatrix(0, 1) = "HPQC Business Component Name"
flxImport.TextMatrix(0, 2) = "Status"
flxImport.TextMatrix(0, 3) = "QA Reviewer"
flxImport.TextMatrix(0, 4) = "Test Type"
flxImport.TextMatrix(0, 5) = "Planned Scripting Start Date"
flxImport.TextMatrix(0, 6) = "Planned Scripting End Date"
flxImport.TextMatrix(0, 7) = "Business Component Name"
flxImport.TextMatrix(0, 8) = "Scripter"
flxImport.TextMatrix(0, 9) = "Peer Reviewer"
flxImport.TextMatrix(0, 10) = "Step Number"
flxImport.TextMatrix(0, 11) = "Step Description"
flxImport.TextMatrix(0, 12) = "Expected Results"
flxImport.TextMatrix(0, 13) = "Validation Check"
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
  curTab = "BC_STEPS-01"
  xlObject.Sheets("Sheet1").Name = curTab
  flxImport.FixedCols = 0
  flxImport.FixedRows = 0
  flxImport.row = 0
  flxImport.col = 0
  Pause 1
  flxImport.RowSel = flxImport.Rows - 1
  flxImport.ColSel = flxImport.Cols - 1
  Clipboard.Clear
  For i = 1 To 5
    flxImport.Clip = Replace(flxImport.Clip, vbCrLf, "<br>", 1, , vbTextCompare)
    flxImport.Clip = Replace(flxImport.Clip, vbNewLine, "<br>", 1, , vbTextCompare)
    flxImport.Clip = Replace(flxImport.Clip, Chr(10) & Chr(13), "<br>", 1, , vbTextCompare)
    flxImport.Clip = Replace(flxImport.Clip, Chr(10), "<br>", 1, , vbTextCompare)
    flxImport.Clip = Replace(flxImport.Clip, Chr(13), "<br>", 1, , vbTextCompare)
    flxImport.Clip = Replace(flxImport.Clip, vbCr, "<br>", 1, , vbTextCompare)
    flxImport.Clip = Replace(flxImport.Clip, vbLf, "<br>", 1, , vbTextCompare)
  Next
  Clipboard.SetText flxImport.Clip
  flxImport.FixedCols = 1
  flxImport.FixedRows = 1

  xlObject.Sheets(curTab).Range("A1").Select
  xlObject.Sheets(curTab).Paste

'On Error Resume Next
    xlObject.Sheets(curTab).Range("N1").Value = "Validation"
    xlObject.Sheets(curTab).Range("A:N").Select

    xlObject.Sheets(curTab).Range("A:N").Borders(xlDiagonalDown).LineStyle = xlNone
    xlObject.Sheets(curTab).Range("A:N").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlObject.Sheets(curTab).Range("A:N").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:N").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:N").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:N").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:N").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:N").Borders(xlInsideHorizontal)
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
    xlObject.Sheets(curTab).Range("A:N").Select
    xlObject.Sheets(curTab).Range("A:N").EntireColumn.AutoFit
    xlObject.Sheets(curTab).Range("A1").Select

    xlObject.Sheets(curTab).Range("A1").AddComment
    xlObject.Sheets(curTab).Range("A1").Comment.Visible = False
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & mdiMain.Caption & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""

    xlObject.Sheets(curTab).Range("N:N").Interior.ColorIndex = 3
    'xlObject.Sheets(curTab).Protection.AllowEditRanges.Add Title:="Range1", Range:=xlObject.Sheets(curTab).Range("A:M")
    'xlObject.Sheets(curTab).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    xlObject.Sheets(curTab).Rows("1:1").Select
    xlObject.Sheets(curTab).Rows("1:1").Insert Shift:=xlDown
    xlObject.Sheets(curTab).Range("A1:G1").Select
    With xlObject.Sheets(curTab).Range("A1:G1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    xlObject.Sheets(curTab).Range("A1:G1").Merge
    xlObject.Sheets(curTab).Range("A1").Interior.ColorIndex = 3
    xlObject.Sheets(curTab).Range("A1").Value = "To Be Populated"
    xlObject.Sheets(curTab).Range("H1:N1").Select
    With xlObject.Sheets(curTab).Range("H1:N1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    xlObject.Sheets(curTab).Range("H1:N1").Merge
    xlObject.Sheets(curTab).Range("H1").Value = "To Be Populated"
    xlObject.Sheets(curTab).Range("H1").Interior.ColorIndex = 5
  xlObject.Workbooks(1).SaveAs "BC_STEPS-01" & "-" & Format(Now, "mmddyyyy HHMMSS AMPM")
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate
  FXGirl.EZPlay FXExportToExcel
  Set xlWB = Nothing
  Set xlObject = Nothing

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

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.Key = "cmdByColor" Then
    MsgBox "Under construction"
ElseIf ButtonMenu.Key = "cmdByParams" Then
    MsgBox "Under construction"
End If
End Sub

Private Function CreateComponentFolder(FolderName As String)
    Dim generalComponentFolderFactory As ComponentFolderFactory
    Dim rootComponentFolderFactory As ComponentFolderFactory
    Dim cFolder As ComponentFolder, rootCFolder As ComponentFolder
  Dim i
  Dim strFol
  Dim X
  Dim stru
  Dim strPath
    Set generalComponentFolderFactory = QCConnection.ComponentFolderFactory
  '*****SPLITTING THE PATH*****
    strFol = Split(FolderName, "\")
    stru = UBound(strFol)
    strPath = "Components\"
  For X = 1 To stru
    Set rootCFolder = generalComponentFolderFactory.FolderByPath(strPath)
    Set rootComponentFolderFactory = rootCFolder.ComponentFolderFactory
    On Error Resume Next
    Set cFolder = rootComponentFolderFactory.AddItem(Null)
    cFolder.Name = strFol(X)
    Debug.Print Err.Number
    If Err.Number = "-2147220199" Then
        On Error GoTo 0
    End If
    On Error Resume Next
    cFolder.Post
    Debug.Print Err.Number
    If Err.Number = "-2147214502" Then
        On Error GoTo 0
    End If
    strPath = strPath & strFol(X) & "\"
  Next
  CreateComponentFolder = True
End Function

