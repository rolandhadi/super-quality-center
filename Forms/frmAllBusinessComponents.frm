VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAllBusinessComponents 
   Caption         =   "Download/Upload Business Component Steps Module"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12705
   Icon            =   "frmAllBusinessComponents.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   12705
   Tag             =   "Download/Upload Business Component Steps Module"
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFilter 
      Height          =   375
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Filter by Test Instance (L4)"
      Top             =   600
      Width           =   4335
   End
   Begin VB.CheckBox chkCSV 
      Caption         =   "Download to CSV"
      Height          =   315
      Left            =   1980
      TabIndex        =   6
      Top             =   120
      Width           =   1635
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
      Left            =   4560
      Picture         =   "frmAllBusinessComponents.frx":08CA
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
      Width           =   12705
      _ExtentX        =   22410
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
            Object.ToolTipText     =   "Upload to HPQC"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin VB.CheckBox chkDeleteOnly 
         Caption         =   "Step Delete Only"
         Height          =   315
         Left            =   3900
         TabIndex        =   7
         Top             =   120
         Width           =   2655
      End
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
            Picture         =   "frmAllBusinessComponents.frx":1070
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   21184
            MinWidth        =   17639
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
            Picture         =   "frmAllBusinessComponents.frx":15C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllBusinessComponents.frx":1853
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllBusinessComponents.frx":1AE5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView QCTree 
      Height          =   4935
      Left            =   60
      TabIndex        =   0
      Top             =   1020
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8705
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
      Height          =   4935
      Left            =   4560
      TabIndex        =   3
      Top             =   1020
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   8705
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
            Picture         =   "frmAllBusinessComponents.frx":1D73
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllBusinessComponents.frx":2485
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllBusinessComponents.frx":2B97
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllBusinessComponents.frx":32A9
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
   Begin MSComctlLib.ImageList imgList_Sts 
      Left            =   11940
      Top             =   540
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
            Picture         =   "frmAllBusinessComponents.frx":39BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllBusinessComponents.frx":3C9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllBusinessComponents.frx":41EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllBusinessComponents.frx":473F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAllBusinessComponents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BC_Component
    Group As String
    Step_ID As Long
    BC_ID As Long
    StepName_Name As String
    Step_Order() As Integer
    Step_Name() As String
    STEP_DESCRIPTION() As String
    Step_ExpectedResult() As String
    AllParams() As String
    Log As String
End Type

Private Type BC_ParAction
    BC_ID As Long
    AddPar() As Variant
    RemovePar() As Variant
    RemoveParID() As Variant
End Type

Private Type BC_PAR_CHANGE
    BC_ID As Long
    BC_PAR As BC_ParAction
End Type

Private All_BC() As BC_Component
Private HasIssue As Boolean
Private HasUploadIssue  As Integer
Private StepsForDeletion()
Private UploadList() As String
Private UploadStepsList() As String
Private UpdateParameters() As BC_PAR_CHANGE

Private Function LoadToArray()
Dim lastrow, i, LastVal, EndArr
Dim StepCounter
lastrow = flxImport.Rows - 1
ReDim All_BC(0)
ReDim StepsForDeletion(0)
EndArr = -1
LastVal = 0
StepCounter = 1

For i = 1 To lastrow
    If Trim(flxImport.TextMatrix(i, 0)) = "" Or Trim(flxImport.TextMatrix(i, 1)) = "" Then
        All_BC(EndArr).Log = All_BC(EndArr).Log & vbCrLf & "Line " & i & " is blank"
    ElseIf Trim(UCase(flxImport.TextMatrix(i, 8))) = "DEL" Then
        StepsForDeletion(UBound(StepsForDeletion)) = flxImport.TextMatrix(i, 0)
        ReDim Preserve StepsForDeletion(UBound(StepsForDeletion) + 1)
    ElseIf (Trim(UCase(flxImport.TextMatrix(i, 8))) <> Trim(UCase(flxImport.TextMatrix(i - 1, 8)))) Then
        EndArr = EndArr + 1
        ReDim Preserve All_BC(EndArr)
        LastVal = LastVal + 1
        All_BC(EndArr).Group = flxImport.TextMatrix(i, 8)
        All_BC(EndArr).Step_ID = flxImport.TextMatrix(i, 0)
        All_BC(EndArr).BC_ID = flxImport.TextMatrix(i, 1)
        
        ReDim All_BC(EndArr).Step_Order(0)
        ReDim All_BC(EndArr).Step_Name(0)
        ReDim All_BC(EndArr).STEP_DESCRIPTION(0)
        ReDim All_BC(EndArr).Step_ExpectedResult(0)
        
        All_BC(EndArr).Step_Order(0) = 1
        All_BC(EndArr).Step_Name(0) = "Step " & UBound(All_BC(EndArr).Step_Name) + 1
        If (Trim(UCase(flxImport.TextMatrix(i, 1))) <> Trim(UCase(flxImport.TextMatrix(i - 1, 1)))) Then
            StepCounter = 1
            All_BC(EndArr).StepName_Name = "Step " & StepCounter
            StepCounter = StepCounter + 1
        Else
            All_BC(EndArr).StepName_Name = "Step " & StepCounter
            StepCounter = StepCounter + 1
        End If
        All_BC(EndArr).STEP_DESCRIPTION(0) = flxImport.TextMatrix(i, 6)
        All_BC(EndArr).Step_ExpectedResult(0) = flxImport.TextMatrix(i, 7)
    Else
        StepsForDeletion(UBound(StepsForDeletion)) = flxImport.TextMatrix(i, 0)
        ReDim Preserve StepsForDeletion(UBound(StepsForDeletion) + 1)
        
        ReDim Preserve All_BC(EndArr).Step_Order(UBound(All_BC(EndArr).Step_Order) + 1)
        ReDim Preserve All_BC(EndArr).Step_Name(UBound(All_BC(EndArr).Step_Name) + 1)
        ReDim Preserve All_BC(EndArr).STEP_DESCRIPTION(UBound(All_BC(EndArr).STEP_DESCRIPTION) + 1)
        ReDim Preserve All_BC(EndArr).Step_ExpectedResult(UBound(All_BC(EndArr).Step_ExpectedResult) + 1)
        
        All_BC(EndArr).Step_Order(UBound(All_BC(EndArr).Step_Order)) = UBound(All_BC(EndArr).Step_Order) + 1
        All_BC(EndArr).Step_Name(UBound(All_BC(EndArr).Step_Name)) = "Step " & UBound(All_BC(EndArr).Step_Name) + 1
        If (Trim(UCase(flxImport.TextMatrix(i, 1))) <> Trim(UCase(flxImport.TextMatrix(i - 1, 1)))) Then
            StepCounter = 1
            All_BC(EndArr).StepName_Name = "Step " & StepCounter
            StepCounter = StepCounter + 1
        Else
            All_BC(EndArr).StepName_Name = "Step " & StepCounter - 1
        End If
        All_BC(EndArr).StepName_Name = "Step " & StepCounter - 1
        All_BC(EndArr).STEP_DESCRIPTION(UBound(All_BC(EndArr).STEP_DESCRIPTION)) = flxImport.TextMatrix(i, 6)
        All_BC(EndArr).Step_ExpectedResult(UBound(All_BC(EndArr).Step_ExpectedResult)) = flxImport.TextMatrix(i, 7)
    End If
Next
End Function

Private Sub GenerateUploadList()
Dim i, j, z, X
Dim tmpDesc, tmpExp, tmp, tmpStep, cFact As ComponentFactory, cComp As Component
Dim compParamFactory As ComponentParamFactory, ALL_BC_PAR() As String
Dim compParam As ComponentParam
Dim compList As List
Dim compList_()
Dim compList_ID()
Dim tmpParam, AddPar(), DelPar()
Dim ParCounter, stringFunct As clsStrings
ReDim UploadList(0)
ReDim UploadStepsList(0)
ReDim UpdateParameters(0)
ReDim ALL_BC_PAR(0)
For i = LBound(All_BC) To UBound(All_BC)
    tmpDesc = ""
    tmpExp = ""
    tmpStep = 1
    If i = 0 Then
        Set cFact = QCConnection.ComponentFactory
        Set cComp = cFact.Item(All_BC(i).BC_ID)
        Set compParamFactory = cComp.ComponentParamFactory
        Set compList = compParamFactory.NewList("") ' HERE!!!
        ReDim compList_(0)
        ReDim compList_ID(0)
        For X = 1 To compList.Count
                compList_(UBound(compList_)) = compList.Item(X).Name
                ReDim Preserve compList_(UBound(compList_) + 1)
                compList_ID(UBound(compList_ID)) = compList.Item(X).ID
                ReDim Preserve compList_ID(UBound(compList_ID) + 1)
        Next
        If i <> 0 Then
            If All_BC(i).BC_ID <> All_BC(i - 1).BC_ID Then ReDim ALL_BC_PAR(0)
        End If
        ReDim All_BC(i).AllParams(0)
    Else
        If All_BC(i).BC_ID <> All_BC(i - 1).BC_ID Then
            Set cFact = QCConnection.ComponentFactory
            Set cComp = cFact.Item(All_BC(i).BC_ID)
            Set compParamFactory = cComp.ComponentParamFactory
            Set compList = compParamFactory.NewList("") ' HERE!!!
            
            ReDim compList_(0)
            ReDim compList_ID(0)
            For X = 1 To compList.Count
                compList_(UBound(compList_)) = compList.Item(X).Name
                ReDim Preserve compList_(UBound(compList_) + 1)
                compList_ID(UBound(compList_ID)) = compList.Item(X).ID
                ReDim Preserve compList_ID(UBound(compList_ID) + 1)
            Next
            If i <> 0 Then
                If All_BC(i).BC_ID <> All_BC(i - 1).BC_ID Then ReDim ALL_BC_PAR(0)
            End If
            ReDim All_BC(i).AllParams(0)
        Else
            All_BC(i).AllParams = All_BC(i - 1).AllParams
        End If
    End If
    For j = LBound(All_BC(i).STEP_DESCRIPTION) To UBound(All_BC(i).STEP_DESCRIPTION)
                tmp = All_BC(i).STEP_DESCRIPTION(j)
                If HasParameters(tmp) = True Then
                    tmpParam = ExtractParametersWithFix(tmp)
                    For X = LBound(tmpParam) To UBound(tmpParam)
                        If Trim(tmpParam(X)) <> "" Then
                            tmp = Replace(tmp, "<<<" & tmpParam(X) & ">>>", LCase("<<<" & tmpParam(X) & ">>>"), , , vbTextCompare)
                            If IsParameterDeclared(All_BC(i).AllParams, tmpParam(X)) = False Then
                                All_BC(i).AllParams(UBound(All_BC(i).AllParams)) = tmpParam(X)
                                ReDim Preserve All_BC(i).AllParams(UBound(All_BC(i).AllParams) + 1)
                                ALL_BC_PAR(UBound(ALL_BC_PAR)) = tmpParam(X)
                                ReDim Preserve ALL_BC_PAR(UBound(ALL_BC_PAR) + 1)
                            End If
                        End If
                    Next
                End If
                For z = 1 To 100
                    tmp = Replace(tmp, vbCrLf, "<br>", , , vbTextCompare)
                Next
                For z = 1 To 100
                    tmp = Replace(tmp, Chr(10) & Chr(13), "<br>", , , vbTextCompare)
                Next
                For z = 1 To 100
                    tmp = Replace(tmp, Chr(10), "<br>", , , vbTextCompare)
                Next
                For z = 1 To 100
                    tmp = Replace(tmp, Chr(13), "<br>", , , vbTextCompare)
                Next
                For z = 1 To 100
                    tmp = Replace(tmp, vbTab, "     ", , , vbTextCompare)
                Next
                If j = UBound(All_BC(i).STEP_DESCRIPTION) Then
                    tmpDesc = tmpDesc & tmp
                Else
                    tmpDesc = tmpDesc & tmp & "<br>"
                End If
                'If Right(tmpDesc, 4) = "<br>" Then tmpDesc = Left(tmpDesc, Len(tmpDesc) - 4)
                tmp = All_BC(i).Step_ExpectedResult(j)
                If HasParameters(tmp) = True Then
                    tmpParam = ExtractParametersWithFix(tmp)
                    For X = LBound(tmpParam) To UBound(tmpParam)
                        If Trim(tmpParam(X)) <> "" Then
                            tmp = Replace(tmp, "<<<" & tmpParam(X) & ">>>", LCase("<<<" & tmpParam(X) & ">>>"), , , vbTextCompare)
                            If IsParameterDeclared(All_BC(i).AllParams, tmpParam(X)) = False Then
                                All_BC(i).AllParams(UBound(All_BC(i).AllParams)) = tmpParam(X)
                                ReDim Preserve All_BC(i).AllParams(UBound(All_BC(i).AllParams) + 1)
                                ALL_BC_PAR(UBound(ALL_BC_PAR)) = tmpParam(X)
                                ReDim Preserve ALL_BC_PAR(UBound(ALL_BC_PAR) + 1)
                            End If
                        End If
                    Next
                End If
                For z = 1 To 100
                    tmp = Replace(tmp, vbCrLf, "<br>", , , vbTextCompare)
                Next
                For z = 1 To 100
                    tmp = Replace(tmp, Chr(10) & Chr(13), "<br>", , , vbTextCompare)
                Next
                For z = 1 To 100
                    tmp = Replace(tmp, Chr(10), "<br>", , , vbTextCompare)
                Next
                For z = 1 To 100
                    tmp = Replace(tmp, Chr(13), "<br>", , , vbTextCompare)
                Next
                For z = 1 To 100
                    tmp = Replace(tmp, vbTab, "     ", , , vbTextCompare)
                Next
                If j = UBound(All_BC(i).STEP_DESCRIPTION) Then
                    tmpExp = tmpExp & tmp
                Else
                    tmpExp = tmpExp & tmp & "<br>"
                End If
                'If Right(tmpExp, 4) = "<br>" Then tmpExp = Left(tmpExp, Len(tmpExp) - 4)
    Next
    ReDim Preserve UploadList(i)
    ReDim Preserve UploadStepsList(i)
    
    If i = 0 Then
        ParCounter = 0
        ReDim UpdateParameters(ParCounter)
        UpdateParameters(ParCounter).BC_ID = All_BC(i).BC_ID
        UpdateParameters(ParCounter).BC_PAR.AddPar() = GetAddPars(compList_, All_BC(i).AllParams)
        For j = LBound(UpdateParameters(ParCounter).BC_PAR.AddPar()) To UBound(UpdateParameters(ParCounter).BC_PAR.AddPar()) - 1
            If IsParameterDeclared(All_BC(i).AllParams(UBound(All_BC(i).AllParams)), UpdateParameters(ParCounter).BC_PAR.AddPar(j)) = False Then
                All_BC(i).AllParams(UBound(All_BC(i).AllParams)) = UpdateParameters(ParCounter).BC_PAR.AddPar(j)
                ReDim Preserve All_BC(i).AllParams(UBound(compList_) + 1)
            End If
        Next
        UpdateParameters(ParCounter).BC_PAR.RemovePar() = GetRemovePars(compList_, ALL_BC_PAR)
        UpdateParameters(ParCounter).BC_PAR.RemoveParID() = GetRemoveParsID(compList_, compList_ID, ALL_BC_PAR)
        tmpStep = CStr(All_BC(i).StepName_Name)
    Else
        If All_BC(i).BC_ID <> All_BC(i - 1).BC_ID Then
            ParCounter = ParCounter + 1
            ReDim Preserve UpdateParameters(ParCounter)
            UpdateParameters(ParCounter).BC_ID = All_BC(i).BC_ID
            UpdateParameters(ParCounter).BC_PAR.AddPar() = GetAddPars(compList_, All_BC(i).AllParams)
            For j = LBound(UpdateParameters(ParCounter).BC_PAR.AddPar()) To UBound(UpdateParameters(ParCounter).BC_PAR.AddPar()) - 1
                If IsParameterDeclared(All_BC(i).AllParams(UBound(All_BC(i).AllParams)), UpdateParameters(ParCounter).BC_PAR.AddPar(j)) = False Then
                    All_BC(i).AllParams(UBound(All_BC(i).AllParams)) = UpdateParameters(ParCounter).BC_PAR.AddPar(j)
                    ReDim Preserve All_BC(i).AllParams(UBound(compList_) + 1)
                End If
            Next
            UpdateParameters(ParCounter).BC_PAR.RemovePar() = GetRemovePars(compList_, ALL_BC_PAR)
            UpdateParameters(ParCounter).BC_PAR.RemoveParID() = GetRemoveParsID(compList_, compList_ID, ALL_BC_PAR)
            tmpStep = CStr(All_BC(i).StepName_Name)
        Else
            ReDim Preserve UpdateParameters(ParCounter)
            UpdateParameters(ParCounter).BC_ID = All_BC(i).BC_ID
            UpdateParameters(ParCounter).BC_PAR.AddPar() = GetAddPars(compList_, All_BC(i).AllParams)
            For j = LBound(UpdateParameters(ParCounter).BC_PAR.AddPar()) To UBound(UpdateParameters(ParCounter).BC_PAR.AddPar()) - 1
                If IsParameterDeclared(All_BC(i).AllParams(UBound(All_BC(i).AllParams)), UpdateParameters(ParCounter).BC_PAR.AddPar(j)) = False Then
                    All_BC(i).AllParams(UBound(All_BC(i).AllParams)) = UpdateParameters(ParCounter).BC_PAR.AddPar(j)
                    ReDim Preserve All_BC(i).AllParams(UBound(compList_) + 1)
                End If
            Next
            UpdateParameters(ParCounter).BC_PAR.RemovePar() = GetRemovePars(compList_, ALL_BC_PAR)
            UpdateParameters(ParCounter).BC_PAR.RemoveParID() = GetRemoveParsID(compList_, compList_ID, ALL_BC_PAR)
            tmpStep = CStr(All_BC(i).StepName_Name)
        End If
    End If
    
    tmpDesc = Replace(tmpDesc, "<br><br><br><br><br>", "<br>", , , vbTextCompare)
    tmpDesc = Replace(tmpDesc, "<br><br><br><br>", "<br>", , , vbTextCompare)
    tmpDesc = Replace(tmpDesc, "<br><br><br>", "<br>", , , vbTextCompare)
    
    tmpExp = Replace(tmpExp, "<br><br><br><br><br>", "<br>", , , vbTextCompare)
    tmpExp = Replace(tmpExp, "<br><br><br><br>", "<br>", , , vbTextCompare)
    tmpExp = Replace(tmpExp, "<br><br><br>", "<br>", , , vbTextCompare)
    
    UploadList(i) = "UPDATE COMPONENT_STEP SET CS_DESCRIPTION = '<html><body>" & CleanHTML_(CStr(tmpDesc)) & "</body></html>', CS_EXPECTED = '<html><body>" & CleanHTML_(CStr(tmpExp)) & "</body></html>' WHERE CS_STEP_ID = " & All_BC(i).Step_ID
    UploadStepsList(i) = "UPDATE COMPONENT_STEP SET CS_STEP_NAME = '" & tmpStep & "' WHERE CS_STEP_ID = " & CStr(All_BC(i).Step_ID)
Next
End Sub

Private Function GetAddPars(BC_PARS, STEP_PAR)
Dim i As Integer, j  As Integer, found As Boolean
Dim tmp
On Error Resume Next
    ReDim tmp(0)
    For i = LBound(STEP_PAR) To UBound(STEP_PAR)
        For j = LBound(BC_PARS) To UBound(BC_PARS)
            If Trim(UCase(STEP_PAR(i))) = Trim(UCase(BC_PARS(j))) Then
                found = True
                Exit For
            End If
        Next
        If found = False Then
            If IsParameterDeclared(tmp, LCase(STEP_PAR(i))) = False Then
             tmp(UBound(tmp)) = LCase(STEP_PAR(i))
            ReDim Preserve tmp(UBound(tmp) + 1)
            End If
        End If
        found = False
    Next
    GetAddPars = tmp
End Function

Private Function GetRemovePars(BC_PARS, STEP_PAR)
Dim i As Integer, j  As Integer, found As Boolean
Dim tmp
On Error Resume Next
    ReDim tmp(0)
    For i = LBound(BC_PARS) To UBound(BC_PARS)
        For j = LBound(STEP_PAR) To UBound(STEP_PAR)
            If Trim(UCase(BC_PARS(i))) = Trim(UCase(STEP_PAR(j))) Then
                found = True
                Exit For
            End If
        Next
        If found = False Then
            tmp(UBound(tmp)) = LCase(BC_PARS(i))
            ReDim Preserve tmp(UBound(tmp) + 1)
        End If
        found = False
    Next
    GetRemovePars = tmp
End Function

Private Function GetRemoveParsID(BC_PARS, BC_PARS_ID, STEP_PAR)
Dim i As Integer, j  As Integer, found As Boolean
Dim tmp
On Error Resume Next
    ReDim tmp(0)
    For i = LBound(BC_PARS) To UBound(BC_PARS)
        For j = LBound(STEP_PAR) To UBound(STEP_PAR)
            If Trim(UCase(BC_PARS(i))) = Trim(UCase(STEP_PAR(j))) Then
                found = True
                Exit For
            End If
        Next
        If found = False Then
            tmp(UBound(tmp)) = LCase(BC_PARS_ID(i))
            ReDim Preserve tmp(UBound(tmp) + 1)
        End If
        found = False
    Next
    GetRemoveParsID = tmp
End Function


Function LoadToQC()
Dim i, j, k
Dim tmpComp, objCommand, rs
Dim cFact As ComponentFactory, cComp As Component
Dim compParamFactory As ComponentParamFactory
Dim compParam() As ComponentParam

If chkDeleteOnly.Value = Unchecked Then
  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = ""
  mdiMain.pBar.Max = UBound(All_BC) + 3
  For i = LBound(All_BC) To UBound(All_BC)
      On Error Resume Next
          If UploadList(i) = "" Then Exit For
          Set objCommand = QCConnection.Command
          objCommand.CommandText = UploadList(i)
          Set rs = objCommand.Execute
      If Err.Number <> 0 Then
          FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[UPDATE STEPS: (FAILED) " & Now & " " & All_BC(i).Group & "-" & All_BC(i).Step_ID & "] " & Err.Description
          HasUploadIssue = HasUploadIssue + 1
          Err.Clear
          On Error GoTo 0
      Else
          FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[UPDATE STEPS: (PASSED) " & Now & " " & All_BC(i).Group & "-" & All_BC(i).Step_ID & "] "
      End If
      
      On Error Resume Next
      If UploadStepsList(i) = "" Then Exit For
          Set objCommand = QCConnection.Command
          objCommand.CommandText = UploadStepsList(i)
          Set rs = objCommand.Execute
      If Err.Number <> 0 Then
          FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[UPDATE STEPS: (FAILED) " & Now & " " & All_BC(i).Group & "-" & All_BC(i).Step_ID & "] " & Err.Description
          HasUploadIssue = HasUploadIssue + 1
          Err.Clear
          On Error GoTo 0
      Else
          FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[UPDATE STEPS: (PASSED) " & Now & " " & All_BC(i).Group & "-" & All_BC(i).Step_ID & "]"
      End If
      
      stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Updating Business Component Steps " & i + 1 & " out of " & UBound(All_BC) + 1 & " (" & All_BC(i).Step_ID & ")"
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
  FXGirl.EZPlay FXDataUploadCompleted
  mdiMain.pBar.Value = mdiMain.pBar.Max
End If
If chkDeleteOnly.Value = Checked Then
  mdiMain.pBar.Max = UBound(StepsForDeletion) + 1
  For i = LBound(StepsForDeletion) To UBound(StepsForDeletion)
      On Error Resume Next
      If StepsForDeletion(i) = "" Then Exit For
          Set objCommand = QCConnection.Command
          objCommand.CommandText = "DELETE FROM COMPONENT_STEP WHERE CS_STEP_ID = " & StepsForDeletion(i)
          Set rs = objCommand.Execute
      If Err.Number <> 0 Then
          FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[UPDATE STEPS: (FAILED)" & Now & " " & All_BC(i).Group & "-" & All_BC(i).Step_ID & "] " & Err.Description
          HasUploadIssue = HasUploadIssue + 1
          Err.Clear
          On Error GoTo 0
      Else
          FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[UPDATE STEPS: (PASSED)" & Now & " " & All_BC(i).Group & "-" & All_BC(i).Step_ID & "]"
      End If
      mdiMain.pBar.Value = i + 1
      Debug.Print mdiMain.pBar.Value
  Next
End If

If chkDeleteOnly.Value = Unchecked Then
  For i = LBound(UpdateParameters) To UBound(UpdateParameters)
      On Error Resume Next
      If UpdateParameters(i).BC_ID = 0 Then Exit For
      Set cFact = QCConnection.ComponentFactory
      Set cComp = cFact.Item(UpdateParameters(i).BC_ID)
      Set compParamFactory = cComp.ComponentParamFactory
      For j = LBound(UpdateParameters(i).BC_PAR.RemovePar) To UBound(UpdateParameters(i).BC_PAR.RemovePar)
          If Trim(UpdateParameters(i).BC_PAR.RemoveParID(j)) <> "" Or UpdateParameters(i).BC_PAR.RemoveParID(j) <> 0 Then
              compParamFactory.RemoveItem UpdateParameters(i).BC_PAR.RemoveParID(j)
              FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[REMOVE BC PARAMETERS: (PASSED)" & Now & " " & All_BC(i).BC_ID & "-" & All_BC(i).Step_ID & "-" & CStr(UpdateParameters(i).BC_PAR.RemovePar(j)) & "]"
          End If
      Next
  Next
  
  For i = LBound(UpdateParameters) To UBound(UpdateParameters)
      On Error Resume Next
      If UpdateParameters(i).BC_ID = 0 Then Exit For
      Set cFact = QCConnection.ComponentFactory
      Set cComp = cFact.Item(UpdateParameters(i).BC_ID)
      Set compParamFactory = cComp.ComponentParamFactory
      For j = LBound(UpdateParameters(i).BC_PAR.AddPar) To UBound(UpdateParameters(i).BC_PAR.AddPar)
          If Trim(UpdateParameters(i).BC_PAR.AddPar(j)) <> "" Or UpdateParameters(i).BC_PAR.AddPar(j) <> 0 Then
              ReDim Preserve compParam(j + 1)
              Set compParam(j) = compParamFactory.AddItem(Null)
              If UCase(Left(UpdateParameters(i).BC_PAR.AddPar(j), 1)) = "O" Then
                  compParam(j).IsOut = 1
              Else
                  compParam(j).IsOut = 0
              End If
              compParam(j).Name = Replace(LCase(UpdateParameters(i).BC_PAR.AddPar(j)), "-", "_")
              compParam(j).Desc = LCase(UpdateParameters(i).BC_PAR.AddPar(j))
              compParam(j).ValueType = "String"
              compParam(j).Post
              FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[ADD BC PARAMETERS: (PASSED)" & Now & " " & All_BC(i).BC_ID & "-" & All_BC(i).Step_ID & "-" & compParam(j).Name & "]"
          End If
      Next
  Next
End If

mdiMain.pBar.Value = mdiMain.pBar.Max
stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = UBound(All_BC) + 1 & " Business Component Step(s) updated successfully. (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log"
QCConnection.SendMail "user@companyemail.com", "", "[HPQC UPDATES] Business Component Step(s) updated successfully by " & curUser & " in " & curDomain & "-" & curProject, UBound(All_BC) + 1 & " Business Component Step(s) updated successfully. (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
QCConnection.SendMail curUser, "", "[HPQC UPDATES] Business Component Step(s) updated successfully by " & curUser & " in " & curDomain & "-" & curProject, UBound(All_BC) + 1 & " Business Component Step(s) updated successfully. (" & HasUploadIssue & ") uploading issue(s) found. See " & App.path & "\SQC DAT" & "\" & Format(Now, "mm-dd-yyyy") & ".log" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
End Function

Sub Start()
Debug.Print "New Session: " & Now
LoadToArray
If chkDeleteOnly.Value = Unchecked Then
  GenerateUploadList
End If
LoadToQC
Debug.Print "New Finished: " & Now
End Sub


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
Dim z
Dim mySTEP_ID
Dim myBC_ID
Dim myFOLDER_NAME
Dim myCOMPONENT_NAME
Dim mySTEP_ORDER
Dim mySTEP_NAME
Dim myDESCRIPTION
Dim myEXPECTED
Dim LastBCID, LastBCFolder

    TimeSt = Format(Now, "mmm-dd-yyyy hhmmss")
    If chkCSV.Value = Checked Then
      AllF = InputBox("Enter file name", "File name", "[BC Steps] ")
    Else
      AllF = "[BC Steps] "
    End If
    
    ReDim CheckedItems(0): strPath = ""
    GetAllCheckedItems QCTree.Nodes(1)
    For j = LBound(CheckedItems) To UBound(CheckedItems) - 1
        If Left(CheckedItems(j), 1) = "F" Then
            strPath = strPath & "FC_PATH LIKE '" & GetFromTable(Right(CheckedItems(j), Len(CheckedItems(j)) - 1), "FC_ID", "FC_PATH", "COMPONENT_FOLDER") & "%' OR "
        ElseIf Left(CheckedItems(j), 1) = "C" Then
            strPath = strPath & "CO_ID = " & Right(CheckedItems(j), Len(CheckedItems(j)) - 1) & " OR "
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
    If Trim(txtFilter.Text) <> "" Then
      objCommand.CommandText = "SELECT CS_STEP_ID, FC_NAME, CO_ID, CO_NAME, CS_STEP_ORDER, CS_STEP_NAME, CS_DESCRIPTION, CS_EXPECTED FROM COMPONENT_STEP, COMPONENT, COMPONENT_FOLDER WHERE CO_ID = CS_COMPONENT_ID AND CO_FOLDER_ID = FC_ID AND " & strPath & " AND CO_NAME LIKE '%" & txtFilter.Text & "%' ORDER BY FC_NAME, CO_ID, CS_STEP_ORDER"
    Else
      objCommand.CommandText = "SELECT CS_STEP_ID, FC_NAME, CO_ID, CO_NAME, CS_STEP_ORDER, CS_STEP_NAME, CS_DESCRIPTION, CS_EXPECTED FROM COMPONENT_STEP, COMPONENT, COMPONENT_FOLDER WHERE CO_ID = CS_COMPONENT_ID AND CO_FOLDER_ID = FC_ID AND " & strPath & " ORDER BY FC_NAME, CO_ID, CS_STEP_ORDER"
    End If
    Debug.Print Me.Caption & "-" & objCommand.CommandText
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[SQL] " & Now & " " & objCommand.CommandText
    Set rs = objCommand.Execute                                                                                                                                                                                                                                                 'HERE!!!!!! <<<-------------
    If rs.RecordCount > 10000 And chkCSV.Value = Unchecked Then
        MsgBox "The records found exceeds 2500 records. It will be automatically generated as a CSV file.", vbOKOnly
        chkCSV.Value = Checked
        Exit Sub
        GenerateOutput
    End If
    AllScript = """" & "Step ID" & """" & "," & """" & "Component ID" & """" & "," & """" & "Test Set Folder" & """" & "," & """" & "Component Name" & """" & "," & """" & "Step Order" & """" & "," & """" & "Step Name" & """" & "," & """" & "Description" & """" & "," & """" & "Expected" & """" & "," & """" & "Group" & """" & "," & """" & "Validation" & """"
    ClearTable
    If chkCSV.Value = Unchecked Then
        flxImport.Rows = rs.RecordCount + 1
    End If
        k = 0
        mdiMain.pBar.Max = rs.RecordCount + 3
    For i = 1 To rs.RecordCount
        'If i = 1 Then FileWrite App.path & "\SQC DAT" & "\" & QCTree.SelectedItem.Key & "-" & TimeSt & AllF & ".xls", AllScript
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Processing " & i & " out of " & rs.RecordCount
                
                mySTEP_ID = rs.FieldValue("CS_STEP_ID")
                myBC_ID = rs.FieldValue("CO_ID")
        
                If LastBCID <> rs.FieldValue("CO_ID") Then
                    myFOLDER_NAME = GetBusinessComponentFolderPath(rs.FieldValue("CO_ID"))
                    LastBCID = rs.FieldValue("CO_ID")
                    LastBCFolder = myFOLDER_NAME
                End If
                
                myFOLDER_NAME = LastBCFolder
                
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
                myDESCRIPTION = Replace(myDESCRIPTION, """", "'", , , vbTextCompare)
                myDESCRIPTION = Trim(myDESCRIPTION)
                If Left(myDESCRIPTION, 1) = "-" Then myDESCRIPTION = Right(myDESCRIPTION, Len(myDESCRIPTION) - 1)
                
                For z = 1 To 10
                    myDESCRIPTION = Replace(myDESCRIPTION, vbCrLf, "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 10
                    myDESCRIPTION = Replace(myDESCRIPTION, Chr(10) & Chr(13), "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 10
                    myDESCRIPTION = Replace(myDESCRIPTION, Chr(10), "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 10
                    myDESCRIPTION = Replace(myDESCRIPTION, Chr(13), "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 10
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
                myEXPECTED = Replace(myEXPECTED, """", "'", , , vbTextCompare)
                myEXPECTED = Trim(myEXPECTED)
                If Left(myEXPECTED, 1) = "-" Then myEXPECTED = Right(myEXPECTED, Len(myEXPECTED) - 1)
                
                For z = 1 To 10
                    myEXPECTED = Replace(myEXPECTED, vbCrLf, "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 10
                    myEXPECTED = Replace(myEXPECTED, Chr(10) & Chr(13), "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 10
                    myEXPECTED = Replace(myEXPECTED, Chr(10), "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 10
                    myEXPECTED = Replace(myEXPECTED, Chr(13), "<br>", , , vbTextCompare)
                Next
                
                For z = 1 To 10
                    myEXPECTED = Replace(myEXPECTED, vbTab, "     ", , , vbTextCompare)
                Next
                        
                If chkCSV.Value = Unchecked Then
                    k = k + 1
                    flxImport.Rows = k + 1
                    flxImport.TextMatrix(k, 0) = mySTEP_ID
                    flxImport.TextMatrix(k, 1) = myBC_ID
                    flxImport.TextMatrix(k, 2) = myFOLDER_NAME
                    flxImport.TextMatrix(k, 3) = myCOMPONENT_NAME
                    flxImport.TextMatrix(k, 4) = mySTEP_ORDER
                    flxImport.TextMatrix(k, 5) = mySTEP_NAME
                    flxImport.TextMatrix(k, 6) = myDESCRIPTION
                    flxImport.TextMatrix(k, 7) = myEXPECTED
                    flxImport.TextMatrix(k, 8) = k
                Else
                    If Trim(AllScript) <> "" Then
                        AllScript = AllScript & vbCrLf & """" & mySTEP_ID & """" & "," & """" & myBC_ID & """" & "," & """" & myFOLDER_NAME & """" & "," & """" & myCOMPONENT_NAME & """" & "," & """" & mySTEP_ORDER & """" & "," & """" & mySTEP_NAME & """" & "," & """" & Replace(myDESCRIPTION, "<br>", vbCrLf) & """" & "," & """" & Replace(myEXPECTED, "<br>", vbCrLf) & """"
                    Else
                        AllScript = AllScript & """" & mySTEP_ID & """" & "," & """" & myBC_ID & """" & "," & """" & myFOLDER_NAME & """" & "," & """" & myCOMPONENT_NAME & """" & "," & """" & mySTEP_ORDER & """" & "," & """" & mySTEP_NAME & """" & "," & """" & Replace(myDESCRIPTION, "<br>", vbCrLf) & """" & "," & """" & Replace(myEXPECTED, "<br>", vbCrLf) & """"
                    End If
                End If
        If chkCSV.Value = Checked Then
            If i Mod 500 = 0 Then
                FileAppend App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv", AllScript
                AllScript = ""
            End If
        End If
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
        rs.Next
    Next
    FXGirl.EZPlay FXSQCExtractCompleted
    mdiMain.pBar.Value = mdiMain.pBar.Max
If chkCSV.Value = Checked Then '***
    FileAppend App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv", AllScript: If MsgBox("Successfully exported to " & App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv" & vbCrLf & "Do you want to launch the extracted file?", vbYesNo) = vbYes Then Shell "explorer.exe " & App.path & "\SQC Logs" & "\", vbNormalFocus
    AllScript = vbCrLf & ","
    AllScript = AllScript & """" & "SQL Code:" & """" & "," & """" & Replace(objCommand.CommandText, """", "'") & """"
    FileAppend App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv", AllScript
End If '***
stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = mdiMain.pBar.Max & " record(s) generated"
End Sub


Private Sub chkCSV_Click()
If chkCSV.Value = Checked Then
    If MsgBox("Are you sure you want to download directly to CSV?", vbYesNo) = vbYes Then
        chkCSV.Value = Checked
    Else
        chkCSV.Value = Unchecked
    End If
End If
End Sub

Private Sub chkDeleteOnly_Click()
If chkDeleteOnly.Value = Checked Then
    If MsgBox("Are you sure you want to select the Step Delete Only?", vbYesNo) = vbYes Then
        chkDeleteOnly.Value = Checked
    Else
        chkDeleteOnly.Value = Unchecked
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
Dim curLetterA, curLetterB
HasIssue = False
curLetterA = 1
curLetterB = 1
On Error Resume Next
    xlWB.Close
    xlObject.Application.Quit
On Error GoTo 0
On Error GoTo ErrLoad
    dlgOpenExcel.filename = "": dlgOpenExcel.ShowOpen
    fname = dlgOpenExcel.filename: stsBar.Tag = fname
    If fname = "" Then Exit Sub Else Me.Caption = Me.Caption & " (" & dlgOpenExcel.FileTitle & ")"
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Open(fname) 'Open your book here
                
    Clipboard.Clear

    With xlObject.ActiveWorkbook.ActiveSheet
            Debug.Print xlObject.ActiveWorkbook.Sheets(2).Range("B7").Value
         If UCase(Trim(curDomain & "-" & curProject)) <> UCase(Trim(xlObject.ActiveWorkbook.Sheets(2).Range("B7").Value)) Then
            MsgBox "The spreadsheet is from a different Domain or Project"
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
         If UCase(Trim(.Range("A1").Value)) <> UCase(Trim("Step ID")) Then
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
        flxImport.Cols = 10
        flxImport.Redraw = False
        
        'A - Load HPQC Folder Path
        'Should not be blank
        For i = 2 To lastrow
            
            flxImport.TextMatrix(i - 1, 0) = strFunct.RemoveBackslash(((.Range("A" & i).Value)))          'Change number and letter
            If Trim(.Range("A" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step ID=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 1) = strFunct.RemoveBackslash(((.Range("B" & i).Value)))          'Change number and letter
            If Trim(.Range("B" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step ID=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 2) = strFunct.RemoveBackslash(((.Range("C" & i).Value)))          'Change number and letter
            If Trim(.Range("C" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Test Set Folder=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 3) = strFunct.RemoveBackslash(((.Range("D" & i).Value)))          'Change number and letter
            If Trim(.Range("D" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Component Name=BLANK]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 4) = strFunct.RemoveBackslash(((.Range("E" & i).Value)))          'Change number and letter
            flxImport.TextMatrix(i - 1, 5) = strFunct.RemoveBackslash(((.Range("F" & i).Value)))          'Change number and letter
            
            flxImport.TextMatrix(i - 1, 6) = Trim(.Range("G" & i).Value)    'Change number and letter
            If Trim(.Range("F" & i).Value) = "" Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step Description=BLANK]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, .Range("G" & i).Value, "<<<<", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step Description=<<<< FOUND]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, .Range("G" & i).Value, ">>>>", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step Description=>>>> FOUND]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, UCase(.Range("G" & i).Value), " <<P_", vbTextCompare) <> 0 Or InStr(1, UCase(.Range("F" & i).Value), vbCrLf & "<<P_", vbTextCompare) <> 0 Or InStr(1, UCase(.Range("F" & i).Value), " <<O_", vbTextCompare) <> 0 Or InStr(1, UCase(.Range("F" & i).Value), vbCrLf & "<<O_", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step Description=<< FOUND]"
                tmpSts = tmpSts + 1
            End If
            For j = 1 To 26
                If InStr(1, .Range("G" & i).Value, "<<<" & Chr(j + 64), vbTextCompare) <> 0 And (LCase(Chr(j + 64)) <> "p" And LCase(Chr(j + 64)) <> "o") Then
                    flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step Description=PARAMETER FORMAT FAIL]"
                    tmpSts = tmpSts + 1
                End If
            Next
            If HasParameters(.Range("G" & i).Value) = True Then
                tmpParam = ExtractParameters(.Range("G" & i).Value)
                For j = LBound(tmpParam) To UBound(tmpParam)
                    If InvalidParameterCheck(tmpParam(j)) = True Then
                        flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Parameter=INVALID FORMAT/CHAR]"
                        tmpSts = tmpSts + 1
                    End If
                Next
            End If
            If InStr(1, .Range("G" & i).Value, "<<< ", vbTextCompare) Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step Description=PARAMETER FORMAT FAIL]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, .Range("G" & i).Value, " >>>", vbTextCompare) Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step Description=PARAMETER FORMAT FAIL]"
                tmpSts = tmpSts + 1
            End If
            
            flxImport.TextMatrix(i - 1, 7) = Trim(.Range("H" & i).Value)    'Change number and letter
            If InStr(1, .Range("G" & i).Value, "<<<<", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step Description=<<<< FOUND]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, .Range("H" & i).Value, ">>>>", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step Description=>>>> FOUND]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, UCase(.Range("H" & i).Value), " <<P_", vbTextCompare) <> 0 Or InStr(1, UCase(.Range("G" & i).Value), vbCrLf & "<<P_", vbTextCompare) <> 0 Or InStr(1, UCase(.Range("G" & i).Value), " <<O_", vbTextCompare) <> 0 Or InStr(1, UCase(.Range("G" & i).Value), vbCrLf & "<<O_", vbTextCompare) <> 0 Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step Description=<< FOUND]"
                tmpSts = tmpSts + 1
            End If
            For j = 1 To 26
                If InStr(1, .Range("H" & i).Value, "<<<" & Chr(j + 64), vbTextCompare) <> 0 And (LCase(Chr(j + 64)) <> "p" And LCase(Chr(j + 64)) <> "o") Then
                    flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step Description=PARAMETER FORMAT FAIL]"
                    tmpSts = tmpSts + 1
                End If
            Next
            If HasParameters(.Range("H" & i).Value) = True Then
                tmpParam = ExtractParameters(.Range("H" & i).Value)
                For j = LBound(tmpParam) To UBound(tmpParam)
                    If InvalidParameterCheck(tmpParam(j)) = True Then
                        flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Parameter=INVALID FORMAT/CHAR]"
                        tmpSts = tmpSts + 1
                    End If
                Next
            End If
            If InStr(1, .Range("H" & i).Value, "<<< ", vbTextCompare) Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step Description=PARAMETER FORMAT FAIL]"
                tmpSts = tmpSts + 1
            End If
            If InStr(1, .Range("H" & i).Value, " >>>", vbTextCompare) Then
                flxImport.TextMatrix(i - 1, 9) = flxImport.TextMatrix(i - 1, 9) & "[Step Description=PARAMETER FORMAT FAIL]"
                tmpSts = tmpSts + 1
            End If
            If .Range("G" & i).Interior.color = "65280" Or .Range("G" & i).Interior.color = "10147522" Then
                If flxImport.TextMatrix(i - 1, 3) = flxImport.TextMatrix(i - 2, 3) Then
                    flxImport.TextMatrix(i - 1, 8) = "A" & (curLetterA)
                Else
                    curLetterA = curLetterA + 1
                    flxImport.TextMatrix(i - 1, 8) = "A" & (curLetterA)
                End If
            ElseIf .Range("G" & i).Interior.color = "6723891" Or .Range("G" & i).Interior.color = "3969653" Then
               If flxImport.TextMatrix(i - 1, 3) = flxImport.TextMatrix(i - 2, 3) Then
                    flxImport.TextMatrix(i - 1, 8) = "B" & (curLetterB)
                Else
                    curLetterB = curLetterB + 1
                    flxImport.TextMatrix(i - 1, 8) = "B" & (curLetterB)
                End If
            ElseIf .Range("G" & i).Interior.color = "26367" Or .Range("G" & i).Interior.color = "683492" Or .Range("G" & i).Interior.color = "5066944" Then
                flxImport.TextMatrix(i - 1, 8) = "DEL"
            Else
                flxImport.TextMatrix(i - 1, 8) = i - 1
            End If
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = i - 1 & " out of " & lastrow - 1 & " validated " & Format(i / lastrow, "0.0%") & " (" & tmpSts & ") errors found."
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
Exit Sub
ErrLoad:
MsgBox "There was an error while importing the file. Please refresh and close all excel and try again" & vbCrLf & Err.Description, vbCritical
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
    
    If Left(QCTree.SelectedItem.Key, 1) = "F" Then
        Set objCommand = QCConnection.Command
        objCommand.CommandText = "SELECT CO_ID, CO_NAME FROM COMPONENT, COMPONENT_FOLDER WHERE CO_FOLDER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " AND CO_FOLDER_ID = FC_ID ORDER BY CO_NAME"
        Set rs = objCommand.Execute
        For i = 1 To rs.RecordCount
            QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("C" & rs.FieldValue("CO_ID")), rs.FieldValue("CO_NAME"), 3
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
    If Trim(flxImport.TextMatrix(1, 0)) <> "" Then
        If IncorrectHeaderDetails = False Then
            If MsgBox("Are you sure you want to upload this to HPQC?", vbYesNo) = vbYes Then
                HasUploadIssue = 0
                If HasIssue = True Then
                    If MsgBox("There are some issues found in the upload sheet. Do you want to proceed?", vbYesNo) = vbYes Then
                        Randomize: tmpR = CInt(Rnd(1000) * 10000)
                        If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
                            If MsgBox("Are you really sure that you want to upload this sheet? There are some issues found in the upload sheet. Do you want to proceed?", vbYesNo) = vbYes Then Start
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
    Me.Caption = Me.Tag
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.TextMatrix(0, 0) = "Step ID"
flxImport.TextMatrix(0, 1) = "Component ID"
flxImport.TextMatrix(0, 2) = "Test Set Folder"
flxImport.TextMatrix(0, 3) = "Component Name"
flxImport.TextMatrix(0, 4) = "Step Order"
flxImport.TextMatrix(0, 5) = "Step Name"
flxImport.TextMatrix(0, 6) = "Description"
flxImport.TextMatrix(0, 7) = "Expected"
flxImport.TextMatrix(0, 8) = "Group"
flxImport.TextMatrix(0, 9) = "Validation"
flxImport.Rows = 2
End Sub

Public Function IncorrectHeaderDetails() As Boolean
    If flxImport.TextMatrix(0, 0) <> "Step ID" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 1) <> "Component ID" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 2) <> "Test Set Folder" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 3) <> "Component Name" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 4) <> "Step Order" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 5) <> "Step Name" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 6) <> "Description" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 7) <> "Expected" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 8) <> "Group" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 9) <> "Validation" Then IncorrectHeaderDetails = True
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
    xlObject.Sheets(curTab).Range("A:J").Select

    xlObject.Sheets(curTab).Range("A:J").Borders(xlDiagonalDown).LineStyle = xlNone
    xlObject.Sheets(curTab).Range("A:J").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlObject.Sheets(curTab).Range("A:J").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:J").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:J").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:J").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:J").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:J").Borders(xlInsideHorizontal)
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
    xlObject.Sheets(curTab).Range("A:J").Select
    xlObject.Sheets(curTab).Range("A:J").EntireColumn.AutoFit
    xlObject.Sheets(curTab).Range("A1").Select

    xlObject.Sheets(curTab).Range("A1").AddComment
    xlObject.Sheets(curTab).Range("A1").Comment.Visible = False
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & mdiMain.Caption & "] " & Format(Now, "mmm-dd-yyyy HHMMSS AMPM") & ""
    
    xlObject.Sheets(curTab).Range("A:B").Interior.ColorIndex = 3
    xlObject.Sheets(curTab).Range("I:J").Interior.ColorIndex = 3

  xlObject.Sheets(curTab).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
  xlObject.Workbooks(1).SaveAs "BC_STEPS-01 " & CleanTheString(QCTree.SelectedItem.Text) & " - " & Format(Now, "mmm-dd-yyyy HHMM AMPM")
  xlObject.Visible = True
  xlObject.ActiveWindow.Activate

  Set xlWB = Nothing
  Set xlObject = Nothing

  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export To Excel Completed": Exit Sub:
OutErr:     MsgBox Err.Description, vbCritical: xlObject.Visible = True: xlObject.ActiveWindow.Activate: Set xlWB = Nothing: Set xlObject = Nothing
On Error GoTo 0
End Sub

Private Function GetBusinessComponentFolderPath(strID As String) As String
Dim Fact As ComponentFactory
Dim Obj As Component
If Trim(strID) = "" Then Exit Function
Set Fact = QCConnection.ComponentFactory
Set Obj = Fact.Item(strID)
GetBusinessComponentFolderPath = Obj.folder.path
End Function

Public Function CleanHTML_(strText As String) As String
        Dim tmp As String
        tmp = Replace(strText, "<br>", vbCrLf, 1, , vbTextCompare)
        tmp = Replace(tmp, "<html><body>", "", 1, , vbTextCompare)
        tmp = Replace(tmp, "</body></html>", "", 1, , vbTextCompare)
        tmp = Replace(tmp, "&", "&amp;", 1, , vbTextCompare)
        tmp = Replace(tmp, "'", "''", 1, , vbTextCompare)
        tmp = Replace(tmp, "<", "&lt;", 1, , vbTextCompare)
        tmp = Replace(tmp, ">", "&gt;", 1, , vbTextCompare)
        tmp = Replace(tmp, """", "&quot;", 1, , vbTextCompare)
        tmp = Replace(tmp, vbCrLf, "<br>", 1, , vbTextCompare)
        tmp = ReplaceAllEnter(CStr(tmp))
        CleanHTML_ = tmp
End Function

Private Sub txtFilter_DblClick()
txtFilter.Locked = False
End Sub
