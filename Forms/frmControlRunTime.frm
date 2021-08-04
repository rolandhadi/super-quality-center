VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmControlRunTime 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New Component"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   13275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRun 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Run"
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
      Left            =   11580
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton cmdHighlight 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Highlight"
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
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   60
      Width           =   1035
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   14
      Left            =   3240
      TabIndex        =   33
      Top             =   8280
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   13
      Left            =   3240
      TabIndex        =   31
      Top             =   7920
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   12
      Left            =   3240
      TabIndex        =   29
      Top             =   7560
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   11
      Left            =   3240
      TabIndex        =   27
      Top             =   7200
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   10
      Left            =   3240
      TabIndex        =   25
      Top             =   6840
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   9
      Left            =   3240
      TabIndex        =   23
      Top             =   6480
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   8
      Left            =   3240
      TabIndex        =   21
      Top             =   6120
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   7
      Left            =   3240
      TabIndex        =   19
      Top             =   5760
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   6
      Left            =   3240
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   5
      Left            =   3240
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   4
      Left            =   3240
      TabIndex        =   13
      Top             =   4680
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   2
      Left            =   3240
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.CommandButton cmdSkip 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add Component"
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
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   1755
   End
   Begin VB.TextBox txtFilter 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   60
      Width           =   6675
   End
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   2715
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   4789
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      WordWrap        =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   14
      Left            =   120
      TabIndex        =   34
      Top             =   8340
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   13
      Left            =   120
      TabIndex        =   32
      Top             =   7980
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   30
      Top             =   7620
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   28
      Top             =   7260
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   26
      Top             =   6900
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   24
      Top             =   6540
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   6180
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   20
      Top             =   5820
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   5460
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   5100
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   4740
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   4380
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   4020
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   3660
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblFieldName 
      Caption         =   "Field Properties"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   3300
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Component Name"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   1935
   End
End
Attribute VB_Name = "frmControlRunTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tmpProperties_ As String

Private Function GetLastInx()
GetLastInx = CInt(frmConsolidate.QCTree.Nodes.Count)
'GetLastInx = CInt(Right(frmConsolidate.QCTree.Nodes(1).Child.LastSibling.Child.LastSibling.LastSibling.Key, Len(frmConsolidate.QCTree.Nodes(1).Child.LastSibling.Child.LastSibling.Key) - 1))
End Function

Private Function GetLastInxC()
GetLastInxC = CInt(frmConsolidate.QCTree.Nodes.Count)
'GetLastInxC = CInt(Right(frmConsolidate.QCTree.Nodes(1).Child.LastSibling.Key, Len(frmConsolidate.QCTree.Nodes(1).Child.LastSibling.Key) - 1))
End Function

Private Sub cmdAdd_Click()
Dim i, tmp
Dim CNV, c
Dim A, B
On Error GoTo Err1
If flxImport.RowSel > 0 Then
    If (Left(frmConsolidate.QCTree.SelectedItem.Key, 1) = "C" Or Left(frmConsolidate.QCTree.SelectedItem.Key, 1) = "R") And flxImport.TextMatrix(flxImport.RowSel, 1) <> "" Then
        frmConsolidate.AutoRedraw = False
        CNV = GetLastInx + 1
        c = GetLastInxC + 1
        A = frmConsolidate.QCTree.SelectedItem.Key
        B = CStr("C" & c)
        If REALTIME = True Then
            frmConsolidate.QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "000") & "] " & GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Name, 4
            For i = lblFieldName.LBound To lblFieldName.UBound
                If lblFieldName(i).Visible = True Then
                    frmConsolidate.QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), lblFieldName(i).Caption, 2
                    If InStr(1, txtProperties(i).Text, "|") = 0 Then
                        frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtProperties(i).Text, 3
                    Else
                        tmp = Split(txtProperties(i).Text, "|")
                        frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), "DT_" & UCase(Trim(ResolveParameter(Trim(tmp(0))))), 3
                        frmConsolidate.QCTree.Nodes(CStr("V" & CNV)).Tag = tmp(1)
                    End If
                Else
                    Exit For
                End If
                CNV = CNV + 1
            Next
        End If
        If Left(frmConsolidate.QCTree.SelectedItem.Key, 1) = "R" Then
          frmConsolidate.RenumberTree
        Else
          frmConsolidate.DragMove2 A, B
        End If
        frmConsolidate.QCTree.Nodes(CStr("C" & c)).Selected = True
        frmConsolidate.QCTree.Nodes(CStr("C" & c)).Tag = flxImport.TextMatrix(flxImport.RowSel, 2)
        frmConsolidate.AutoRedraw = True
    Else
    End If
    If Control_Auto = True Then Unload Me
End If
Exit Sub
Err1:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdHighlight_Click()
Dim qtp 'As New QuickTest.Application
Dim fileFunct As New clsFiles, tmp
On Error GoTo Err1
If flxImport.RowSel <= 0 Then Exit Sub
Set qtp = CreateObject("QuickTest.Application")
fileFunct.FileCopy App.path & "\SQC DAT\BC Template\Action1\Highlight.mts", App.path & "\SQC DAT\BC Template\Action1\Script.mts"
tmp = fileFunct.ReadFromFile(App.path & "\SQC DAT\BC Template\Action1\Script.mts")
tmp = Replace(tmp, "[FrameProperties]", GetProperty("FrameProperties"))
tmp = Replace(tmp, "[FieldProperties]", GetProperty("FieldProperties"))
fileFunct.FileWrite App.path & "\SQC DAT\BC Template\Action1\Script.mts", CStr(tmp)
qtp.Visible = True
qtp.New False
qtp.Open App.path & "\SQC DAT\BC Template"
qtp.Test.Run
If qtp.Test.LastRunResults.Status <> "Failed" Then
    MsgBox "Object found", vbInformation
Else
    MsgBox "Object not found", vbCritical
End If
qtp.New False
fileFunct.FileCopy App.path & "\SQC DAT\BC Template\Action1\Blank.mts", App.path & "\SQC DAT\BC Template\Action1\Script.mts"
Exit Sub
Err1:
MsgBox Err.Description
End Sub

Private Function GetProperty(PropertyName As String)
Dim i
    For i = lblFieldName.LBound To lblFieldName.UBound
        If UCase(Trim(lblFieldName(i).Caption)) = UCase(Trim(PropertyName)) Then
            GetProperty = txtProperties(i).Text
            Exit Function
        End If
        
    Next
End Function

Private Sub cmdRun_Click()
Dim qtp 'As New QuickTest.Application
Dim fileFunct As New clsFiles, tmp, i
On Error GoTo Err1
If flxImport.RowSel <= 0 Then Exit Sub
Set qtp = CreateObject("QuickTest.Application")
fileFunct.FileCopy App.path & "\SQC Logs\bin\" & curDomain & "-" & curProject & "\" & flxImport.TextMatrix(flxImport.RowSel, 2) & "\Action1\Script.mts", App.path & "\SQC DAT\BC Template\Action1\Script.mts"
tmp = fileFunct.ReadFromFile(App.path & "\SQC DAT\BC Template\Action1\Script.mts")

For i = lblFieldName.LBound To lblFieldName.UBound
    If lblFieldName(i).Caption <> "" Then
        tmp = Replace(tmp, "GetParameterValue(" & Chr(34) & lblFieldName(i) & Chr(34) & ")", Chr(34) & GetProperty(lblFieldName(i).Caption) & Chr(34))
        tmp = Replace(tmp, "ResolveParameter(" & Chr(34) & lblFieldName(i) & Chr(34) & ")", Chr(34) & GetProperty(lblFieldName(i).Caption) & Chr(34))
        tmp = Replace(tmp, "GetParameter(" & Chr(34) & lblFieldName(i) & Chr(34) & ")", Chr(34) & GetProperty(lblFieldName(i).Caption) & Chr(34))
        tmp = Replace(tmp, "Parameter(" & Chr(34) & lblFieldName(i) & Chr(34) & ")", Chr(34) & GetProperty(lblFieldName(i).Caption) & Chr(34))
    Else
        Exit For
    End If
Next

fileFunct.FileWrite App.path & "\SQC DAT\BC Template\Action1\Script.mts", CStr(tmp)
qtp.Visible = True
qtp.New False
qtp.Open App.path & "\SQC DAT\BC Template"
qtp.Test.Run
If qtp.Test.LastRunResults.Status <> "Failed" Then
    MsgBox "Object found", vbInformation
Else
    MsgBox "Object not found", vbCritical
End If
qtp.New False
fileFunct.FileCopy App.path & "\SQC DAT\BC Template\Action1\Blank.mts", App.path & "\SQC DAT\BC Template\Action1\Script.mts"
Exit Sub
Err1:
MsgBox Err.Description
End Sub

Private Sub cmdSkip_Click()
Unload Me
End Sub

Private Sub flxImport_Click()
If flxImport.Rows <= 2 Then EnterField
End Sub

Private Sub flxImport_EnterCell()
EnterField
End Sub

Private Sub EnterField()
Dim i
Dim CNV, c
Dim A, B
If REALTIME = True Then
    ClearFields
    If flxImport.row = 0 Then Exit Sub
    flxImport.TextMatrix(flxImport.row, 0) = GetBusinessComponentFolderPath(flxImport.TextMatrix(flxImport.row, 2))
    On Error GoTo Err1
    If flxImport.RowSel > 0 Then
        If (Left(frmConsolidate.QCTree.SelectedItem.Key, 1) = "C" Or Left(frmConsolidate.QCTree.SelectedItem.Key, 1) = "R") And flxImport.TextMatrix(flxImport.RowSel, 1) <> "" Then
            frmConsolidate.AutoRedraw = False
            If REALTIME = True Then
                LoadAllBusinessComponent flxImport.TextMatrix(flxImport.RowSel, 2), flxImport.TextMatrix(flxImport.RowSel, 1)
                For i = LBound(GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters) To UBound(GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters)
                    lblFieldName(i).Visible = True
                    txtProperties(i).Visible = True
                    lblFieldName(i).Caption = GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters(i).ParameterName
                    txtProperties(i).Text = GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters(i).ParameterValue
                    Me.height = lblFieldName(i).Top + lblFieldName(i).height + 650
                    Me.Top = (Screen.height - Me.height) / 2
                    Me.Left = (Screen.width - Me.width) / 2
                    txtFilter.Text = flxImport.TextMatrix(flxImport.row, 1)
                Next
            End If
        End If
End If
Exit Sub
Err1:
MsgBox Err.Description, vbCritical
End If
End Sub

Private Sub ClearFields()
Dim i
For i = lblFieldName.LBound To lblFieldName.UBound
    lblFieldName(i).Visible = False
    txtProperties(i).Visible = False
    lblFieldName(i).Caption = ""
    txtProperties(i).Text = ""
Next
On Error Resume Next
Me.height = lblFieldName(i).Top + lblFieldName(i).height + 650
Me.Top = (Screen.height - Me.height) / 2
Me.Left = (Screen.width - Me.width) / 2
End Sub

Private Sub Form_DblClick()
'Me.Left = mdiMain.width \ 2
'Me.Top = mdiMain.height \ 2
End Sub

Private Sub Form_Load()
Dim fileFunct As New clsFiles
Dim WinFunct As New clsWindow
Dim tmp, i, tmpStr
WinFunct.Ontop Me
End Sub


Public Sub txtFilter_KeyPress(KeyAscii As Integer)
Dim i
If KeyAscii = 13 And Len(txtFilter) >= 3 Then
    flxImport.Clear
    flxImport.TextMatrix(0, 0) = "Component Path"
    flxImport.TextMatrix(0, 1) = "Component Name"
    flxImport.TextMatrix(0, 2) = "ID"
    flxImport.ColWidth(0) = 4000
    flxImport.ColWidth(1) = 7500
    flxImport.ColWidth(2) = 100
    flxImport.Rows = 2
    If REALTIME = True Then
        For i = LBound(All_BC_QC) To UBound(All_BC_QC)
            If InStr(1, All_BC_QC(i).BC_Name, txtFilter.Text, vbTextCompare) <> 0 Then
                flxImport.TextMatrix(flxImport.Rows - 1, 1) = All_BC_QC(i).BC_Name
                flxImport.TextMatrix(flxImport.Rows - 1, 2) = All_BC_QC(i).BC_ID
                flxImport.Rows = flxImport.Rows + 1
            End If
        Next
        flxImport.Rows = flxImport.Rows - 1
        cmdSort_Click
        For i = 1 To flxImport.Rows - 1
            If Trim(UCase(flxImport.TextMatrix(i, 1))) = Trim(UCase(txtFilter.Text)) Then
                flxImport.col = 0
                flxImport.row = i
                flxImport.ColSel = 2
                Exit For
            End If
        Next
        If flxImport.Rows <> 1 Then
            If flxImport.TextMatrix(1, 1) <> "" Then flxImport.TextMatrix(1, 0) = GetBusinessComponentFolderPath(flxImport.TextMatrix(1, 2))
        End If
    Else
         For i = LBound(GetAllBusinessComponents_MyComps) To UBound(GetAllBusinessComponents_MyComps)
            If InStr(1, GetAllBusinessComponents_MyComps(i).BC_Name, txtFilter.Text, vbTextCompare) <> 0 Then
                flxImport.TextMatrix(flxImport.Rows - 1, 0) = GetAllBusinessComponents_MyComps(i).BC_Path
                flxImport.TextMatrix(flxImport.Rows - 1, 1) = GetAllBusinessComponents_MyComps(i).BC_Name
                flxImport.TextMatrix(flxImport.Rows - 1, 2) = i
                flxImport.Rows = flxImport.Rows + 1
            End If
         Next
         flxImport.Rows = flxImport.Rows - 1
    End If
End If
End Sub


Private Sub cmdSort_Click()
 Dim aData() As String
  Dim lRow As Long, LCol As Long
  Dim cColumn As Collection, cOrder As Collection
  
  flxImport.FixedCols = 0
  
  ' Put the grid data in a string array
  With flxImport
    If .Rows <= 1 Then Exit Sub
    ReDim aData(.Rows - .FixedRows - 1, .Cols - .FixedCols - 1)
    For lRow = .FixedRows To .Rows - 1
      For LCol = .FixedCols To .Cols - 1
        aData(lRow - .FixedRows, LCol - .FixedCols) = .TextMatrix(lRow, LCol)
      Next LCol
    Next lRow
  End With
  
  ' Set the sorting parameters
  Set cColumn = New Collection
  Set cOrder = New Collection
  
  cColumn.Add 1
  cOrder.Add 1  ' sort Ascending
  
  cColumn.Add 0
  cOrder.Add 1 ' sort Ascending

  cColumn.Add 3
  cOrder.Add 1  ' sort Ascending
  
  ' Sort the grid
  ShellSortMultiColumn aData, cColumn, cOrder
  
  ' Put the data back in the grid
  With flxImport
    For lRow = .FixedRows To .Rows - 1
      For LCol = .FixedCols To .Cols - 1
        .TextMatrix(lRow, LCol) = aData(lRow - .FixedRows, LCol - .FixedCols)
      Next LCol
    Next lRow
  End With

End Sub

Private Sub txtProperties_DblClick(Index As Integer)
frmWebObjectURI.myIndex = Index
frmWebObjectURI.Show
End Sub
