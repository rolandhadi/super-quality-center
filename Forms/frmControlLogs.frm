VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmControlLogs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New Log Component"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   11580
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtImagePath 
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   7980
      Width           =   9555
   End
   Begin VB.CheckBox chkInclude 
      Caption         =   "Include FAIL"
      Height          =   315
      Left            =   1920
      TabIndex        =   17
      Top             =   8460
      Width           =   1515
   End
   Begin VB.TextBox txtDescription 
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   7560
      Width           =   9555
   End
   Begin VB.TextBox txtComponentName 
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   7140
      Width           =   9555
   End
   Begin VB.TextBox txtStepSummary 
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   6720
      Width           =   9555
   End
   Begin VB.TextBox txtResult 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   6300
      Width           =   9555
   End
   Begin VB.TextBox txtExeTime 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   5460
      Width           =   9555
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8460
      Width           =   1035
   End
   Begin VB.TextBox txtElapseTime 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   5880
      Width           =   9555
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add Log"
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8460
      Width           =   1755
   End
   Begin VB.TextBox txtFilter 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   60
      Width           =   9555
   End
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   4875
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   8599
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      WordWrap        =   -1  'True
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
   Begin VB.Label Label8 
      Caption         =   "IMAGE_PATH"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   8100
      Width           =   1635
   End
   Begin VB.Label Label7 
      Caption         =   "STEP_DESCRIPTION"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   7680
      Width           =   1635
   End
   Begin VB.Label Label6 
      Caption         =   "COMPONENT_NAME"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   7260
      Width           =   1635
   End
   Begin VB.Label Label5 
      Caption         =   "STEP_SUMMARY"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   6840
      Width           =   1635
   End
   Begin VB.Label Label4 
      Caption         =   "STEP_RESULT"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   6420
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "ELAPSED_TIME"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   6000
      Width           =   1635
   End
   Begin VB.Label Label3 
      Caption         =   "EXECUTION_TIME"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   5580
      Width           =   1635
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
Attribute VB_Name = "frmControlLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Log
     EXECUTION_TIME As String
     ELAPSED_TIME As String
     STEP_RESULT As String
     STEP_SUMMARY As String
     Component_Name As String
     STEP_DESCRIPTION  As String
End Type

Public EXECUTION_TIME As String
Public ELAPSED_TIME As String
Public STEP_RESULT As String
Public STEP_SUMMARY As String
Public Component_Name As String
Public STEP_DESCRIPTION  As String

Dim tmpLogs() As Log

Private Sub cmdAdd_Click()
Dim i
Dim CNV, c
Dim A, B
On Error GoTo Err1
If flxImport.RowSel > 0 Then
    If (Left(frmSinger.QCTree.SelectedItem.Key, 1) = "C" Or Left(frmSinger.QCTree.SelectedItem.Key, 1) = "R") And flxImport.TextMatrix(flxImport.RowSel, 1) <> "" Then
        frmSinger.AutoRedraw = False
        CNV = GetLastInx + 1
        c = GetLastInxC + 1
        A = frmSinger.QCTree.SelectedItem.Key
        B = CStr("C" & c)
        frmSinger.QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "0000") & "] " & "Log Entry", 4
        frmSinger.QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "EXECUTION_TIME", 2
        frmSinger.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtExeTime.Text, 3
        CNV = CNV + 1
        frmSinger.QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "ELAPSED_TIME", 2
        frmSinger.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtElapseTime.Text, 3
        CNV = CNV + 1
        frmSinger.QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STEP_RESULT", 2
        frmSinger.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtResult.Text, 3
        CNV = CNV + 1
        frmSinger.QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STEP_SUMMARY", 2
        frmSinger.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtStepSummary.Text, 3
        CNV = CNV + 1
        frmSinger.QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "COMPONENT_NAME", 2
        frmSinger.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtComponentName.Text, 3
        CNV = CNV + 1
        frmSinger.QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "STEP_DESCRIPTION", 2
        frmSinger.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtDescription.Text, 3
        CNV = CNV + 1
        If Trim(txtImagePath.Text) <> "" And InStr(1, txtImagePath.Text, "./images/", vbTextCompare) <> 0 Then
            frmSinger.QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), "IMAGE_PATH", 2
            frmSinger.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), "Captured Image Stored in: " & txtImagePath.Text, 3
            CNV = CNV + 1
        End If
        If Left(frmSinger.QCTree.SelectedItem.Key, 1) = "R" Then
          frmSinger.RenumberTree
        Else
          frmSinger.DragMove2 A, B
        End If
        frmSinger.flxImport.Rows = frmSinger.flxImport.Rows + 1
        frmSinger.PushToFlex
        frmSinger.QCTree.Nodes(CStr("C" & c)).Selected = True
        frmSinger.AutoRedraw = True
        frmSinger.flxImport.row = frmSinger.flxImport.row + 1
        frmSinger.flxImport.col = 0
        frmSinger.flxImport.RowSel = frmSinger.flxImport.row
        frmSinger.flxImport.ColSel = 7
    End If
End If
Exit Sub
Err1:
MsgBox Err.Description, vbCritical
End Sub

Private Function GetLastInx()
GetLastInx = CInt(frmSinger.QCTree.Nodes.Count)
'GetLastInx = CInt(Right(frmSinger.QCTree.Nodes(1).Child.LastSibling.Child.LastSibling.LastSibling.Key, Len(frmSinger.QCTree.Nodes(1).Child.LastSibling.Child.LastSibling.Key) - 1))
End Function

Private Function GetLastInxC()
GetLastInxC = CInt(frmSinger.QCTree.Nodes.Count)
'GetLastInxC = CInt(Right(frmSinger.QCTree.Nodes(1).Child.LastSibling.Key, Len(frmSinger.QCTree.Nodes(1).Child.LastSibling.Key) - 1))
End Function

Private Sub cmdSkip_Click()
Unload Me
End Sub

Private Sub flxImport_Click()
If flxImport.RowSel <> 0 Then
    txtStepSummary.Text = flxImport.TextMatrix(flxImport.RowSel, 0)
    txtComponentName.Text = flxImport.TextMatrix(flxImport.RowSel, 1)
    txtStepSummary.Text = flxImport.TextMatrix(flxImport.RowSel, 2)
    txtDescription.Text = flxImport.TextMatrix(flxImport.RowSel, 3)
End If
End Sub

Private Sub flxImport_dblClick()
cmdAdd_Click
End Sub

Private Sub Form_Load()
Dim FileFunct As New clsFiles
Dim WinFunct As New clsWindow
Dim tmp, tmp2, i, j, tmpStr
WinFunct.Ontop Me
tmp = Split(FileFunct.ReadFromFile(App.path & "\SQC DAT" & "\" & "LogsLegend.hxh"), vbCrLf)
j = 0
ReDim tmpLogs(j)
For i = LBound(tmp) To UBound(tmp) - 1
    tmp2 = Split(tmp(i), "|")
    ReDim Preserve tmpLogs(j)
    tmpLogs(j).STEP_RESULT = tmp2(0)
    tmpLogs(j).Component_Name = tmp2(3)
    tmpLogs(j).STEP_SUMMARY = tmp2(1)
    tmpLogs(j).STEP_DESCRIPTION = tmp2(2)
    j = j + 1
Next
With Me
    .txtExeTime.Text = EXECUTION_TIME
    .txtElapseTime.Text = ELAPSED_TIME
    .txtResult.Text = STEP_RESULT
    .txtStepSummary.Text = STEP_SUMMARY
    .txtComponentName.Text = Component_Name
    .txtDescription.Text = STEP_DESCRIPTION
    .flxImport.ColWidth(0) = 1500
    .flxImport.ColWidth(1) = 3000
    .flxImport.ColWidth(2) = 3000
    .flxImport.ColWidth(3) = 3000
End With
End Sub

Public Sub txtFilter_KeyPress(KeyAscii As Integer)
Dim i
If KeyAscii = 13 And Len(txtFilter) >= 2 Then
    flxImport.Clear
    flxImport.TextMatrix(0, 0) = "STEP_RESULT"
    flxImport.TextMatrix(0, 1) = "COMPONENT_NAME"
    flxImport.TextMatrix(0, 2) = "STEP_SUMMARY"
    flxImport.TextMatrix(0, 3) = "STEP_DESCRIPTION"
    flxImport.Rows = 2
    For i = LBound(tmpLogs) To UBound(tmpLogs)
        If InStr(1, tmpLogs(i).Component_Name, txtFilter.Text, vbTextCompare) <> 0 Or InStr(1, tmpLogs(i).STEP_DESCRIPTION, txtFilter.Text, vbTextCompare) <> 0 Or InStr(1, tmpLogs(i).STEP_SUMMARY, txtFilter.Text, vbTextCompare) <> 0 Then
            If chkInclude.Value = Checked Then
                flxImport.TextMatrix(flxImport.Rows - 1, 0) = tmpLogs(i).STEP_RESULT
                flxImport.TextMatrix(flxImport.Rows - 1, 1) = tmpLogs(i).Component_Name
                flxImport.TextMatrix(flxImport.Rows - 1, 2) = tmpLogs(i).STEP_SUMMARY
                flxImport.TextMatrix(flxImport.Rows - 1, 3) = tmpLogs(i).STEP_DESCRIPTION
                flxImport.Rows = flxImport.Rows + 1
            Else
                If (Trim(UCase(tmpLogs(i).STEP_RESULT)) <> "FAIL" And Trim(UCase(tmpLogs(i).STEP_RESULT)) <> "FAILED") Then
                    flxImport.TextMatrix(flxImport.Rows - 1, 0) = tmpLogs(i).STEP_RESULT
                    flxImport.TextMatrix(flxImport.Rows - 1, 1) = tmpLogs(i).Component_Name
                    flxImport.TextMatrix(flxImport.Rows - 1, 2) = tmpLogs(i).STEP_SUMMARY
                    flxImport.TextMatrix(flxImport.Rows - 1, 3) = tmpLogs(i).STEP_DESCRIPTION
                    flxImport.Rows = flxImport.Rows + 1
                End If
            End If
        End If
    Next
    flxImport.Rows = flxImport.Rows - 1
    cmdSort_Click
    For i = 1 To flxImport.Rows - 1
        If Trim(UCase(flxImport.TextMatrix(i, 1))) = Trim(UCase(txtFilter.Text)) Then
            flxImport.col = 0
            flxImport.row = i
            flxImport.ColSel = 3
            Exit For
        End If
    Next
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
  
  cColumn.Add 2
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
