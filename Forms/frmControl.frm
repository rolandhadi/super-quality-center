VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmControl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New Component"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   11580
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFieldName 
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   5820
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
      TabIndex        =   13
      Top             =   6720
      Width           =   1035
   End
   Begin VB.TextBox txtOthers 
      Height          =   375
      Left            =   6540
      TabIndex        =   12
      Text            =   "Others"
      Top             =   6720
      Width           =   1995
   End
   Begin VB.OptionButton optOthers 
      Caption         =   "Others"
      Height          =   315
      Left            =   5100
      TabIndex        =   11
      Top             =   6780
      Width           =   1335
   End
   Begin VB.OptionButton optTag 
      Caption         =   "logical name"
      Height          =   315
      Index           =   2
      Left            =   3660
      TabIndex        =   10
      Top             =   6780
      Width           =   1335
   End
   Begin VB.OptionButton optTag 
      Caption         =   "text"
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   9
      Top             =   6780
      Width           =   1335
   End
   Begin VB.OptionButton optTag 
      Caption         =   "name"
      Height          =   315
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   6780
      Width           =   1335
   End
   Begin VB.TextBox txtValue 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   6240
      Width           =   9555
   End
   Begin VB.TextBox txtProperties 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   5400
      Width           =   9555
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
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
      Cols            =   3
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
   Begin VB.Label Label4 
      Caption         =   "Field Name"
      Height          =   195
      Left            =   780
      TabIndex        =   15
      Top             =   5940
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Field Value"
      Height          =   195
      Left            =   780
      TabIndex        =   7
      Top             =   6360
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Field Properties"
      Height          =   195
      Left            =   780
      TabIndex        =   5
      Top             =   5520
      Width           =   1155
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
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tmpProperties_ As String

Private Sub cmdAdd_Click()
Dim i
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
            LoadAllBusinessComponent flxImport.TextMatrix(flxImport.RowSel, 2), flxImport.TextMatrix(flxImport.RowSel, 1)
            frmConsolidate.QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "000") & "] " & GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Name, 4
            For i = LBound(GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters) To UBound(GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters)
                If UCase(Trim(GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters(i).ParameterName)) = "EMPTY" Then
                    'Do Nothing
                Else
                    frmConsolidate.QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters(i).ParameterName, 2
                    If Control_Auto = True Then
                        If InStr(1, GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters(i).ParameterName, "FieldProperties") <> 0 Then
                            frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtProperties.Text, 3
                        ElseIf InStr(1, GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters(i).ParameterName, "Uri") <> 0 Then
                            frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtProperties.Text, 3
                        ElseIf InStr(1, GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters(i).ParameterName, "FieldValue") <> 0 Then
                            If txtFieldName.Text <> "" Then
                                frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), "DT_" & UCase(Trim(ResolveParameter(Trim(txtFieldName.Text)))), 3
                                frmConsolidate.QCTree.Nodes(CStr("V" & CNV)).Tag = Trim(txtValue.Text)
                            Else
                                frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtValue.Text, 3
                            End If
                        ElseIf InStr(1, GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters(i).ParameterName, "TheValue") <> 0 Then
                            If txtFieldName.Text <> "" Then
                                frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), "DT_" & UCase(Trim(ResolveParameter(Trim(txtFieldName.Text)))), 3
                                frmConsolidate.QCTree.Nodes(CStr("V" & CNV)).Tag = Trim(txtValue.Text)
                            Else
                                frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtValue.Text, 3
                            End If
                        Else
                            frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters(i).ParameterValue, 3
                        End If
                    Else
                        frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), GetAllBusinessComponents_MyComps(UBound(GetAllBusinessComponents_MyComps)).BC_Parameters(i).ParameterValue, 3
                    End If
                    CNV = CNV + 1
                End If
            Next
        Else
            frmConsolidate.QCTree.Nodes.Add "Root", tvwChild, CStr("C" & c), "[" & Format(c, "000") & "] " & GetAllBusinessComponents_MyComps(flxImport.TextMatrix(flxImport.RowSel, 2)).BC_Name, 4
            For i = LBound(GetAllBusinessComponents_MyComps(flxImport.TextMatrix(flxImport.RowSel, 2)).BC_Parameters) To UBound(GetAllBusinessComponents_MyComps(flxImport.TextMatrix(flxImport.RowSel, 2)).BC_Parameters)
                If UCase(Trim(GetAllBusinessComponents_MyComps(flxImport.TextMatrix(flxImport.RowSel, 2)).BC_Parameters(i).ParameterName)) = "EMPTY" Then
                    'Do Nothing
                Else
                    frmConsolidate.QCTree.Nodes.Add CStr("C" & c), tvwChild, CStr("N" & CNV), GetAllBusinessComponents_MyComps(flxImport.TextMatrix(flxImport.RowSel, 2)).BC_Parameters(i).ParameterName, 2
                    If Control_Auto = True Then
                        If InStr(1, GetAllBusinessComponents_MyComps(flxImport.TextMatrix(flxImport.RowSel, 2)).BC_Parameters(i).ParameterName, "FieldProperties") <> 0 Then
                            frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtProperties.Text, 3
                        ElseIf InStr(1, GetAllBusinessComponents_MyComps(flxImport.TextMatrix(flxImport.RowSel, 2)).BC_Parameters(i).ParameterName, "Uri") <> 0 Then
                            frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtProperties.Text, 3
                        ElseIf InStr(1, GetAllBusinessComponents_MyComps(flxImport.TextMatrix(flxImport.RowSel, 2)).BC_Parameters(i).ParameterName, "FieldValue") <> 0 Then
                            If txtFieldName.Text <> "" Then
                                frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), "DT_" & UCase(Trim(ResolveParameter(Trim(txtFieldName.Text)))), 3
                                frmConsolidate.QCTree.Nodes(CStr("V" & CNV)).Tag = Trim(txtValue.Text)
                            Else
                                frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtValue.Text, 3
                            End If
                        ElseIf InStr(1, GetAllBusinessComponents_MyComps(flxImport.TextMatrix(flxImport.RowSel, 2)).BC_Parameters(i).ParameterName, "TheValue") <> 0 Then
                            If txtFieldName.Text <> "" Then
                                frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), "DT_" & UCase(Trim(ResolveParameter(Trim(txtFieldName.Text)))), 3
                                frmConsolidate.QCTree.Nodes(CStr("V" & CNV)).Tag = Trim(txtValue.Text)
                            Else
                                frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), txtValue.Text, 3
                            End If
                        Else
                            frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), GetAllBusinessComponents_MyComps(flxImport.TextMatrix(flxImport.RowSel, 2)).BC_Parameters(i).ParameterValue, 3
                        End If
                    Else
                        frmConsolidate.QCTree.Nodes.Add CStr("N" & CNV), tvwChild, CStr("V" & CNV), GetAllBusinessComponents_MyComps(flxImport.TextMatrix(flxImport.RowSel, 2)).BC_Parameters(i).ParameterValue, 3
                    End If
                    CNV = CNV + 1
                End If
            Next
        End If
        If Left(frmConsolidate.QCTree.SelectedItem.Key, 1) = "R" Then
          frmConsolidate.RenumberTree
        Else
          frmConsolidate.DragMove2 A, B
        End If
        frmConsolidate.QCTree.Nodes(CStr("C" & c)).Selected = True
        frmConsolidate.QCTree.Nodes(CStr("C" & c)).Tag = GetAllBusinessComponents_MyComps(flxImport.TextMatrix(flxImport.RowSel, 2)).BC_ID
        frmConsolidate.AutoRedraw = True
    Else
    End If
    If Control_Auto = True Then Unload Me
End If
Exit Sub
Err1:
MsgBox Err.Description, vbCritical
End Sub

Private Function GetLastInx()
GetLastInx = CInt(frmConsolidate.QCTree.Nodes.Count)
'GetLastInx = CInt(Right(frmConsolidate.QCTree.Nodes(1).Child.LastSibling.Child.LastSibling.LastSibling.Key, Len(frmConsolidate.QCTree.Nodes(1).Child.LastSibling.Child.LastSibling.Key) - 1))
End Function

Private Function GetLastInxC()
GetLastInxC = CInt(frmConsolidate.QCTree.Nodes.Count)
'GetLastInxC = CInt(Right(frmConsolidate.QCTree.Nodes(1).Child.LastSibling.Key, Len(frmConsolidate.QCTree.Nodes(1).Child.LastSibling.Key) - 1))
End Function

Private Sub cmdSkip_Click()
Unload Me
End Sub

Private Sub flxImport_DblClick()
cmdAdd_Click
End Sub

Private Sub flxImport_EnterCell()
If REALTIME = True Then
    flxImport.TextMatrix(flxImport.row, 0) = GetBusinessComponentFolderPath(flxImport.TextMatrix(flxImport.row, 2))
End If
End Sub

Private Sub Form_DblClick()
'Me.Left = mdiMain.width \ 2
'Me.Top = mdiMain.height \ 2
End Sub

Private Sub Form_Load()
Dim FileFunct As New clsFiles
Dim WinFunct As New clsWindow
Dim tmp, i, tmpStr
WinFunct.Ontop Me
If LastOthers = "" Then LastOthers = "OTHERS"
txtOthers.Text = LastOthers
End Sub

Private Sub optOthers_Click()
On Error Resume Next
txtOthers.SetFocus
txtOthers.SelStart = 1
txtOthers.SelLength = Len(txtOthers.Text)
LastOthers = Trim(txtOthers.Text)
txtProperties.Text = Replace(tmpProperties_, "logical name/text/name", LastOthers)
End Sub

Private Sub optTag_Click(Index As Integer)
txtProperties.Text = Replace(tmpProperties_, "logical name/text/name", optTag(Index).Caption)
frmControl_Option = Index
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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Dim i
'If KeyAscii = 13 And Len(txtFilter) >= 2 Then
'    flxImport.Clear
'    flxImport.TextMatrix(0, 0) = "Component Path"
'    flxImport.TextMatrix(0, 1) = "Component Name"
'    flxImport.TextMatrix(0, 2) = "ID"
'    flxImport.ColWidth(0) = 4000
'    flxImport.ColWidth(1) = 7500
'    flxImport.ColWidth(2) = 100
'    flxImport.Rows = 2
'    For i = LBound(GetAllBusinessComponents_MyComps) To UBound(GetAllBusinessComponents_MyComps)
'        If InStr(1, GetAllBusinessComponents_MyComps(i).BC_Name, txtFilter.Text, vbTextCompare) <> 0 Then
'            flxImport.TextMatrix(flxImport.Rows - 1, 0) = GetAllBusinessComponents_MyComps(i).BC_Path
'            flxImport.TextMatrix(flxImport.Rows - 1, 1) = GetAllBusinessComponents_MyComps(i).BC_Name
'            flxImport.TextMatrix(flxImport.Rows - 1, 2) = i
'            flxImport.Rows = flxImport.Rows + 1
'        End If
'    Next
'    flxImport.Rows = flxImport.Rows - 1
'    cmdSort_Click
'    For i = 1 To flxImport.Rows - 1
'        If Trim(UCase(flxImport.TextMatrix(i, 1))) = Trim(UCase(txtFilter.Text)) Then
'            flxImport.col = 0
'            flxImport.row = i
'            flxImport.ColSel = 2
'            Exit For
'        End If
'    Next
'End If
End Sub

Private Sub txtOthers_Change()
LastOthers = Trim(txtOthers.Text)
txtProperties.Text = Replace(tmpProperties_, "logical name/text/name", LastOthers)
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
