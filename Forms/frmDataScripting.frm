VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDataScripting 
   Caption         =   "Data Scripting Module"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12675
   Icon            =   "frmDataScripting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   12675
   Tag             =   "Data Scripting Module"
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
      Picture         =   "frmDataScripting.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Import step description and expected results from an excel file"
      Top             =   1620
      Width           =   2205
   End
   Begin VB.ListBox lstUpdateFields 
      Columns         =   4
      Height          =   735
      Left            =   4560
      Style           =   1  'Checkbox
      TabIndex        =   2
      ToolTipText     =   "Select Values to update"
      Top             =   780
      Width           =   7995
   End
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
         NumButtons      =   5
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
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdDeletePassedPars"
            Object.ToolTipText     =   "Delete Passed Parameters"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin VB.CheckBox chkCSV 
         Caption         =   "Download to CSV"
         Height          =   315
         Left            =   4455
         TabIndex        =   11
         Top             =   135
         Width           =   2655
      End
      Begin VB.CheckBox chkSelective 
         Caption         =   "Selective Upload"
         Height          =   315
         Left            =   2475
         TabIndex        =   9
         Top             =   135
         Width           =   2655
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
            Picture         =   "frmDataScripting.frx":1070
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataScripting.frx":1302
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataScripting.frx":1594
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
            Picture         =   "frmDataScripting.frx":1822
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataScripting.frx":1F34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataScripting.frx":2646
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataScripting.frx":2D58
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
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   11580
      TabIndex        =   5
      Top             =   1140
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   3915
      Left            =   4500
      TabIndex        =   6
      Top             =   2100
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   6906
      _Version        =   393216
      Cols            =   12
      WordWrap        =   -1  'True
      AllowUserResizing=   3
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   6060
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
            Picture         =   "frmDataScripting.frx":346A
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   21132
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
            Picture         =   "frmDataScripting.frx":39BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataScripting.frx":3C9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataScripting.frx":41EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataScripting.frx":473F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid tmpImport 
      Height          =   1155
      Left            =   10980
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   2037
      _Version        =   393216
      Cols            =   12
      WordWrap        =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      Caption         =   "Fields to Update"
      Height          =   195
      Left            =   4560
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmDataScripting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AllRunTimeParams() As RunTimeParams
Dim Last_TC_TEST_INSTANCE_ID
 
Private Const Header_01 = "<DATAPACKET><CONFIGURATION><SELECTION first_sel_row=""-1"" last_sel_row=""-1""/></CONFIGURATION><METADATA><COLUMNS>"
Const Footer_01 = "</COLUMNS></METADATA>"
Const Header_02 = "<ROWADATA>"
Const Footer_02 = "</ROWADATA>"
Const Footer_03 = "</DATAPACKET>"
Const ColumnHeader = "<COLUMN column_name=""XXX_CNAME_XXX"" column_value_type=""String""/>"
 
Dim UpdateList() As String
Dim UpdateList_Iter() As String
Dim UpdateFParamList() As String

Private Sub Load_Data_To_RunTime_Parameters()
stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Consolidating Records..."
Step2
stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Updating Database..."
Step3
End Sub

Private Sub Load_Data_To_RunTime_Parameters_SELECTIVE()
stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Consolidating Records... Please wait"
Step2_SELECTIVE
stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Updating Database..."
End Sub

Private Function Step2_SELECTIVE()
Dim objCommand
Dim rs As TDAPIOLELib.Recordset
Dim lastrow, i, DataOfString_OLD, DataOfString_NEW, DataOfString_UPLOAD, ColChange, start_, end_, strUploadData
lastrow = flxImport.Rows
flxImport.Rows = flxImport.Rows + 1
For i = 1 To lastrow - 1
    If flxImport.TextMatrix(i, 0) = "" Then Exit Function
    If UCase(Trim(flxImport.TextMatrix(i, 11))) = "YES" Then
        Set objCommand = QCConnection.Command
        objCommand.CommandText = "SELECT TC_DATA_OBJ FROM TESTCYCL WHERE TC_TESTCYCL_ID = " & Trim(flxImport.TextMatrix(i, 0))
        Set rs = objCommand.Execute
        If rs.RecordCount > 0 Then
                DataOfString_OLD = rs.FieldValue("TC_DATA_OBJ")
                If Trim(DataOfString_OLD) = "" Then
                    GenerateOutput_NoIteration CStr(Trim(flxImport.TextMatrix(i, 0)))
                    DataOfString_OLD = GetFromTable(CStr(Trim(flxImport.TextMatrix(i, 0))), "TC_TESTCYCL_ID", "TC_DATA_OBJ", "TESTCYCL")
                    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Selective Parameter Set (DATA POPULATED) " & Now & " " & flxImport.TextMatrix(i, 0)
                End If
                If Trim(DataOfString_OLD) <> "" Then
                    start_ = InStr(1, DataOfString_OLD, "col" & flxImport.TextMatrix(i, 3) & "=" & """", vbTextCompare)
                    If start_ <> 0 Then
                        end_ = InStr(start_ + 1 + Abs((start_ - InStr(start_, DataOfString_OLD, """", vbTextCompare))), DataOfString_OLD, """", vbTextCompare)
                        ColChange = Mid(DataOfString_OLD, start_, end_ - start_) & """"
                        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Selective Parameter Set (ORIG DATA) " & Now & " " & flxImport.TextMatrix(i, 0) & " - " & ColChange
                        DataOfString_NEW = flxImport.TextMatrix(i, 6)
                        DataOfString_NEW = Replace(DataOfString_NEW, "&", "&amp;")
                        DataOfString_NEW = Replace(DataOfString_NEW, "'", "''")
                        DataOfString_NEW = Replace(DataOfString_NEW, "<", "&lt;")
                        DataOfString_NEW = Replace(DataOfString_NEW, ">", "&gt;")
                        DataOfString_NEW = Replace(DataOfString_NEW, """", "&quot;")
                        DataOfString_NEW = "col" & flxImport.TextMatrix(i, 3) & "=" & """" & DataOfString_NEW & """"
                        If InStr(1, DataOfString_OLD, ColChange, vbTextCompare) <> 0 And InStr(1, DataOfString_OLD, Trim(flxImport.TextMatrix(i, 4)), vbTextCompare) <> 0 Then
                            strUploadData = ((Replace(DataOfString_OLD, ColChange, DataOfString_NEW, , , vbTextCompare)))
                            UpdateSelectiveParameters CStr(Trim(flxImport.TextMatrix(i, 0))), CStr(Trim(strUploadData)), CInt(i)
                        Else
                            FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Selective Parameter Set (NO MATCH) " & Now & " " & flxImport.TextMatrix(i, 0)
                            flxImport.TextMatrix(i, 11) = "Not Uploaded (No Match Found)"
                        End If
                    Else
                        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Selective Parameter Set (NO PARAMETER COLUMN FOUND) " & Now & " " & flxImport.TextMatrix(i, 0)
                        flxImport.TextMatrix(i, 11) = "Not Uploaded (NO PARAMETER COLUMN FOUND)"
                    End If
                Else
                    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Selective Parameter Set (NO DATA FOUND) " & Now & " " & flxImport.TextMatrix(i, 0)
                    flxImport.TextMatrix(i, 11) = "Not Uploaded (Data is Blank)"
                End If
        End If
        Debug.Print i
    End If
Next
flxImport.Rows = flxImport.Rows - 1
End Function

Private Function UpdateSelectiveParameters(ID As String, TC_OBJ_DATA As String, RowNum As Integer)
Dim objCommand1
Dim rs1 As TDAPIOLELib.Recordset
Dim objCommand2
Dim rs2 As TDAPIOLELib.Recordset
On Error GoTo errHandler
If Trim(TC_OBJ_DATA) = "" Or Trim(ID) = "" Then FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Selective Parameter Set (NO DATA) " & Now & " " & ID: Exit Function
Set objCommand1 = QCConnection.Command
Set objCommand2 = QCConnection.Command
objCommand1.CommandText = "SELECT TC_TESTCYCL_ID FROM TESTCYCL WHERE TC_TESTCYCL_ID = " & ID
Set rs1 = objCommand1.Execute
If rs1.RecordCount = 1 Then
    objCommand2.CommandText = "UPDATE TESTCYCL SET tc_data_obj = '" & TC_OBJ_DATA & "' WHERE tc_testcycl_id = " & ID
    Set rs2 = objCommand2.Execute
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Selective Parameter Set (PASSED) " & Now & " " & ID
    flxImport.TextMatrix(RowNum, 11) = "Uploaded"
Else
    FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Selective Parameter Set (MULTIPLE OR NO MATCH FOUND) " & Now & " " & ID
    flxImport.TextMatrix(RowNum, 11) = "Not Uploaded (Multiple Entries or No Match Found)"
End If
Exit Function
errHandler:
FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Selective Parameter Set (FAILED) " & Now & " " & ID & " " & Err.Description & " - " & objCommand2.CommandText
flxImport.TextMatrix(RowNum, 11) = "Not Uploaded (FAILED)"
End Function
 
Private Function Step2()
    Dim GlobalLoop
    Dim i As Long
    Dim tmpColumnHeader
    Dim tmpRowData()
    Dim LastItemRow
    Dim AlltmpData
    Dim IT_Start As Boolean
    Dim DataOfString
    Dim lastrow
    Dim k
    Dim tmp
    
    lastrow = flxImport.Rows
    flxImport.Rows = flxImport.Rows + 1
    ReDim UpdateList(0)
    ReDim UpdateList_Iter(0)
    ReDim UpdateFParamList(0)
    ReDim tmpRowData(0)
 
    For GlobalLoop = 1 To lastrow
    tmpColumnHeader = ""
    tmpRowData(UBound(tmpRowData)) = ""
    IT_Start = True
        For i = GlobalLoop To lastrow - 1
            DataOfString = flxImport.TextMatrix(i, 6)
            DataOfString = Replace(DataOfString, "&", "&amp;")
            DataOfString = Replace(DataOfString, "'", "''")
            DataOfString = Replace(DataOfString, "<", "&lt;")
            DataOfString = Replace(DataOfString, ">", "&gt;")
            DataOfString = Replace(DataOfString, """", "&quot;")
            If flxImport.TextMatrix(i, 0) = flxImport.TextMatrix(i + 1, 0) And flxImport.TextMatrix(i, 2) = flxImport.TextMatrix(i + 1, 2) Then
                If IT_Start = True Then tmpColumnHeader = tmpColumnHeader & Replace(ColumnHeader, "XXX_CNAME_XXX", flxImport.TextMatrix(i, 4), , , vbTextCompare)
                tmpRowData(UBound(tmpRowData)) = tmpRowData(UBound(tmpRowData)) & "col" & flxImport.TextMatrix(i, 3) & "=""" & DataOfString & """ "
            ElseIf flxImport.TextMatrix(i, 0) = flxImport.TextMatrix(i + 1, 0) And flxImport.TextMatrix(i, 2) <> flxImport.TextMatrix(i + 1, 2) Then
                If IT_Start = True Then tmpColumnHeader = tmpColumnHeader & Replace(ColumnHeader, "XXX_CNAME_XXX", flxImport.TextMatrix(i, 4), , , vbTextCompare)
                tmpRowData(UBound(tmpRowData)) = "<ROW " & tmpRowData(UBound(tmpRowData)) & "col" & flxImport.TextMatrix(i, 3) & "=""" & DataOfString & """ " & "/>"
                ReDim Preserve tmpRowData(UBound(tmpRowData) + 1)
                IT_Start = False
            Else
                tmpColumnHeader = tmpColumnHeader & Replace(ColumnHeader, "XXX_CNAME_XXX", flxImport.TextMatrix(i, 4), , , vbTextCompare)
                tmpRowData(UBound(tmpRowData)) = "<ROW " & tmpRowData(UBound(tmpRowData)) & "col" & flxImport.TextMatrix(i, 3) & "=""" & DataOfString & """ " & "/>"
                For k = 0 To UBound(tmpRowData)
                    AlltmpData = AlltmpData & tmpRowData(k) & " "
                Next
                ReDim tmpRowData(0)
                UpdateList(UBound(UpdateList)) = "UPDATE TESTCYCL SET tc_data_obj = '" & Header_01 & tmpColumnHeader & Footer_01 & Header_02 & AlltmpData & Footer_02 & Footer_03 & "' WHERE tc_testcycl_id = " & flxImport.TextMatrix(i, 0)
                ReDim Preserve UpdateList(UBound(UpdateList) + 1)
                
                tmp = Split(Header_01 & tmpColumnHeader & Footer_01 & Header_02 & AlltmpData & Footer_02 & Footer_03, "<ROW Col1", , vbTextCompare)
                If UBound(tmp) > 0 Then
                    UpdateList_Iter(UBound(UpdateList_Iter)) = "UPDATE TESTCYCL SET TC_ITERATIONS = '" & UBound(tmp) & ";-1;-1" & "' WHERE TC_TESTCYCL_ID = " & flxImport.TextMatrix(i, 0)
                Else
                    UpdateList_Iter(UBound(UpdateList_Iter)) = "UPDATE TESTCYCL SET TC_ITERATIONS = '' WHERE TC_TESTCYCL_ID = " & flxImport.TextMatrix(i, 0)
                End If
                ReDim Preserve UpdateList_Iter(UBound(UpdateList_Iter) + 1)
                
                UpdateFParamList(UBound(UpdateFParamList)) = flxImport.TextMatrix(i, 10)
                ReDim Preserve UpdateFParamList(UBound(UpdateFParamList) + 1)
                
                GlobalLoop = i
                IT_Start = True
                AlltmpData = ""
                Exit For
            End If
        Next
    Next
    flxImport.Rows = flxImport.Rows - 1
End Function
 
Private Function Step3()
Dim i
Dim objCommand
Dim rs
Dim numErr, tmp
    mdiMain.pBar.Max = UBound(UpdateList) + 3
    For i = 0 To UBound(UpdateList)
        On Error Resume Next
        If UpdateList(i) = "" Then mdiMain.pBar.Value = mdiMain.pBar.Max: FXGirl.EZPlay FXDataUploadCompleted: Exit Function
        Set objCommand = QCConnection.Command
'        Clipboard.Clear
'        Clipboard.SetText UpdateList(i)
        objCommand.CommandText = UpdateList(i)
        Set rs = objCommand.Execute
        objCommand.CommandText = UpdateList_Iter(i)
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[SQL] " & Now & " " & objCommand.CommandText
        Set rs = objCommand.Execute
        If InStr(1, tmp, UpdateFParamList(i), vbTextCompare) = 0 Then
            FixAllFrameworkParameter UpdateFParamList(i)
            tmp = tmp & UpdateFParamList(i) & " "
        End If
        Debug.Print i
        stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Updating Database... " & i + 1 & " of " & UBound(UpdateList) + 1
        If Err.Number <> 0 Then
          FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Parameter Set (FAILED) " & Now & " " & i + 1 & Err.Description
          numErr = numErr + 1
        Else
          FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Parameter Set (PASSED) " & Now & " " & i + 1
        End If
        Err.Clear
        On Error GoTo 0
        Set objCommand = Nothing
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
    If numErr <> 0 Then
      Dim tmpFile As New clsFiles
      frmLogs.Caption = App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log"
      frmLogs.txtLogs.Text = tmpFile.ReadFromFile_FAILED(App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log")
      frmLogs.Show 1
    End If
End Function

Public Function DecompileRunTimeParams(TC_TEST_INSTANCE_ID, TC_OBJ_DATA)
'On Error Resume Next
Dim tmp, i, j, k
Dim tmp2
Const ColumnDivider = "<COLUMN column_name="
Const IterationDivider = "<ROW "

If Trim(TC_OBJ_DATA) = "" Then
    DecompileRunTimeParams = "#N/A"
    Exit Function
End If

If Last_TC_TEST_INSTANCE_ID = TC_TEST_INSTANCE_ID Then
    DecompileRunTimeParams = "DONE"
    Exit Function
End If
    

Last_TC_TEST_INSTANCE_ID = TC_TEST_INSTANCE_ID

tmp = Split(TC_OBJ_DATA, IterationDivider)


ReDim AllRunTimeParams(0)
If UBound(tmp) > 0 Then
    ReDim AllRunTimeParams(UBound(tmp) - 1)
End If
'I stopped here >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

For j = LBound(AllRunTimeParams) To UBound(AllRunTimeParams)
    tmp = Split(TC_OBJ_DATA, ColumnDivider)
    For i = 1 To UBound(tmp)
        AllRunTimeParams(j).TestInstanceID = TC_TEST_INSTANCE_ID
        ReDim Preserve AllRunTimeParams(j).ParamName(i - 1)
        ReDim Preserve AllRunTimeParams(j).ParamValue(i - 1)
        AllRunTimeParams(j).ParamName(i - 1) = Trim(Replace(Replace(tmp(i), "column_value_type=""String""/>", ""), """", ""))
        If i = UBound(tmp) Then
            AllRunTimeParams(j).ParamName(i - 1) = Trim(Left(AllRunTimeParams(j).ParamName(i - 1), InStr(1, AllRunTimeParams(j).ParamName(i - 1), "</COLUMNS>") - 1))
        End If
    Next
Next

tmp = Split(TC_OBJ_DATA, IterationDivider)

For j = LBound(AllRunTimeParams) To UBound(AllRunTimeParams)
    For i = 1 To 1000
        tmp(j + 1) = Replace(tmp(j + 1), "col" & i & "=""", "<|")
    Next
    For i = 1 To 1000
        tmp(j + 1) = Replace(tmp(j + 1), "/>", "")
    Next
    For i = 1 To 1000
        tmp(j + 1) = Replace(tmp(j + 1), """", "|>")
    Next
    tmp2 = Split(tmp(j + 1), "<|")
    For k = 1 To UBound(tmp2)
        AllRunTimeParams(j).ParamValue(k - 1) = Trim(Replace(tmp2(k), "|>", ""))
    Next
Next
DecompileRunTimeParams = "DONE"
End Function

Private Function Step3_NoIteration()
Dim i
Dim objCommand
Dim rs
Dim numErr, tmp
    mdiMain.pBar.Max = UBound(UpdateList) + 3
    For i = 0 To UBound(UpdateList)
        On Error Resume Next
        If UpdateList(i) = "" Then Exit Function
        Set objCommand = QCConnection.Command
        objCommand.CommandText = UpdateList(i)
        Set rs = objCommand.Execute
        objCommand.CommandText = UpdateList_Iter(i)
        FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "[SQL] " & Now & " " & objCommand.CommandText
        Set rs = objCommand.Execute
        If InStr(1, tmp, UpdateFParamList(i), vbTextCompare) = 0 Then
            FixAllFrameworkParameter UpdateFParamList(i)
            tmp = tmp & UpdateFParamList(i) & " "
        End If
        Debug.Print i
        If Err.Number <> 0 Then
          FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Parameter Set (FAILED) " & Now & " " & i + 1 & Err.Description
          numErr = numErr + 1
        Else
          FileAppend App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log", "Uploading Parameter Set (PASSED) " & Now & " " & i + 1
        End If
        Err.Clear
        On Error GoTo 0
        Set objCommand = Nothing
    Next
    If numErr <> 0 Then
      Dim tmpFile As New clsFiles
      frmLogs.Caption = App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log"
      frmLogs.txtLogs.Text = tmpFile.ReadFromFile_FAILED(App.path & "\SQC Logs" & "\" & Format(Now, "mm-dd-yyyy") & ".log")
      frmLogs.Show 1
    End If
End Function

Private Function Step2_NoIteration()
    Dim GlobalLoop
    Dim i As Long
    Dim tmpColumnHeader
    Dim tmpRowData()
    Dim LastItemRow
    Dim AlltmpData
    Dim IT_Start As Boolean
    Dim DataOfString
    Dim lastrow
    Dim k
    Dim tmp
    
    lastrow = tmpImport.Rows
    tmpImport.Rows = tmpImport.Rows + 1
    ReDim UpdateList(0)
    ReDim UpdateList_Iter(0)
    ReDim UpdateFParamList(0)
    ReDim tmpRowData(0)
 
    For GlobalLoop = 1 To lastrow
    tmpColumnHeader = ""
    tmpRowData(UBound(tmpRowData)) = ""
    IT_Start = True
        For i = GlobalLoop To lastrow - 1
            DataOfString = tmpImport.TextMatrix(i, 6)
            DataOfString = Replace(DataOfString, "&", "&amp;")
            DataOfString = Replace(DataOfString, "'", "''")
            DataOfString = Replace(DataOfString, "<", "&lt;")
            DataOfString = Replace(DataOfString, ">", "&gt;")
            DataOfString = Replace(DataOfString, """", "&quot;")
            If tmpImport.TextMatrix(i, 0) = tmpImport.TextMatrix(i + 1, 0) And tmpImport.TextMatrix(i, 2) = tmpImport.TextMatrix(i + 1, 2) Then
                If IT_Start = True Then tmpColumnHeader = tmpColumnHeader & Replace(ColumnHeader, "XXX_CNAME_XXX", tmpImport.TextMatrix(i, 4), , , vbTextCompare)
                tmpRowData(UBound(tmpRowData)) = tmpRowData(UBound(tmpRowData)) & "col" & tmpImport.TextMatrix(i, 3) & "=""" & DataOfString & """ "
            ElseIf tmpImport.TextMatrix(i, 0) = tmpImport.TextMatrix(i + 1, 0) And tmpImport.TextMatrix(i, 2) <> tmpImport.TextMatrix(i + 1, 2) Then
                If IT_Start = True Then tmpColumnHeader = tmpColumnHeader & Replace(ColumnHeader, "XXX_CNAME_XXX", tmpImport.TextMatrix(i, 4), , , vbTextCompare)
                tmpRowData(UBound(tmpRowData)) = "<ROW " & tmpRowData(UBound(tmpRowData)) & "col" & tmpImport.TextMatrix(i, 3) & "=""" & DataOfString & """ " & "/>"
                ReDim Preserve tmpRowData(UBound(tmpRowData) + 1)
                IT_Start = False
            Else
                tmpColumnHeader = tmpColumnHeader & Replace(ColumnHeader, "XXX_CNAME_XXX", tmpImport.TextMatrix(i, 4), , , vbTextCompare)
                tmpRowData(UBound(tmpRowData)) = "<ROW " & tmpRowData(UBound(tmpRowData)) & "col" & tmpImport.TextMatrix(i, 3) & "=""" & DataOfString & """ " & "/>"
                For k = 0 To UBound(tmpRowData)
                    AlltmpData = AlltmpData & tmpRowData(k) & " "
                Next
                ReDim tmpRowData(0)
                UpdateList(UBound(UpdateList)) = "UPDATE TESTCYCL SET tc_data_obj = '" & Header_01 & tmpColumnHeader & Footer_01 & Header_02 & AlltmpData & Footer_02 & Footer_03 & "' WHERE tc_testcycl_id = " & tmpImport.TextMatrix(i, 0)
                ReDim Preserve UpdateList(UBound(UpdateList) + 1)
                
                tmp = Split(Header_01 & tmpColumnHeader & Footer_01 & Header_02 & AlltmpData & Footer_02 & Footer_03, "<ROW Col1", , vbTextCompare)
                If UBound(tmp) > 0 Then
                    UpdateList_Iter(UBound(UpdateList_Iter)) = "UPDATE TESTCYCL SET TC_ITERATIONS = '" & UBound(tmp) & ";-1;-1" & "' WHERE TC_TESTCYCL_ID = " & tmpImport.TextMatrix(i, 0)
                Else
                    UpdateList_Iter(UBound(UpdateList_Iter)) = "UPDATE TESTCYCL SET TC_ITERATIONS = '' WHERE TC_TESTCYCL_ID = " & tmpImport.TextMatrix(i, 0)
                End If
                ReDim Preserve UpdateList_Iter(UBound(UpdateList_Iter) + 1)
                
                UpdateFParamList(UBound(UpdateFParamList)) = tmpImport.TextMatrix(i, 10)
                ReDim Preserve UpdateFParamList(UBound(UpdateFParamList) + 1)
                
                GlobalLoop = i
                IT_Start = True
                AlltmpData = ""
                Exit For
            End If
        Next
    Next
    Step3_NoIteration
End Function

Private Sub GenerateOutput_NoIteration(TC_TESTCYCL_ID As String)
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
Dim tmpAllPars

tmpImport.Clear
tmpImport.Rows = 1
tmpImport.Cols = 12

    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT RTP_TEST_ID, TC_TEST_INSTANCE, CY_CYCLE_ID, TC_TESTCYCL_ID, RTP_ID, '1' AS ""Iteration"",  RTP_ORDER, RTP_NAME, RTP_BPTA_LONG_VALUE, '' AS ""RTP_ACTUAL_VALUE"", CF_ITEM_NAME, CY_CYCLE, TS_NAME, TC_DATA_OBJ FROM RUNTIME_PARAM, TEST, TESTCYCL, CYCLE, CYCL_FOLD WHERE TC_TEST_ID = TS_TEST_ID AND RTP_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND TC_TESTCYCL_ID = " & TC_TESTCYCL_ID & " ORDER BY CF_ITEM_ID, TC_CYCLE_ID, TC_TEST_ORDER, RTP_TEST_ID, RTP_ORDER "
    Debug.Print Me.Caption & "-" & objCommand.CommandText
    Set rs = objCommand.Execute
    tmpImport.Rows = rs.RecordCount + 1
    k = 0
    For i = 1 To rs.RecordCount
        Decompile = DecompileRunTimeParams(rs.FieldValue("TC_TESTCYCL_ID"), rs.FieldValue("TC_DATA_OBJ"))
        If Decompile = "DONE" Then
            For iter = LBound(AllRunTimeParams) To UBound(AllRunTimeParams)
                If k = 0 Then
                    tmpAllPars = "~!@#$%^&*()"
                    tmpAllPars = GetAllFrameworkParameter(rs.FieldValue("RTP_TEST_ID"))
                ElseIf tmpImport.TextMatrix(k - 1, 0) <> rs.FieldValue("TC_TESTCYCL_ID") Then
                    tmpAllPars = "~!@#$%^&*()"
                    tmpAllPars = GetAllFrameworkParameter(rs.FieldValue("RTP_TEST_ID"))
                End If
                If InStr(1, tmpAllPars, rs.FieldValue("RTP_NAME"), vbTextCompare) <> 0 Then
                    k = k + 1
                    tmpImport.Rows = k + 1
                    tmpImport.TextMatrix(k, 0) = rs.FieldValue("TC_TESTCYCL_ID")
                    tmpImport.TextMatrix(k, 1) = rs.FieldValue("RTP_ID")
                    tmpImport.TextMatrix(k, 2) = iter + 1
                    tmpImport.TextMatrix(k, 3) = rs.FieldValue("RTP_ORDER")
                    tmpImport.TextMatrix(k, 4) = rs.FieldValue("RTP_NAME")
                    tmpImport.TextMatrix(k, 5) = ReplaceAllEnter(rs.FieldValue("RTP_BPTA_LONG_VALUE"))
                    tmpImport.TextMatrix(k, 6) = ReplaceAllEnter(ReverseCleanHTML(GetParamValue(iter, rs.FieldValue("RTP_NAME"))))
                    tmpImport.TextMatrix(k, 8) = ReplaceAllEnter(rs.FieldValue("CY_CYCLE"))
                    If rs.FieldValue("CY_CYCLE_ID") <> tmpImport.TextMatrix(k - 1, 8) Then
                        tmpImport.TextMatrix(k, 7) = GetTestSetFolderPath(rs.FieldValue("CY_CYCLE_ID"))
                    Else
                        tmpImport.TextMatrix(k, 7) = tmpImport.TextMatrix(k - 1, 7)
                    End If
                    tmpImport.TextMatrix(k, 9) = "[" & rs.FieldValue("TC_TEST_INSTANCE") & "] " & ReplaceAllEnter(rs.FieldValue("TS_NAME"))
                    tmpImport.TextMatrix(k, 10) = ReplaceAllEnter(rs.FieldValue("RTP_TEST_ID"))
                End If
            Next
        ElseIf Decompile = "#N/A" Then
                If k = 0 Then
                    tmpAllPars = "~!@#$%^&*()"
                    tmpAllPars = GetAllFrameworkParameter(rs.FieldValue("RTP_TEST_ID"))
                ElseIf tmpImport.TextMatrix(k - 1, 0) <> rs.FieldValue("TC_TESTCYCL_ID") Then
                    tmpAllPars = "~!@#$%^&*()"
                    tmpAllPars = GetAllFrameworkParameter(rs.FieldValue("RTP_TEST_ID"))
                End If
                If InStr(1, tmpAllPars, rs.FieldValue("RTP_NAME"), vbTextCompare) <> 0 Then
                    k = k + 1
                    tmpImport.Rows = k + 1
                    tmpImport.TextMatrix(k, 0) = rs.FieldValue("TC_TESTCYCL_ID")
                    tmpImport.TextMatrix(k, 1) = rs.FieldValue("RTP_ID")
                    tmpImport.TextMatrix(k, 2) = 1
                    tmpImport.TextMatrix(k, 3) = rs.FieldValue("RTP_ORDER")
                    tmpImport.TextMatrix(k, 4) = rs.FieldValue("RTP_NAME")
                    tmpImport.TextMatrix(k, 5) = ReplaceAllEnter(rs.FieldValue("RTP_BPTA_LONG_VALUE"))
                    tmpImport.TextMatrix(k, 6) = ""
                    tmpImport.TextMatrix(k, 8) = ReplaceAllEnter(rs.FieldValue("CY_CYCLE"))
                    If rs.FieldValue("CY_CYCLE_ID") <> tmpImport.TextMatrix(k - 1, 8) Then
                        tmpImport.TextMatrix(k, 7) = GetTestSetFolderPath(rs.FieldValue("CY_CYCLE_ID"))
                    Else
                        tmpImport.TextMatrix(k, 7) = tmpImport.TextMatrix(k - 1, 7)
                    End If
                    tmpImport.TextMatrix(k, 9) = "[" & rs.FieldValue("TC_TEST_INSTANCE") & "] " & ReplaceAllEnter(rs.FieldValue("TS_NAME"))
                    tmpImport.TextMatrix(k, 10) = ReplaceAllEnter(rs.FieldValue("RTP_TEST_ID"))
                End If
        End If
        rs.Next
    Next
    Step2_NoIteration
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
Dim tmpAllPars, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, last_a7, last_a8
    
    TimeSt = Format(Now, "mmm-dd-yyyy hhmmss") & "-"
    If chkCSV.Value = Checked Then
      AllF = InputBox("Enter file name", "File name", "[Data Scripting] ")
    Else
      AllF = "[Data Scripting] "
    End If
    
    ReDim CheckedItems(0): strPath = ""
    GetAllCheckedItems QCTree.Nodes(1)
    For j = LBound(CheckedItems) To UBound(CheckedItems) - 1
        If Left(CheckedItems(j), 1) = "F" Then
            strPath = strPath & "CF_ITEM_PATH LIKE '" & GetFromTable(Right(CheckedItems(j), Len(CheckedItems(j)) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%'" & " OR "
        ElseIf Left(CheckedItems(j), 1) = "T" Then
            strPath = strPath & "TC_TESTCYCL_ID = " & Right(CheckedItems(j), Len(CheckedItems(j)) - 1) & " OR "
        Else
            strPath = strPath & "TC_CYCLE_ID = " & Right(CheckedItems(j), Len(CheckedItems(j)) - 1) & " OR "
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
      objCommand.CommandText = "SELECT RTP_TEST_ID, TC_TEST_INSTANCE, CY_CYCLE_ID, TC_TESTCYCL_ID, RTP_ID, '1' AS ""Iteration"",  RTP_ORDER, RTP_NAME, RTP_BPTA_LONG_VALUE, '' AS ""RTP_ACTUAL_VALUE"", CF_ITEM_NAME, CY_CYCLE, TS_NAME, TC_DATA_OBJ FROM RUNTIME_PARAM, TEST, TESTCYCL, CYCLE, CYCL_FOLD WHERE TC_TEST_ID = TS_TEST_ID AND RTP_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " AND " & Trim(txtFilter.Text) & " ORDER BY CF_ITEM_ID, TC_CYCLE_ID, TC_TEST_ORDER, RTP_TEST_ID, RTP_ORDER "
    Else
      objCommand.CommandText = "SELECT RTP_TEST_ID, TC_TEST_INSTANCE, CY_CYCLE_ID, TC_TESTCYCL_ID, RTP_ID, '1' AS ""Iteration"",  RTP_ORDER, RTP_NAME, RTP_BPTA_LONG_VALUE, '' AS ""RTP_ACTUAL_VALUE"", CF_ITEM_NAME, CY_CYCLE, TS_NAME, TC_DATA_OBJ FROM RUNTIME_PARAM, TEST, TESTCYCL, CYCLE, CYCL_FOLD WHERE TC_TEST_ID = TS_TEST_ID AND RTP_TEST_ID = TC_TEST_ID AND CY_CYCLE_ID = TC_CYCLE_ID AND CY_FOLDER_ID = CF_ITEM_ID AND " & strPath & " ORDER BY CF_ITEM_ID, TC_CYCLE_ID, TC_TEST_ORDER, RTP_TEST_ID, RTP_ORDER "
    End If
    Debug.Print Me.Caption & "-" & objCommand.CommandText
    Set rs = objCommand.Execute
    AllScript = """" & "Test Instance ID" & """" & "," & """" & "RunTime ID" & """" & "," & """" & "Iteration" & """" & "," & """" & "Parameter Order" & """" & "," & """" & "Parameter Name" & """" & "," & """" & "Default Value" & """" & "," & """" & "Actual Value" & """" & "," & """" & "Folder Name" & """" & "," & """" & "Test Set" & """" & "," & """" & "Test Instance" & """" & "," & """" & "Run Time Test ID" & """" & "," & """" & "Selective Upload" & """"
    ClearTable
    If chkCSV.Value = Unchecked Then '***
        flxImport.Rows = rs.RecordCount + 1
    End If '***
    k = 0
    mdiMain.pBar.Max = rs.RecordCount + 3
    For i = 1 To rs.RecordCount
        Decompile = DecompileRunTimeParams(rs.FieldValue("TC_TESTCYCL_ID"), rs.FieldValue("TC_DATA_OBJ"))
        If Decompile = "DONE" Then
            For iter = LBound(AllRunTimeParams) To UBound(AllRunTimeParams)
                If chkCSV.Value = Unchecked Then
                  If k = 0 Then
                      tmpAllPars = "~!@#$%^&*()"
                      tmpAllPars = GetAllFrameworkParameter(rs.FieldValue("RTP_TEST_ID"))
                  ElseIf flxImport.TextMatrix(k - 1, 0) <> rs.FieldValue("TC_TESTCYCL_ID") Then
                      tmpAllPars = "~!@#$%^&*()"
                      tmpAllPars = GetAllFrameworkParameter(rs.FieldValue("RTP_TEST_ID"))
                  End If
                  If InStr(1, tmpAllPars, rs.FieldValue("RTP_NAME"), vbTextCompare) <> 0 Then
                      stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Processing " & i & " out of " & rs.RecordCount & " - Iteration " & iter + 1 & " of " & UBound(AllRunTimeParams) + 1
                      k = k + 1
                      flxImport.Rows = k + 1
                      flxImport.TextMatrix(k, 0) = rs.FieldValue("TC_TESTCYCL_ID")
                      flxImport.TextMatrix(k, 1) = rs.FieldValue("RTP_ID")
                      flxImport.TextMatrix(k, 2) = iter + 1
                      flxImport.TextMatrix(k, 3) = rs.FieldValue("RTP_ORDER")
                      flxImport.TextMatrix(k, 4) = rs.FieldValue("RTP_NAME")
                      flxImport.TextMatrix(k, 5) = ReplaceAllEnter(rs.FieldValue("RTP_BPTA_LONG_VALUE"))
                      flxImport.TextMatrix(k, 6) = ReplaceAllEnter(ReverseCleanHTML(GetParamValue(iter, rs.FieldValue("RTP_NAME"))))
                      flxImport.TextMatrix(k, 8) = ReplaceAllEnter(rs.FieldValue("CY_CYCLE"))
                      If rs.FieldValue("CY_CYCLE_ID") <> flxImport.TextMatrix(k - 1, 8) Then
                          flxImport.TextMatrix(k, 7) = GetTestSetFolderPath(rs.FieldValue("CY_CYCLE_ID")) 'rs.FieldValue("CF_ITEM_NAME")
                      Else
                          flxImport.TextMatrix(k, 7) = flxImport.TextMatrix(k - 1, 7)
                      End If
                      flxImport.TextMatrix(k, 9) = "[" & rs.FieldValue("TC_TEST_INSTANCE") & "] " & ReplaceAllEnter(rs.FieldValue("TS_NAME"))
                      flxImport.TextMatrix(k, 10) = ReplaceAllEnter(rs.FieldValue("RTP_TEST_ID"))
                  End If
                Else
                  a0 = rs.FieldValue("TC_TESTCYCL_ID")
                  a1 = rs.FieldValue("RTP_ID")
                  a2 = iter + 1
                  a3 = rs.FieldValue("RTP_ORDER")
                  a4 = rs.FieldValue("RTP_NAME")
                  a5 = ReplaceAllEnter(rs.FieldValue("RTP_BPTA_LONG_VALUE"))
                  a6 = ReplaceAllEnter(ReverseCleanHTML(GetParamValue(iter, rs.FieldValue("RTP_NAME"))))
                  a8 = ReplaceAllEnter(rs.FieldValue("CY_CYCLE"))
                  If a8 <> last_a8 Then
                    a7 = GetTestSetFolderPath(rs.FieldValue("CY_CYCLE_ID"))
                    last_a7 = a7
                  Else
                    a7 = last_a7
                  End If
                  a9 = "[" & rs.FieldValue("TC_TEST_INSTANCE") & "] " & ReplaceAllEnter(rs.FieldValue("TS_NAME"))
                  a10 = ReplaceAllEnter(rs.FieldValue("RTP_TEST_ID"))
                  If Trim(AllScript) <> "" Then
                        AllScript = AllScript & vbCrLf & """" & a0 & """" & "," & """" & a1 & """" & "," & """" & a2 & """" & "," & """" & a3 & """" & "," & """" & a4 & """" & "," & """" & a5 & """" & "," & """" & a6 & """" & "," & """" & a7 & """" & "," & """" & a8 & """" & "," & """" & a9 & """" & "," & """" & a10 & """"
                  Else
                        AllScript = AllScript & """" & a0 & """" & "," & """" & a1 & """" & "," & """" & a2 & """" & "," & """" & a3 & """" & "," & """" & a4 & """" & "," & """" & a5 & """" & "," & """" & a6 & """" & "," & """" & a7 & """" & "," & """" & a8 & """" & "," & """" & a9 & """" & "," & """" & a10 & """"
                  End If
                End If
            Next
            mdiMain.pBar.Max = mdiMain.pBar.Max + UBound(AllRunTimeParams)
        ElseIf Decompile = "#N/A" Then
                If chkCSV.Value = Unchecked Then
                  If k = 0 Then
                      tmpAllPars = "~!@#$%^&*()"
                      tmpAllPars = GetAllFrameworkParameter(rs.FieldValue("RTP_TEST_ID"))
                  ElseIf flxImport.TextMatrix(k - 1, 0) <> rs.FieldValue("TC_TESTCYCL_ID") Then
                      tmpAllPars = "~!@#$%^&*()"
                      tmpAllPars = GetAllFrameworkParameter(rs.FieldValue("RTP_TEST_ID"))
                  End If
                  If InStr(1, tmpAllPars, rs.FieldValue("RTP_NAME"), vbTextCompare) <> 0 Then
                      stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Processing " & i & " out of " & rs.RecordCount & " - Iteration 1 of 1"
      '                AllScript = AllScript & vbCrLf & rs.FieldValue("TC_TESTCYCL_ID") & vbTab & rs.FieldValue("RTP_ID") & vbTab & 1 & vbTab & rs.FieldValue("RTP_ORDER") & vbTab & rs.FieldValue("RTP_NAME") & vbTab & rs.FieldValue("RTP_BPTA_LONG_VALUE") & vbTab & "" & vbTab & rs.FieldValue("CF_ITEM_NAME") & vbTab & rs.FieldValue("CY_CYCLE") & vbTab & rs.FieldValue("TS_NAME")
                      k = k + 1
                      flxImport.Rows = k + 1
                      flxImport.TextMatrix(k, 0) = rs.FieldValue("TC_TESTCYCL_ID")
                      flxImport.TextMatrix(k, 1) = rs.FieldValue("RTP_ID")
                      flxImport.TextMatrix(k, 2) = 1
                      flxImport.TextMatrix(k, 3) = rs.FieldValue("RTP_ORDER")
                      flxImport.TextMatrix(k, 4) = rs.FieldValue("RTP_NAME")
                      flxImport.TextMatrix(k, 5) = ReplaceAllEnter(rs.FieldValue("RTP_BPTA_LONG_VALUE"))
                      flxImport.TextMatrix(k, 6) = ""
                      flxImport.TextMatrix(k, 8) = ReplaceAllEnter(rs.FieldValue("CY_CYCLE"))
                      If rs.FieldValue("CY_CYCLE_ID") <> flxImport.TextMatrix(k - 1, 8) Then
                          flxImport.TextMatrix(k, 7) = GetTestSetFolderPath(rs.FieldValue("CY_CYCLE_ID")) 'rs.FieldValue("CF_ITEM_NAME")
                      Else
                          flxImport.TextMatrix(k, 7) = flxImport.TextMatrix(k - 1, 7)
                      End If
                      flxImport.TextMatrix(k, 9) = "[" & rs.FieldValue("TC_TEST_INSTANCE") & "] " & ReplaceAllEnter(rs.FieldValue("TS_NAME"))
                      flxImport.TextMatrix(k, 10) = ReplaceAllEnter(rs.FieldValue("RTP_TEST_ID"))
                  End If
                Else
                  a0 = rs.FieldValue("TC_TESTCYCL_ID")
                  a1 = rs.FieldValue("RTP_ID")
                  a2 = 1
                  a3 = rs.FieldValue("RTP_ORDER")
                  a4 = rs.FieldValue("RTP_NAME")
                  a5 = ReplaceAllEnter(rs.FieldValue("RTP_BPTA_LONG_VALUE"))
                  a6 = ""
                  a8 = ReplaceAllEnter(rs.FieldValue("CY_CYCLE"))
                  If a8 <> last_a8 Then
                    a7 = GetTestSetFolderPath(rs.FieldValue("CY_CYCLE_ID"))
                    last_a7 = a7
                  Else
                    a7 = last_a7
                  End If
                  a9 = "[" & rs.FieldValue("TC_TEST_INSTANCE") & "] " & ReplaceAllEnter(rs.FieldValue("TS_NAME"))
                  a10 = ReplaceAllEnter(rs.FieldValue("RTP_TEST_ID"))
                  If Trim(AllScript) <> "" Then
                        AllScript = AllScript & vbCrLf & """" & a0 & """" & "," & """" & a1 & """" & "," & """" & a2 & """" & "," & """" & a3 & """" & "," & """" & a4 & """" & "," & """" & a5 & """" & "," & """" & a6 & """" & "," & """" & a7 & """" & "," & """" & a8 & """" & "," & """" & a9 & """" & "," & """" & a10 & """"
                  Else
                        AllScript = AllScript & """" & a0 & """" & "," & """" & a1 & """" & "," & """" & a2 & """" & "," & """" & a3 & """" & "," & """" & a4 & """" & "," & """" & a5 & """" & "," & """" & a6 & """" & "," & """" & a7 & """" & "," & """" & a8 & """" & "," & """" & a9 & """" & "," & """" & a10 & """"
                  End If
                End If
        End If
        If chkCSV.Value = Checked Then '***
            If i Mod 10 = 0 Then
                FileAppend App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv", AllScript
                AllScript = ""
            End If
        End If '***
        rs.Next
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
    Next
    mdiMain.pBar.Value = mdiMain.pBar.Max
    FXGirl.EZPlay FXSQCExtractCompleted
If chkCSV.Value = Checked Then '***
    FileAppend App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv", AllScript: If MsgBox("Successfully exported to " & App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv" & vbCrLf & "Do you want to launch the extracted file?", vbYesNo) = vbYes Then Shell "explorer.exe " & App.path & "\SQC Logs" & "\", vbNormalFocus
    AllScript = vbCrLf & " ,"
    AllScript = AllScript & """" & "SQL Code:" & """" & "," & """" & Replace(objCommand.CommandText, """", "'") & """"
    FileAppend App.path & "\SQC Logs" & "\" & AllF & "_" & TimeSt & ".csv", AllScript
End If '***
stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Processed " & mdiMain.pBar.Max & " items"
End Sub

Private Function GetAllFrameworkParameter(Test_ID) As String
Dim tmpSQL, tmp, i, tmp2
Dim rs As TDAPIOLELib.Recordset, rs2 As TDAPIOLELib.Recordset
Dim objCommand

Set objCommand = QCConnection.Command
tmpSQL = "SELECT FP_NAME FROM   BP_ITER_PARAM, BP_PARAM, FRAMEWORK_PARAM, COMPONENT, TEST, BPTEST_TO_COMPONENTS, ALL_LISTS WHERE  BPIP_BPP_ID = BPP_ID AND    BPP_PARAM_ID = FP_ID AND    FP_COMPONENT_ID = CO_ID AND    TS_TEST_ID = BC_BPT_ID AND    BC_CO_ID = CO_ID AND    BC_ID = BPP_BPC_ID AND    TS_SUBJECT = AL_ITEM_ID  AND TS_TEST_ID = " & Test_ID
objCommand.CommandText = tmpSQL
Set rs = objCommand.Execute
For i = 1 To rs.RecordCount
    tmp = tmp & rs.FieldValue("FP_NAME") & " "
    rs.Next
Next
tmp = Trim(tmp)
If tmp = "" Then tmp = "~!@#$%^&*()"
GetAllFrameworkParameter = tmp
End Function

Private Sub FixAllFrameworkParameter(Test_ID)
Dim tmpSQL, tmp, i, tmp2
Dim rs As TDAPIOLELib.Recordset, rs2 As TDAPIOLELib.Recordset
Dim objCommand

Set objCommand = QCConnection.Command
tmpSQL = "SELECT FP_NAME FROM   BP_ITER_PARAM, BP_PARAM, FRAMEWORK_PARAM, COMPONENT, TEST, BPTEST_TO_COMPONENTS, ALL_LISTS WHERE  BPIP_BPP_ID = BPP_ID AND    BPP_PARAM_ID = FP_ID AND    FP_COMPONENT_ID = CO_ID AND    TS_TEST_ID = BC_BPT_ID AND    BC_CO_ID = CO_ID AND    BC_ID = BPP_BPC_ID AND    TS_SUBJECT = AL_ITEM_ID  AND TS_TEST_ID = " & Test_ID
objCommand.CommandText = tmpSQL
Set rs = objCommand.Execute
For i = 1 To rs.RecordCount
    tmp = tmp & rs.FieldValue("FP_NAME") & " "
    rs.Next
Next
tmp = Trim(tmp)
If tmp = "" Then tmp = "~!@#$%^&*()"

Set objCommand = QCConnection.Command
tmpSQL = "SELECT RTP_ID, RTP_NAME FROM RUNTIME_PARAM WHERE RTP_TEST_ID = " & Test_ID
objCommand.CommandText = tmpSQL
Set rs = objCommand.Execute
Do While rs.EOR = False
    If InStr(1, tmp, rs.FieldValue("RTP_NAME"), vbTextCompare) = 0 Then
        tmpSQL = "DELETE FROM RUNTIME_PARAM WHERE RTP_ID = " & rs.FieldValue("RTP_ID")
        objCommand.CommandText = tmpSQL
        Set rs2 = objCommand.Execute
    End If
    rs.Next
Loop
End Sub


 Function GetCommentText(rCommentCell As Range)
     Dim strGotIt As String
         On Error Resume Next
         strGotIt = WorksheetFunction.Clean _
             (rCommentCell.Comment.Text)
         GetCommentText = strGotIt
         On Error GoTo 0
End Function

Private Sub chkCSV_Click()
If chkCSV.Value = Checked Then
    If MsgBox("Are you sure you want to download directly to CSV?", vbYesNo) = vbYes Then
        chkCSV.Value = Checked
    Else
        chkCSV.Value = Unchecked
    End If
End If
End Sub

Private Sub chkSelective_Click()
If chkSelective.Value = Checked Then
    If MsgBox("Are you sure you want to do a selective upload?", vbYesNo) = vbYes Then
        chkSelective.Value = Checked
    Else
        chkSelective.Value = Unchecked
    End If
End If
End Sub

Private Sub cmdLoadExcel_Click()
Dim xlObject    As Excel.Application
Dim xlWB        As Excel.Workbook
Dim fname As String
Dim lastrow, i, j
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
    
    For i = 0 To lstUpdateFields.ListCount - 1
        lstUpdateFields.Selected(i) = False
    Next
    
    With xlObject.ActiveWorkbook.ActiveSheet
         If UCase(Trim(curDomain & "-" & curProject)) <> UCase(Trim(xlObject.ActiveWorkbook.Sheets(2).Range("B7").Value)) Then
            MsgBox "The spreadsheet is from a different Domain or Project"
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
         If InStr(1, GetCommentText(.Range("A1")), "Data Scripting") = 0 Then
            MsgBox "Import file is invalid. Please use only sheets generated by the SuperQC"
            xlWB.Close
            xlObject.Application.Quit
            Set xlWB = Nothing
            Set xlObject = Nothing
            Exit Sub
         End If
         For i = 1 To 12
            If .Range(ColumnLetter(CInt(i)) & 1).Interior.ColorIndex = 35 Then
                For j = 0 To lstUpdateFields.ListCount - 1
                    If .Range(ColumnLetter(CInt(i)) & 1).Value = lstUpdateFields.List(j) Then
                        lstUpdateFields.Selected(j) = True
                    End If
                Next
            End If
         Next
         lastrow = .Range("A" & .Rows.Count).End(xlUp).row
        .Range("A1:L" & lastrow).COPY 'Set selection to Copy
    End With
       
    With flxImport
        .Clear
        .Redraw = False     'Dont draw until the end, so we avoid that flash
        .row = 0            'Paste from first cell
        .col = 0
        .Rows = lastrow
        .Cols = 12
        .RowSel = lastrow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
        .ColSel = 12 - 1
    End With
    
     With flxImport
        .Clear
        .Redraw = False     'Dont draw until the end, so we avoid that flash
        .row = 0            'Paste from first cell
        .col = 0
        .Rows = lastrow
        .Cols = 12
        .RowSel = lastrow - 1 'Select maximum allowed (your selection shouldnt be greater than this)
        .ColSel = 12 - 1
        .Clip = Replace(Clipboard.GetText, vbNewLine, vbCr)   'Replace carriage return with the correct one
        .col = 1            'Just to remove that blue selection from Flexgrid
        .Redraw = True      'Now draw
    End With
        
    xlObject.DisplayAlerts = False 'To avoid "Save woorkbook" messagebox
    
    'Close Excel
    xlWB.Close
    xlObject.Application.Quit
    Set xlWB = Nothing
    Set xlObject = Nothing
    mdiMain.pBar.Max = 100
    mdiMain.pBar.Value = 100
Exit Sub
ErrLoad:
MsgBox "There was an error while importing the file. Please refresh and close all excel and try again" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub cmdSort_Click()
 Dim aData() As String
  Dim lRow As Long, LCol As Long
  Dim cColumn As Collection, cOrder As Collection
  
  flxImport.FixedCols = 0
  
  ' Put the grid data in a string array
  With flxImport
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
  
  cColumn.Add 0
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
  
  flxImport.FixedCols = 1
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
QCTree.height = stsBar.Top - 1150
lstUpdateFields.width = Me.width - lstUpdateFields.Left - 350
flxImport.height = stsBar.Top - flxImport.Top - 250
flxImport.width = Me.width - flxImport.Left - 350
End Sub

Private Sub Label1_Click()
Dim tmpPath, tmpID: On Error Resume Next
tmpID = Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1)
tmpPath = GetFromTable(Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1), "CF_ITEM_ID", "CF_ITEM_PATH", "CYCL_FOLD") & "%"
frmLogs.txtLogs.Text = "Test Instace ID: " & tmpID & vbCrLf & "CF_ITEM_PATH: " & tmpPath & vbCrLf & "Folder Path: " & QCTree.SelectedItem.FullPath
frmLogs.Show 1
End Sub

Private Sub QCTree_DblClick()
Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
Dim nodx As Node

    If QCTree.SelectedItem.Children <> 0 Then Exit Sub
    
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    
    If Left(QCTree.SelectedItem.Key, 1) = "F" Then
        Set objCommand = QCConnection.Command
        objCommand.CommandText = "SELECT DISTINCT CY_CYCLE, CY_CYCLE_ID FROM CYCLE, CYCL_FOLD WHERE CY_FOLDER_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " ORDER BY CY_CYCLE"
        Set rs = objCommand.Execute
        For i = 1 To rs.RecordCount
            QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("C" & rs.FieldValue("CY_CYCLE_ID")), rs.FieldValue("CY_CYCLE"), 2
            rs.Next
        Next
    End If
    
    If Left(QCTree.SelectedItem.Key, 1) = "C" Then
        Set objCommand = QCConnection.Command
        objCommand.CommandText = "SELECT DISTINCT TS_NAME, TC_TESTCYCL_ID, TC_TEST_ORDER FROM TEST, TESTCYCL WHERE TC_CYCLE_ID = " & Right(QCTree.SelectedItem.Key, Len(QCTree.SelectedItem.Key) - 1) & " AND TC_TEST_ID = TS_TEST_ID ORDER BY TC_TEST_ORDER "
        Set rs = objCommand.Execute
        For i = 1 To rs.RecordCount
            QCTree.Nodes.Add QCTree.SelectedItem.Key, tvwChild, CStr("T" & rs.FieldValue("TC_TESTCYCL_ID")), "[" & rs.FieldValue("TC_TEST_ORDER") & "]" & rs.FieldValue("TS_NAME"), 3
            rs.Next
        Next
    End If
End Sub

Private Function GetParamValue(iter, ParName)
Dim i
For i = LBound(AllRunTimeParams(iter).ParamName) To UBound(AllRunTimeParams(iter).ParamName)
    If Trim(UCase(AllRunTimeParams(iter).ParamName(i))) = Trim(UCase(ParName)) Then
        GetParamValue = Replace(AllRunTimeParams(iter).ParamValue(i), "</ROWADATA></DATAPACKET>", "")
        Exit Function
    End If
Next
End Function

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
    cmdSort_Click
    ReNumber
    Exit Sub
OutputErr:
    MsgBox "Data was been truncated because of an error." & vbCrLf & Err.Description
Case "cmdOutput"
    If flxImport.Rows <= 1 Then
        MsgBox "Nothing to output", vbInformation
    Else
        If GetEditableFields() = "IV:IV" Then
            If MsgBox("You have selected nothing to update on this sheet. The whole sheet will be read-only. Do you want to proceed?", vbYesNo) = vbYes Then
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
                OutputTable
            End If
        Else
            stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
            OutputTable
        End If
    End If
Case "cmdUpload"
If IncorrectHeaderDetails = False Then
    If GetEditableFields <> "IV:IV" Then
        If MsgBox("Are you sure you want to mass update " & flxImport.Rows - 1 & " record(s) to the Run Time Parameter?", vbYesNo) = vbYes Then
            'If CheckForWrongOrder = True Then
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(2).Picture: stsBar.Panels(2).Text = "Preparing the process..."
                Randomize: tmpR = CInt(Rnd(1000) * 10000)
                If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
                    If chkSelective.Value = Checked Then
                        Load_Data_To_RunTime_Parameters_SELECTIVE
                    Else
                        If CheckForSelectiveUpload = True Then
                            MsgBox "Selective Upload found in the sheet. Please select Selective Upload"
                            stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Ready"
                            Exit Sub
                        Else
                            Load_Data_To_RunTime_Parameters
                        End If
                    End If
                Else
                    MsgBox "Invalid pass key", vbCritical
                End If
                stsBar.Panels(1).Picture = imgList_Sts.ListImages(1).Picture: stsBar.Panels(2).Text = flxImport.Rows - 1 & " Data Scripts records loaded successfully"
                QCConnection.SendMail "user@companyemail.com", "", "[HPQC UPDATES] Data Scripts records loaded successfully by " & curUser & " in " & curDomain & "-" & curProject, flxImport.Rows - 1 & " Data Scripts records loaded successfully" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
                QCConnection.SendMail curUser, "", "[HPQC UPDATES] Data Scripts records loaded successfully by " & curUser & " in " & curDomain & "-" & curProject, flxImport.Rows - 1 & " Data Scripts records loaded successfully" & "<br><br>" & "Source Data FileName: " & dlgOpenExcel.filename, "", "HTML"
            'End If
        End If
    Else
        MsgBox "No fields to update", vbCritical
    End If
Else
    MsgBox "The template has an invalid/incorrect headers", vbCritical
End If
Case "cmdDeletePassedPars"
    If MsgBox("Are you sure you want to delete all Passed Parameter Data in the script?", vbYesNo) = vbYes Then
        Randomize: tmpR = CInt(Rnd(1000) * 10000)
        If InputBox("Enter pass key '" & tmpR & "'") = tmpR Then
            Remove_Parameter_Values_From_Par
        Else
            MsgBox "Invalid pass key", vbCritical
        End If
    End If
End Select
End Sub

Private Function CheckForSelectiveUpload() As Boolean
Dim i
For i = 1 To flxImport.Rows - 1
    If UCase(Trim(flxImport.TextMatrix(i, 11))) = "YES" Then
        CheckForSelectiveUpload = True
        Exit Function
    End If
Next
End Function

Private Sub ClearForm()
ClearTable
QCTree.Nodes.Clear
'Dim rs As TDAPIOLELib.Recordset
'Dim objCommand
'Dim i As Long
'    QCTree.Nodes.Add , , "Root", "Root"
'    Set objCommand = QCConnection.Command
'    objCommand.CommandText = "SELECT FC_ID, FC_NAME FROM COMPONENT_FOLDER WHERE FC_FATHER_ID = 1 ORDER BY FC_NAME"
'    Set rs = objCommand.Execute
'    For i = 1 To rs.RecordCount
'        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("FC_ID")), rs.FieldValue("FC_NAME")
'        rs.Next
'    Next

Dim rs As TDAPIOLELib.Recordset
Dim objCommand
Dim i As Long
    QCTree.Nodes.Add , , "Root", "Root", 1
    Set objCommand = QCConnection.Command
    objCommand.CommandText = "SELECT CF_ITEM_ID, CF_ITEM_NAME FROM CYCL_FOLD WHERE CF_FATHER_ID = 0 ORDER BY CF_ITEM_NAME"
    Set rs = objCommand.Execute
    For i = 1 To rs.RecordCount
        QCTree.Nodes.Add "Root", tvwChild, CStr("F" & rs.FieldValue("CF_ITEM_ID")), rs.FieldValue("CF_ITEM_NAME"), 1
        rs.Next
    Next
    
    lstUpdateFields.Clear
    lstUpdateFields.AddItem "Actual Value"
     Me.Caption = Me.Tag
     txtFilter.Text = ""
     txtFilter.Locked = True
End Sub

Private Sub ClearTable()
flxImport.Clear
flxImport.TextMatrix(0, 0) = "Test Instance ID"
flxImport.TextMatrix(0, 1) = "RunTime ID"
flxImport.TextMatrix(0, 2) = "Iteration"
flxImport.TextMatrix(0, 3) = "Parameter Order"
flxImport.TextMatrix(0, 4) = "Parameter Name"
flxImport.TextMatrix(0, 5) = "Default Value"
flxImport.TextMatrix(0, 6) = "Actual Value"
flxImport.TextMatrix(0, 7) = "Folder Name"
flxImport.TextMatrix(0, 8) = "Test Set"
flxImport.TextMatrix(0, 9) = "Test Instance"
flxImport.TextMatrix(0, 10) = "Run Time Test ID"
flxImport.TextMatrix(0, 11) = "Selective Upload"
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


On Error GoTo OutErr:  Set xlWB = xlObject.Workbooks.Add

  'xlObject.Visible = True
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
  curTab = "RUN_TIME01"
  xlObject.Sheets("Sheet1").Name = curTab
  flxImport.FixedCols = 0
  flxImport.FixedRows = 0
  flxImport.row = 0
  flxImport.col = 0
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
'     xlObject.Sheets(curTab).Range("A:L").Select
'     xlObject.Sheets(curTab).Range("A2:J" & flxImport.Rows).Sort Key1:=xlObject.Sheets(curTab).Range("A2"), Order1:=xlAscending, Key2:=xlObject.Sheets(curTab).Range("C2") _
'        , Order2:=xlAscending, Key3:=xlObject.Sheets(curTab).Range("D2"), Order3:=xlAscending, Header:= _
'        xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
'        DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, DataOption3:= _
'        xlSortNormal
        
    xlObject.Sheets(curTab).Range("A:L").Borders(xlDiagonalDown).LineStyle = xlNone
    xlObject.Sheets(curTab).Range("A:L").Borders(xlDiagonalUp).LineStyle = xlNone
    With xlObject.Sheets(curTab).Range("A:L").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:L").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:L").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:L").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:L").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlObject.Sheets(curTab).Range("A:L").Borders(xlInsideHorizontal)
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
    xlObject.Sheets(curTab).Range("A:L").Select
    xlObject.Sheets(curTab).Range("A:L").EntireColumn.AutoFit
    xlObject.Sheets(curTab).Range("A1").Select
    
    xlObject.Sheets(curTab).Range("A1").AddComment
    xlObject.Sheets(curTab).Range("A1").Comment.Visible = False
    xlObject.Sheets(curTab).Range("A1").Comment.Text Text:="" & "[" & mdiMain.Caption & "] " & Format(Now, "mmddyyyy HHMMSS AMPM") & ""
  
    xlObject.Sheets(curTab).Range(GetEditableFields & ", L:L").Interior.ColorIndex = 35
    xlObject.Sheets(curTab).Protection.AllowEditRanges.Add Title:="Range1", Range:=xlObject.Sheets(curTab).Columns(GetEditableFields)
    xlObject.Sheets(curTab).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    xlObject.Workbooks(1).SaveAs "RUN_TIME01-" & CleanTheString(QCTree.SelectedItem.Text) & "-" & Format(Now, "mmddyyyy HHMM AMPM")
    xlObject.Visible = True
    xlObject.ActiveWindow.Activate
  
  Set xlWB = Nothing
  Set xlObject = Nothing
  FXGirl.EZPlay FXExportToExcel
  
  stsBar.Panels(1).Picture = imgList_Sts.ListImages(4).Picture: stsBar.Panels(2).Text = "Export to MS Excel completed.": Exit Sub:
OutErr:     MsgBox Err.Description, vbCritical: xlObject.Visible = True: xlObject.ActiveWindow.Activate: Set xlWB = Nothing: Set xlObject = Nothing
On Error GoTo 0
End Sub

Private Function GetEditableFields()
Dim i
Dim j
Dim tmp
For i = 0 To lstUpdateFields.ListCount - 1
   If lstUpdateFields.Selected(i) = True Then
       For j = 0 To flxImport.Cols - 1
          If lstUpdateFields.List(i) = flxImport.TextMatrix(0, j) Then
             tmp = tmp & ColumnLetter(CInt(j + 1)) & ":" & ColumnLetter(CInt(j + 1)) & ", "
          End If
       Next
   End If
Next
tmp = Trim(tmp)
If tmp <> "" Then
   GetEditableFields = Left(tmp, Len(tmp) - 1)
Else
   GetEditableFields = "IV:IV"
End If
End Function

'Public Function CheckForWrongOrder() As Boolean
'Dim i, Override As Boolean
'For i = 1 To flxImport.Rows - 1
'    If flxImport.TextMatrix(i, 3) = flxImport.TextMatrix(i + 1, 3) And flxImport.TextMatrix(i, 3) <> "1" Then
'        If Override = True Then
'            flxImport.TextMatrix(i + 1, 3) = flxImport.TextMatrix(i, 3) + 1
'        ElseIf MsgBox("Duplicate Parameter Order Found! Do you want to fix it?" & vbCrLf & "Line Number: " & i, vbYesNo) = vbYes And Override = False Then
'            flxImport.TextMatrix(i + 1, 3) = flxImport.TextMatrix(i, 3) + 1
'            Override = True
'        Else
'            CheckForWrongOrder = False
'            Exit Function
'        End If
'    End If
'Next
'CheckForWrongOrder = True
'End Function

Public Sub ReNumber()
Dim i As Integer
Dim cntr
cntr = 1
For i = 1 To flxImport.Rows - 1
    If i = 2 Then
        flxImport.TextMatrix(i, 3) = cntr
        cntr = cntr + 1
    Else
        If flxImport.TextMatrix(i, 0) = flxImport.TextMatrix(i - 1, 0) And flxImport.TextMatrix(i, 2) = flxImport.TextMatrix(i - 1, 2) Then
            flxImport.TextMatrix(i, 3) = cntr
            cntr = cntr + 1
        Else
            cntr = 1
            flxImport.TextMatrix(i, 3) = cntr
            cntr = cntr + 1
        End If
    End If
Next
End Sub


Private Sub Remove_Parameter_Values_From_Par()
Dim i, totalP As Integer
On Error Resume Next
For i = 1 To flxImport.Rows - 1
    If InStr(1, flxImport.TextMatrix(i, 6), "[") <> 0 And InStr(1, flxImport.TextMatrix(i, 6), "]") <> 0 And InStr(1, flxImport.TextMatrix(i, 6), "{") <> 0 And InStr(1, flxImport.TextMatrix(i, 6), "}") <> 0 Then
        flxImport.TextMatrix(i, 6) = Left(flxImport.TextMatrix(i, 6), InStr(1, flxImport.TextMatrix(i, 6), "}"))
        totalP = totalP + 1
    End If
Next
MsgBox "There are " & totalP & " passed parameter data removed"
End Sub

Private Function GetTestSetFolderPath(strID As String) As String
Dim Fact As TestSetFactory
Dim Obj As TestSet
Set Fact = QCConnection.TestSetFactory
Set Obj = Fact.Item(strID)
GetTestSetFolderPath = Obj.TestSetFolder.path
End Function

Private Function IncorrectHeaderDetails() As Boolean
    If flxImport.TextMatrix(0, 0) <> "Test Instance ID" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 1) <> "RunTime ID" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 2) <> "Iteration" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 3) <> "Parameter Order" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 4) <> "Parameter Name" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 5) <> "Default Value" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 6) <> "Actual Value" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 7) <> "Folder Name" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 8) <> "Test Set" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 9) <> "Test Instance" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 10) <> "Run Time Test ID" Then IncorrectHeaderDetails = True
    If flxImport.TextMatrix(0, 11) <> "Selective Upload" Then IncorrectHeaderDetails = True
End Function

Private Sub txtFilter_DblClick()
txtFilter.Locked = False
End Sub
