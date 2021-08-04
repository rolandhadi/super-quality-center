VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "SUPER QUALITY CENTER ULTIMATE EDITION"
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   17310
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdiMain.frx":08CA
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar pBar 
      Align           =   1  'Align Top
      Height          =   180
      Left            =   0
      TabIndex        =   1
      Top             =   1185
      Width           =   17310
      _ExtentX        =   30533
      _ExtentY        =   318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   13500
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Excel Files | *.xls"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13500
      Top             =   5460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   80
      ImageHeight     =   67
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":68681
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":69582
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6A254
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6AC3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6B639
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6CA1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6D781
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6E453
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6F0B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6FEF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":70E0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":71FCB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1185
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17310
      _ExtentX        =   30533
      _ExtentY        =   2090
      ButtonWidth     =   2302
      ButtonHeight    =   1931
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RQ"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "UP_UP"
                  Text            =   "Update Requirements"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "UP_R"
                  Text            =   "Upload Requirements"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "UP_RL"
                  Text            =   "Link Requirements"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "UP_UL"
                  Text            =   "UnLink Requirement"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TAO"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "B_A_C"
                  Text            =   "Build and Consolidate"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "P_M"
                  Text            =   "Procedures Management Module"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DT_M"
                  Text            =   "DataTable Manager"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "T_S"
                  Text            =   "Taylor Swift"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BC"
            ImageIndex      =   6
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "U_B"
                  Text            =   "Update Business Components"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "D_B"
                  Text            =   "Download/Upload Business Components Steps"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "UP_B"
                  Text            =   "Upload Business Component"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BC_PAR"
                  Text            =   "Update Business Components Parameters"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TP"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "T_UP"
                  Text            =   "Update Test Plan Test"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "T_UBO"
                  Text            =   "Update Component Order"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "T_UPP"
                  Text            =   "Update Test Parameters"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "T_CBPT"
                  Text            =   "Create Test Plan BPT"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "T_LBPT"
                  Text            =   "Link Component To Test"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TL"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TL_TS"
                  Text            =   "Update Test Set"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TL_TI"
                  Text            =   "Update Test Instance"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TL_CT"
                  Text            =   "Create Test Set"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "TL_LT"
                  Text            =   "Link Test Plan to Lab (Quick)"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TL_LTP"
                  Text            =   "Link Test Plan to Lab (by Path)"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "U_RS"
                  Text            =   "Update Run Steps"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DS"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RP"
            ImageIndex      =   10
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RP_EX"
                  Text            =   "Extreme Report Generator"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UL"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "US"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RS_P"
                  Text            =   "Change User Properties"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AB"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Activate()
SetHeader
End Sub

Private Sub MDIForm_DblClick()
If LCase(curUser) = "Roland Ross Hadi" And MsgBox("Are you sure you want to do it?", vbYesNo) = vbYes Then
  DownloadAttachments GetTest(GetFromTable("'_AUTO_UPDATE_SQC_'", "TS_NAME", "TS_TEST_ID", "TEST")): MsgBox "Done"
End If
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
MkDir App.path & "\SQC DAT"
MkDir App.path & "\SQC Reports"
MkDir App.path & "\SQC Logs"
MkDir App.path & "\SQC DAT\FX"
SetHeader
If ACCESS_ = "Team" Then
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(3).Visible = True
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = True
    Toolbar1.Buttons(8).Visible = True
    Toolbar1.Buttons(9).Visible = False
    Toolbar1.Buttons(3).ButtonMenus(1).Visible = False
    Toolbar1.Buttons(4).ButtonMenus(1).Visible = False
    Toolbar1.Buttons(4).ButtonMenus(2).Visible = False
    Toolbar1.Buttons(4).ButtonMenus(3).Visible = False
    Toolbar1.Buttons(5).ButtonMenus(1).Visible = False
    Toolbar1.Buttons(5).ButtonMenus(2).Visible = False
    Toolbar1.Buttons(5).ButtonMenus(6).Visible = False
ElseIf ACCESS_ = "User" Then
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(8).Visible = False
    Toolbar1.Buttons(9).Visible = True
ElseIf ACCESS_ = "Reporter" Then
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
    Toolbar1.Buttons(7).Visible = True
    Toolbar1.Buttons(8).Visible = False
    Toolbar1.Buttons(9).Visible = False
ElseIf ACCESS_ = "Auto" Then
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(2).ButtonMenus(1).Visible = False
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(8).Visible = True
    Toolbar1.Buttons(9).Visible = False
End If
 FXSQCExtractCompleted = App.path & "\SQC DAT\FX\SQCExtractCompleted.wav"
 FXScriptConsolidationCompleted = App.path & "\SQC DAT\FX\ScriptConsolidationCompleted.wav"
 FXExportToExcel = App.path & "\SQC DAT\FX\ExportToExcel.wav"
 FXDataUploadCompleted = App.path & "\SQC DAT\FX\DataUploadCompleted.wav"
 FX25 = App.path & "\SQC DAT\FX\25.wav"
 FX50 = App.path & "\SQC DAT\FX\50.wav"
 FX75 = App.path & "\SQC DAT\FX\75.wav"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim PassCode
QCConnection.RefreshConnectionState
If QCConnection.Connected = False Then
    On Error Resume Next
    QCConnection.InitConnectionEx (curQCInstance)
    QCConnection.Logout
    QCConnection.Login ADMIN_ID, ADMIN_PASS
    QCConnection.CONNECT curDomain, curProject
    If QCConnection.Connected = False Then MsgBox "Session was disconnected. Please re-load the SuperQC tool.": If MsgBox("Do you want to close this session?", vbYesNo) = vbYes Then End
End If
If LatestVersionCheck = False And InStr(1, App.EXEName, "prjSuper QualityCenterExplorerUltimate", vbTextCompare) = 0 And curInstance <> "Release 3" Then
    If MsgBox("An updated version of superQC was detected in the server. Do you want to restart the application and update to the latest version of superQC?", vbYesNo) = vbYes Then
        Patch
    End If
End If
If Button.Key = "DS" Then
    frmDataScripting.Show
ElseIf Button.Key = "UL" Then
    frmUnlock.Show
ElseIf Button.Key = "RP" Then
    frmReports.Show
ElseIf Button.Key = "TS" Then
    frmBusinessComponents.Show
ElseIf Button.Key = "US" Then
    frmLoadUsers.Show
ElseIf Button.Key = "AB" Then
    frmAbout.Show 1
End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.Key = "TL_TI" Then
    frmTestInstance.Show
ElseIf ButtonMenu.Key = "D_B" Then
    frmAllBusinessComponents.Show
ElseIf ButtonMenu.Key = "UP_B" Then
    frmLoadManualBC.Show
ElseIf ButtonMenu.Key = "TL_TS" Then
    frmTestSet.Show
ElseIf ButtonMenu.Key = "U_B" Then
    frmBusinessComponents.Show
ElseIf ButtonMenu.Key = "T_CBPT" Then
    frmCreateTestPlanBPT.Show
ElseIf ButtonMenu.Key = "T_LBPT" Then
    frmLinkTestPlanBPT.Show
ElseIf ButtonMenu.Key = "TL_CT" Then
    frmCreateTestLabBPT.Show
ElseIf ButtonMenu.Key = "TL_LT" Then
    frmLinkTestLabBPT.Show
ElseIf ButtonMenu.Key = "TL_LTP" Then
    frmLinkTestLabBPT_ViaPath.Show
ElseIf ButtonMenu.Key = "UP_R" Then
    frmLoadRequirement.Show
ElseIf ButtonMenu.Key = "UP_RL" Then
    frmLinkTestPlanReq.Show
ElseIf ButtonMenu.Key = "T_UP" Then
    frmTest.Show
ElseIf ButtonMenu.Key = "UP_UP" Then
    frmRequirements.Show
ElseIf ButtonMenu.Key = "T_UBO" Then
    frmBcInTestOrder.Show
ElseIf ButtonMenu.Key = "RP_EX" Then
    frmReportGen.Show
ElseIf ButtonMenu.Key = "UP_UL" Then
    frmUnLinkTestPlanReq.Show
ElseIf ButtonMenu.Key = "T_UPP" Then
    frmTestParameters.Show
ElseIf ButtonMenu.Key = "BC_PAR" Then
    frmBCParameters.Show
ElseIf ButtonMenu.Key = "B_A_C" Then
    frmConsolidate.Show
ElseIf ButtonMenu.Key = "P_M" Then
    frmFunctionControl.Show
ElseIf ButtonMenu.Key = "U_RS" Then
    frmRunSteps.Show
ElseIf ButtonMenu.Key = "DT_M" Then
    frmDataTableManager.Show
ElseIf ButtonMenu.Key = "T_S" Then
    frmSinger.Show
ElseIf ButtonMenu.Key = "RS_P" Then
    frmUpdateUsers.Show
Else
    '
End If
End Sub
