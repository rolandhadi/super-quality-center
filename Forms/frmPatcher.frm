VERSION 5.00
Begin VB.Form frmPatcher 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdate 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmPatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Visible = False
End Sub

Private Sub tmrUpdate_Timer()
  Dim FileFunct As New clsFiles
  Dim UpdatedF As String
  Dim CurF As String
  UpdatedF = FileFunct.ReadKeyFromFile(App.Path & "\" & "myPatch.hxh", "<PATCHPATH_UPDATE>")
  CurF = FileFunct.ReadKeyFromFile(App.Path & "\" & "myPatch.hxh", "<PATCHPATH_CUR>")
  'UpdatedF = "C:\Users\user\Documents\Work Related\Super Quality Center R4\SQC DAT\_AUTO_UPDATE_SQC_\attachStorage\TEST_2_Super Quality Center R4v1.exe"
  'CurF = "C:\Users\user\Documents\Work Related\Super Quality Center R4\Super Quality Center R4v1.exe"
  If Trim(UpdatedF) = "" Or Trim(CurF) = "" Then End
    While FileFunct.FileExists(CurF) = True
      FileFunct.FileDelete CurF
    Wend
    FileFunct.FileCopy UpdatedF, CurF
    Shell CurF, vbNormalFocus
  End
End Sub
