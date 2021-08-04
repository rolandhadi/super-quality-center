VERSION 5.00
Begin VB.Form frmLogs 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Logs"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   10485
   Icon            =   "frmLogs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Close"
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
      Left            =   8940
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox txtLogs 
      Height          =   7035
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   10275
   End
End
Attribute VB_Name = "frmLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
cmdClose.Left = Me.width - cmdClose.width - 175
cmdClose.Top = Me.height - cmdClose.height - 650
txtLogs.height = cmdClose.Top - txtLogs.Top - 100
txtLogs.width = Me.width - txtLogs.Left - 175
End Sub
