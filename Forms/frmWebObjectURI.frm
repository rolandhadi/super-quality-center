VERSION 5.00
Begin VB.Form frmWebObjectURI 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Web Object URI"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkIndex 
      Caption         =   "Index"
      Height          =   315
      Left            =   4740
      TabIndex        =   7
      Top             =   1020
      Width           =   855
   End
   Begin VB.TextBox txtURI 
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
      Left            =   60
      TabIndex        =   5
      Top             =   1380
      Width           =   5475
   End
   Begin VB.ComboBox cmbType 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmWebObjectURI.frx":0000
      Left            =   2040
      List            =   "frmWebObjectURI.frx":0037
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   3555
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtObjectName 
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
      Left            =   2040
      TabIndex        =   0
      Top             =   540
      Width           =   3555
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   810
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Object Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   585
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web Object Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   1620
   End
End
Attribute VB_Name = "frmWebObjectURI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myIndex As Integer

Private Sub chkIndex_Click()
UpdateProperty
End Sub

Private Sub cmbType_Change()
UpdateProperty
End Sub

Private Sub cmbType_Click()
UpdateProperty
End Sub

Private Sub cmbType_KeyUp(KeyCode As Integer, Shift As Integer)
UpdateProperty
End Sub

Private Sub cmdOK_Click()
frmControlRunTime.txtProperties(myIndex).Text = Trim(txtURI.Text)
Unload Me
End Sub

Private Sub Form_Load()
Dim fileFunct As New clsFiles
Dim WinFunct As New clsWindow
Dim tmp, i, tmpStr
WinFunct.Ontop Me
End Sub

Private Sub txtObjectName_Change()
UpdateProperty
End Sub

Sub UpdateProperty()
If chkIndex.Value = Checked Then
    txtURI.Text = "micclass:=" & cmbType.Text & ";name:=" & txtObjectName.Text & ";index:=0"
Else
    txtURI.Text = "micclass:=" & cmbType.Text & ";name:=" & txtObjectName.Text
End If
End Sub
