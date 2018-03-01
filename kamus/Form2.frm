VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Atur Transparant Form"
   ClientHeight    =   1200
   ClientLeft      =   6645
   ClientTop       =   5160
   ClientWidth     =   2775
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtLevel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "255"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MAX 255 (Ojo Luwih Tekan Iki )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
On Error Resume Next
TransForm frmMain, txtLevel.Text
    Load frmMain
    frmMain.Show
    Me.Hide
End Sub

Private Sub Form_Load()
    TransForm Me, 255
     RemoveCancelMenuItem Me
End Sub

Private Sub txtLevel_Change()
        On Error Resume Next
    TransForm Me, txtLevel.Text
End Sub


Private Sub txtLevel_Validate(Cancel As Boolean)
    If txtLevel.Text >= 256 Then
       Cancel = txtLevel.Text >= 256
        MsgBox "Harus Kurang dari 255", vbCritical
        End If
End Sub
