VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Log In"
   ClientHeight    =   1425
   ClientLeft      =   5700
   ClientTop       =   5040
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   1425
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Masuk"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   7
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox TxtUser 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1080
      Top             =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Username :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2010
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Form_Load()
On Error Resume Next
txtPassword.MaxLength = 7
TxtUser.MaxLength = 5
End Sub

Private Sub Command1_Click()
If TxtUser.Text = "febri" And txtPassword.Text = "blogger" Then
MsgBox "Selamat Anda Bisa Masuk !!", vbInformation
Form1.Hide
frmMain.Show
TxtUser.Text = ""
txtPassword.Text = ""
frmMain.Command1.Enabled = True
frmMain.login.Visible = False
frmMain.keluar.Visible = True
frmMain.Height = 5655
Else
MsgBox "Rasain Loo..?? User dan Password Salah", vbCritical
Command1.Enabled = False
TxtUser.Text = ""
txtPassword.Text = ""
TxtUser.SetFocus
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
Me.Hide
End Sub

Private Sub Timer1_Timer()
Form1.Left = Form1.Left - 15
If Form1.Left <= -Form1.Left Then
Form1.Left = Form1.Width
End If
End Sub

Private Sub txtPassword_Change()
If txtPassword.Text = "blogger" Then
Command1.Enabled = True
End If
End Sub

