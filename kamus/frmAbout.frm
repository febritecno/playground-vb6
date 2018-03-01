VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H000080FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   4995
   ClientLeft      =   5835
   ClientTop       =   3645
   ClientWidth     =   3750
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3447.637
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.444
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   360
      Top             =   4320
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3240
      Top             =   4320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mungkin Aplikasi jelek dan sederhana ini bermanfaat buat Anda, Semoga Anda nyaman menggunakan aplikasi jelek buatan saya ini."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CREDIT# By Febrian Dwi Putra"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   3225
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "+/Spesial Thanks :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kamus Version 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Red, green, blue As Integer



Private Sub Timer1_Timer()
If blue <= 255 Then blue = blue + 50 Else blue = 0
    green = green + 50
If green >= 255 Then green = 0
    Red = Red + 50
If Red >= 255 Then
    Red = 0

Label4.ForeColor = Int(RGB(Red, green, blue))
Label4.Refresh
End If

End Sub

Private Sub Timer2_Timer()
If blue <= 255 Then blue = blue + 50 Else blue = 0
green = green + 50
If green >= 255 Then green = 0
Red = Red + 50
If Red >= 255 Then
Red = 0
End If

Label5.ForeColor = Int(RGB(Red, green, blue))
Label5.Refresh
End Sub

