VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copy Text"
   ClientHeight    =   3636
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   5376
   DrawMode        =   7  'Invert
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3636
   ScaleWidth      =   5376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ingat Copy Dulu"
      Height          =   852
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3492
      Begin VB.Label Label3 
         Caption         =   "2. Lalu Tekan Copy"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2052
      End
      Begin VB.Label Label2 
         Caption         =   "1. Alt+A Pada Text Dibawah"
         Height          =   252
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3012
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   372
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   1332
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2412
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5172
      _ExtentX        =   9123
      _ExtentY        =   4255
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form3.frx":0000
   End
   Begin VB.Label Label1 
      Caption         =   "2. Tolong CopyText Lalu Paste Di Form Utama"
      Height          =   252
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   3372
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
   Clipboard.SetText RichTextBox1.SelText
Form3.Hide
End Sub


