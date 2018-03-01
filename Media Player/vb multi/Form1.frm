VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2400
      Top             =   4920
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "Open"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   4800
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Multimedia Player RPL 1.0.1"
      BeginProperty Font 
         Name            =   "MurrayHill Bd BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   6855
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6855
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   12091
      _cy             =   7223
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Red, Green, Blue As Integer


Private Sub CmdOpen_Click()
 On Error Resume Next
 CommonDialog1.ShowOpen
 WindowsMediaPlayer1.URL = CommonDialog1.FileName
End Sub


Private Sub Timer1_Timer()
    

If Blue <= 255 Then Blue = Blue + 50 Else Blue = 0
    Green = Green + 50
If Green >= 255 Then Green = 0
    Red = Red + 50
If Red >= 255 Then
    Red = 0

Label1.ForeColor = Int(RGB(Red, Green, Blue))
Label1.Refresh
End If

End Sub

