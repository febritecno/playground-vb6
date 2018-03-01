VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Media Player"
   ClientHeight    =   5712
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   9828
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleMode       =   0  'User
   ScaleWidth      =   1338.235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Keluar"
      DragIcon        =   "form1.frx":0000
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      MaskColor       =   &H000000FF&
      TabIndex        =   7
      Top             =   5280
      Width           =   2052
   End
   Begin VB.ListBox List2 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00C00000&
      Height          =   1428
      Left            =   7680
      TabIndex        =   5
      Top             =   3600
      Width           =   2172
   End
   Begin VB.ListBox List1 
      BackColor       =   &H000040C0&
      ForeColor       =   &H00000000&
      Height          =   1428
      Left            =   7680
      TabIndex        =   4
      Top             =   3600
      Width           =   2172
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   7.2
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1752
      Left            =   7680
      TabIndex        =   2
      Top             =   480
      Width           =   2172
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1224
      Left            =   7680
      TabIndex        =   1
      Top             =   2280
      Width           =   2172
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Cooper Std Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   384
      Left            =   7680
      TabIndex        =   0
      Top             =   0
      Width           =   2172
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   5175
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7695
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
      _cx             =   13568
      _cy             =   9123
   End
   Begin VB.Label febri 
      BackColor       =   &H0000FFFF&
      Caption         =   "BY:Febrian Dwi Putra (10-RPL)"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   5160
      Width           =   7680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_DblClick()
List1.AddItem Dir1.Path & "\" & File1.FileName
List2.AddItem File1.FileName
End Sub
Private Sub List2_DblClick()
List1.ListIndex = List2.ListIndex
WindowsMediaPlayer1.URL = List1.List(List1.ListIndex)
End Sub
Private Sub WindowsMediaPlayer1_EndOfStream(ByVal Result As Long)
List1.ListIndex = List1.ListIndex + 1
List2.ListIndex = List2.ListIndex + 1
WindowsMediaPlayer1.URL = List1.List(List1.ListIndex)
End Sub
Private Sub Timer1_Timer()
BackColor = RGB(Rnd() * 225, Rnd() * 225, Rnd() * 225)
End Sub

