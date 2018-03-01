VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00EAF7F7&
   BorderStyle     =   0  'None
   Caption         =   "-------_-+ :www.FerweB\Inc.Com:+_--------"
   ClientHeight    =   8790
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   ScaleHeight     =   752
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1274
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   5040
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   4440
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   7
      Top             =   240
      Width           =   2415
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   10575
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   19095
      ExtentX         =   33681
      ExtentY         =   18653
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http://www.mysearch.com/jsp/cfg_redir2.jsp?id=NE&psa=FEB63364-EF49-4D2C-9D90-CA5BC52ED6CE&url=http://www.ask.com/web&l=dir&q="
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Masuk"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6840
      TabIndex        =   4
      Text            =   "Masukan Alamat Yang DiTuju Oleh Anda !!!!!"
      Top             =   240
      Width           =   5415
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   13560
      OleObjectBlob   =   "Form1.frx":0088
      Top             =   120
   End
   Begin VB.CommandButton Command4 
      Caption         =   "STOP"
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT "
      BeginProperty Font 
         Name            =   "CarrickGroovy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "CarrickGroovy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":103E49
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label jam 
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2055
   End
   Begin VB.Line Line3 
      BorderColor     =   &H008080FF&
      X1              =   288
      X2              =   288
      Y1              =   0
      Y2              =   48
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   936
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   936
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Menu cmdfile 
      Caption         =   "&File"
      Begin VB.Menu cmdbuka 
         Caption         =   "&Buka"
         Shortcut        =   ^B
      End
      Begin VB.Menu cmdsimpan 
         Caption         =   "&Simpan"
         Shortcut        =   ^S
      End
      Begin VB.Menu cmdkeluar 
         Caption         =   "&Keluar"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu cmdoption 
      Caption         =   "&Option"
      Begin VB.Menu cmdfacebook 
         Caption         =   "Facebook Auto Comment"
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdbuka_Click()
CommonDialog1.CancelError = True
On Error GoTo cancel
CommonDialog1.Filter = "File HTM|*.HTM|File HTML|*.HTML"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
WebBrowser1.Navigate CommonDialog1.FileName
Text1.Text = CommonDialog1.FileName
End If
Exit Sub
cancel:
Exit Sub
End Sub

Private Sub cmdsimpan_Click()
CommonDialog2.CancelError = True
On Error GoTo cancel
CommonDialog2.Filter = "File HTM|*.HTM|File HTML|*.HTML"
CommonDialog2.ShowSave
If CommonDialog2.FileName <> "" Then
strnamafile = CommonDialog2.FileName
intnofile = FreeFile
Open strnamafile For Output As intnofile
Print #intnofile, Inet1.OpenURL(Text1.Text)
Close intnofile
End If
Exit Sub
cancel:
Exit Sub
End Sub

Private Sub cmdkeluar_Click()
Unload Me
End Sub

Private Sub cmdfacebook_Click()
Frmfacebook.Show
End Sub


Private Sub Command1_Click(Index As Integer)
On Error GoTo errorback
WebBrowser1.GoBack
Exit Sub
errorback:
MsgBox "Tidak Ada Halaman Sebelumnya!"
End Sub

Private Sub Command2_Click()
On Error GoTo errorback
WebBrowser1.GoBack
Exit Sub
errorback:
MsgBox "Tidak Ada Halaman Sebelumnya!"
End Sub
Private Sub Command3_Click()
WebBrowser1.Refresh
Text1.Text = WebBrowser1.LocationURL
End Sub
Private Sub Command4_Click()
WebBrowser1.Stop
End Sub
Private Sub Command5_Click()
On Error Resume Next
WebBrowser1.Navigate (Text1.Text)
End Sub

Private Sub Form_Load()
Skin1.ApplySkin Me.hWnd
End Sub
Private Sub Timer1_Timer()
jam.Caption = DateTime.Time
tanggal.Caption = DateTime.Date
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
If Progress = -1 Then ProgressBar1.Value = 100
Label1.Caption = "Done"
ProgressBar1.Visible = False
If Progress > 0 And ProgressMax > 0 Then
ProgressBar1.Visible = True
Image1.Visible = False
ProgressBar1.Value = Progress * 100 / ProgressMax
Label1.Caption = "Loading... " & Int(Progress * 100 / ProgressMax) & "%"
End If
Exit Sub
End Sub



