VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Html Viewer Editor"
   ClientHeight    =   6096
   ClientLeft      =   108
   ClientTop       =   732
   ClientWidth     =   6216
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6096
   ScaleWidth      =   6216
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3132
      Left            =   0
      TabIndex        =   1
      Top             =   6
      Width           =   6132
      ExtentX         =   10816
      ExtentY         =   5524
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
      Location        =   "http:///"
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2532
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   6132
      _ExtentX        =   10816
      _ExtentY        =   4466
      _Version        =   393217
      BackColor       =   12640511
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":6852
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu baru 
         Caption         =   "&Baru"
         Checked         =   -1  'True
         Shortcut        =   ^N
      End
      Begin VB.Menu space0 
         Caption         =   "-"
      End
      Begin VB.Menu buka 
         Caption         =   "&Buka"
         Shortcut        =   ^O
      End
      Begin VB.Menu simpan 
         Caption         =   "&Simpan"
         Shortcut        =   ^S
      End
      Begin VB.Menu print 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu keluar 
         Caption         =   "&Keluar"
      End
   End
   Begin VB.Menu insert 
      Caption         =   "&Insert"
      Begin VB.Menu Field 
         Caption         =   "Tambah Field"
         Checked         =   -1  'True
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu fon 
         Caption         =   "&Font"
         Begin VB.Menu font 
            Caption         =   "&Font Type"
            Begin VB.Menu Calibri 
               Caption         =   "&Calibri"
            End
         End
         Begin VB.Menu Gaya 
            Caption         =   "&Font Gaya"
         End
         Begin VB.Menu ukuran 
            Caption         =   "&Font Size "
         End
      End
      Begin VB.Menu Background 
         Caption         =   "&Warna Background"
         Begin VB.Menu default 
            Caption         =   "&Default"
            Checked         =   -1  'True
         End
         Begin VB.Menu hitam 
            Caption         =   "&Hitam"
         End
         Begin VB.Menu biru 
            Caption         =   "&Biru"
         End
         Begin VB.Menu merah 
            Caption         =   "&Merah"
         End
         Begin VB.Menu kuning 
            Caption         =   "&Kuning"
         End
         Begin VB.Menu hijau 
            Caption         =   "&Hijau"
         End
      End
   End
   Begin VB.Menu alat 
      Caption         =   "&Alat"
      Begin VB.Menu garis 
         Caption         =   "Garis"
      End
      Begin VB.Menu tulis 
         Caption         =   "Tulisan"
         Begin VB.Menu tebal 
            Caption         =   "Tulisan Tebal"
         End
         Begin VB.Menu miring 
            Caption         =   "Tulisan Miring"
         End
         Begin VB.Menu bawah 
            Caption         =   "Tulisan Underline"
         End
         Begin VB.Menu kecil 
            Caption         =   "Tulisan Kecil"
         End
         Begin VB.Menu kuat 
            Caption         =   "Tulisan Kuat"
         End
      End
      Begin VB.Menu heading 
         Caption         =   "Heading"
         Begin VB.Menu h1 
            Caption         =   "H1"
         End
         Begin VB.Menu h2 
            Caption         =   "H2"
         End
         Begin VB.Menu h3 
            Caption         =   "H3"
         End
         Begin VB.Menu h4 
            Caption         =   "H4"
         End
         Begin VB.Menu h5 
            Caption         =   "H5"
         End
      End
   End
   Begin VB.Menu text 
      Caption         =   "&Text"
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu delete 
         Caption         =   "&Delete"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
Form2.Show
End Sub

Private Sub baru_Click()
Call simpan_Click
WebBrowser1.Navigate "about:blank"
RichTextBox1.text = ""
Form3.Show
End Sub

Private Sub bawah_Click()
RichTextBox1.SelText = "<u></u>"
End Sub

Private Sub biru_Click()
Form1.BackColor = vbBlue
End Sub

Private Sub buka_Click()
CommonDialog1.Filter = "Semua Files |*.*|Html Files |*.html"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
Dim TxtBox As Object
Dim load As Boolean
Dim path As String
Dim file As Integer
Dim s As String
Dim ape As Boolean
If Dir(path) = "" Then Exit Sub
On Error GoTo ErrorHandler:
s = RichTextBox1.text
file = FreeFile
Open CommonDialog1.FileName For Input As #file
s = Input(LOF(file), #file)
If ape Then
RichTextBox1.text = RichTextBox1.text & s
Else
RichTextBox1.text = s
End If
load = True
ErrorHandler:
If file > 0 Then Close #file
End Sub

Private Sub Calibri_Click()
RichTextBox1.font = "Cooper"
End Sub

Private Sub copy_Click()
On Error Resume Next
   Clipboard.SetText RichTextBox1.SelText
End Sub

Private Sub cut_Click()
On Error Resume Next
Clipboard.SetText RichTextBox1.SelText
RichTextBox1.SelText = vbNullString
End Sub

Private Sub delete_Click()
Clipboard.Clear
RichTextBox1.SelText = ""
End Sub

Private Sub Field_Click()
LoadNewDoc
End Sub

Private Sub LoadNewDoc()
    Static lDocumentCount As Long
    Dim frmD As Form1
    lDocumentCount = lDocumentCount + 1
    Set frmD = New Form1
    frmD.Caption = "Field" & lDocumentCount
    frmD.Show
End Sub


Private Sub Form_Load()
WebBrowser1.Navigate "about:blank"
End Sub

Private Sub Form_Resize()
WebBrowser1.Top = 40
WebBrowser1.Left = 40
WebBrowser1.Width = Me.Width - 300
WebBrowser1.Height = Me.Height / 2 - 80

RichTextBox1.Top = Me.Height / 2 + 100
RichTextBox1.Left = 40
RichTextBox1.Width = Me.Width - 300
RichTextBox1.Height = Me.Height / 2 - 950
End Sub

Private Sub garis_Click()
RichTextBox1.SelText = "<hr/>"
End Sub

Private Sub h1_Click()
RichTextBox1.SelText = "<h1></h1>"
End Sub

Private Sub h2_Click()
RichTextBox1.SelText = "<h2></h2>"
End Sub

Private Sub h3_Click()
RichTextBox1.SelText = "<h3></h3>"
End Sub

Private Sub h4_Click()
RichTextBox1.SelText = "<h4></h4>"
End Sub

Private Sub h5_Click()
RichTextBox1.SelText = "<h5></h5>"
End Sub

Private Sub hitam_Click()
Form1.BackColor = vbBlack
End Sub

Private Sub kecil_Click()
RichTextBox1.SelText = "<Small></Small>"
End Sub


Private Sub keluar_Click()
End
End Sub

Private Sub kuat_Click()
RichTextBox1.SelText = "<Strong></Strong>"
End Sub

Private Sub miring_Click()
RichTextBox1.SelText = "<i></i>"
End Sub


Private Sub paste_Click()
On Error Resume Next
RichTextBox1.SelText = Clipboard.GetText
End Sub

Private Sub print_Click()
 On Error GoTo ErrHandler
  Dim BeginPage, EndPage, NumCopies, i
   CommonDialog1.CancelError = True
  CommonDialog1.ShowPrinter
  BeginPage = CommonDialog1.FromPage
  EndPage = CommonDialog1.ToPage
  NumCopies = CommonDialog1.Copies
  For i = 1 To NumCopies
 Printer.Print RichTextBox1.text
  Next i
  Exit Sub
ErrHandler:
   Exit Sub
End Sub

Private Sub RichTextBox1_Change()
On Error Resume Next
DoEvents
Open "C:\temp.html" For Output As #1: Print #1, RichTextBox1.text: Close #1
DoEvents
WebBrowser1.Navigate "C:\temp.html"
End Sub

Private Sub simpan_Click()
On Error GoTo ErrorHandler
  CommonDialog1.Filter = "Semua Files |*.*|Html Files |*.html"
    CommonDialog1.FilterIndex = 2
   CommonDialog1.ShowSave
 CommonDialog1.FileName = CommonDialog1.FileName
Dim iFile As Integer
 Dim SaveFileFromTB As Boolean
 Dim TxtBox As Object
 Dim FilePath As String
Dim Append As Boolean
  iFile = FreeFile
If Append Then
    Open CommonDialog1.FileName For Append As #iFile
Else
    Open CommonDialog1.FileName For Output As #iFile
End If
Print #iFile, RichTextBox1.text
SaveFileFromTB = True
ErrorHandler:
Close #iFile
End Sub
Private Sub Form_unload(Cancel As Integer)
On Error GoTo ErrorHandler
Dim Msg, Style, Title, Response, MyString
Msg = "Mau Keluar !! => Ingat Save Data Anda Y/N?"
Style = vbYesNo + vbCritical + vbDefaultButton1
Title = "Peringatan"
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
   MyString = "Yes"
   Call simpan_Click
   End
If Response = vbNo Then
   MyString = "No"
   End
End If
ErrorHandler:
Cancel = 1
End If
End Sub

Private Sub hijau_Click()
Form1.BackColor = vbGreen
End Sub

Private Sub merah_Click()
Form1.BackColor = vbRed
End Sub

Private Sub default_Click()
Form1.BackColor = vbWhite
End Sub

Private Sub kuning_Click()
Form1.BackColor = vbYellow
End Sub
Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    LoadNewDoc
End Sub

Private Sub tebal_Click()
RichTextBox1.SelText = "<b></b>"
End Sub
