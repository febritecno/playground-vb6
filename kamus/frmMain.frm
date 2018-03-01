VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Kamus Istilah Bahasa IT"
   ClientHeight    =   3465
   ClientLeft      =   4080
   ClientTop       =   4020
   ClientWidth     =   10575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":1762
   ScaleHeight     =   3465
   ScaleWidth      =   10575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Unlock"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   615
      Left            =   2280
      Picture         =   "frmMain.frx":23A4
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Save"
      Top             =   3720
      Width           =   1995
   End
   Begin VB.CommandButton cmdNew 
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      Picture         =   "frmMain.frx":25F4
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "New"
      Top             =   3720
      Width           =   1995
   End
   Begin VB.CommandButton cmdUpdate 
      Enabled         =   0   'False
      Height          =   615
      Left            =   6240
      Picture         =   "frmMain.frx":2813
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Update"
      Top             =   3720
      Width           =   1995
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   8400
      Picture         =   "frmMain.frx":299F
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Delete"
      Top             =   3720
      Width           =   1995
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   6360
      TabIndex        =   5
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtSearch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin MSComctlLib.ListView lstView 
         Height          =   2535
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4471
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   1129
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Word"
            Object.Width           =   5997
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmMain.frx":2D61
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Cari :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmMain.frx":2DC3
         TabIndex        =   21
         Top             =   1080
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmMain.frx":2E2D
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   4920
         OleObjectBlob   =   "frmMain.frx":2E9D
         Top             =   3360
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "Sembunyikan (Mode Tray)"
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtWord 
         DataField       =   "Istilah"
         DataSource      =   """ & App.Path & ""\Kamus.mdb"""
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3135
      End
      Begin VB.PictureBox Picture1 
         Height          =   135
         Left            =   240
         ScaleHeight     =   75
         ScaleWidth      =   555
         TabIndex        =   17
         Top             =   3480
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Left            =   5880
         Top             =   3480
      End
      Begin VB.TextBox gintIdItem 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtMeaning 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1320
         Width           =   5895
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BUSEK"
      Height          =   375
      Left            =   8400
      TabIndex        =   16
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Golek Sing Mbuk AYAR'i"
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SIMPEN"
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LEBOK'NO ISTILAH"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderStyle     =   2  'Dash
      DrawMode        =   2  'Blackness
      X1              =   120
      X2              =   10440
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu login 
         Caption         =   "Login"
         Checked         =   -1  'True
         Shortcut        =   ^L
      End
      Begin VB.Menu keluar 
         Caption         =   "Keluar"
         Shortcut        =   ^K
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu ctt 
      Caption         =   "Catat"
      Begin VB.Menu Vnod 
         Caption         =   "Via Notepad"
         Checked         =   -1  'True
         Shortcut        =   ^N
      End
      Begin VB.Menu winword 
         Caption         =   "Via Word"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu gy 
      Caption         =   "Gaya"
      Begin VB.Menu ub 
         Caption         =   "Ubah Pembatas"
         WindowList      =   -1  'True
         Begin VB.Menu dfl 
            Caption         =   "Default"
            Checked         =   -1  'True
         End
         Begin VB.Menu df 
            Caption         =   "Hitam"
         End
         Begin VB.Menu or 
            Caption         =   "Orange"
         End
         Begin VB.Menu hj 
            Caption         =   "Hijau"
         End
         Begin VB.Menu kni 
            Caption         =   "Kuning"
         End
         Begin VB.Menu mrh 
            Caption         =   "Merah"
         End
         Begin VB.Menu br 
            Caption         =   "Biru"
         End
      End
      Begin VB.Menu klt 
         Caption         =   "Kulit"
         Begin VB.Menu mnuListSkin 
            Caption         =   "B-studio"
            Index           =   0
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuListSkin 
            Caption         =   "Galaxy"
            Index           =   1
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuListSkin 
            Caption         =   "Green"
            Index           =   2
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuListSkin 
            Caption         =   "Mac"
            Index           =   3
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuListSkin 
            Caption         =   "Media"
            Index           =   4
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuListSkin 
            Caption         =   "Metallic"
            Index           =   5
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuListSkin 
            Caption         =   "Paper"
            Index           =   6
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuListSkin 
            Caption         =   "Normal"
            Index           =   7
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuListSkin 
            Caption         =   "Top Secret"
            Index           =   8
            Shortcut        =   {F9}
         End
         Begin VB.Menu mnuListSkin 
            Caption         =   "web-II"
            Index           =   9
            Shortcut        =   {F11}
         End
         Begin VB.Menu mnuListSkin 
            Caption         =   "Zhelezo"
            Index           =   10
            Shortcut        =   {F12}
         End
      End
      Begin VB.Menu tran 
         Caption         =   "Transparent"
         Shortcut        =   ^{F1}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private SkinPath As String

Dim oldsize As Long

Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, IpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
cbsize As Long
hWnd As Long
uID As Long
uFlags As Long
uCallbackmessage As Long
hIcon As Long
sZTip As String * 64
End Type

Const NIM_ADD = &H0&
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4
Const WM_MOUSEMOVE = &H200

Dim NI As NOTIFYICONDATA
Dim Result As Long
Private Sub br_Click()
Me.BackColor = vbBlue
End Sub

Private Sub cmdDelete_Click()
a = MsgBox("Yakin...?? Pengen Mbusek", vbOKCancel + vbExclamation)
If a = vbOK Then
    Dim strDelete As String
    
    strDelete = "Delete from EngToMalay Where Id = " & gintIdItem.Text & ""
    gAdoConn.Execute strDelete
    PopData (strTextSearch)
    txtWord.Text = ""
    txtMeaning.Text = ""
    End If
End Sub

Private Sub cmdNew_Click()
    txtWord.Text = ""
    txtMeaning.Text = ""
    cmdSave.Enabled = True
    txtWord.Enabled = True
    txtMeaning.Enabled = True
    lstView.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
    txtSearch.Enabled = False
    cmdNew.Enabled = False
    Command1.Enabled = False
    Label5.BackColor = vbRed
    txtWord.SetFocus
End Sub

Private Sub cmdSave_Click()
Dim strSQL As String
Dim rs As ADODB.Recordset

If txtWord.Text = "" Then
    MsgBox "Sorry Bos, Masukan Kata !!", vbExclamation, "Alert"
    Exit Sub
End If
If txtMeaning.Text = "" Then
    MsgBox "Sorry Bos, Masukan Arti Kata !!", vbExclamation, "Alert"
    Exit Sub
End If

strSQL = "Insert into EngToMalay(Istilah,IstilahDesc)Values('" & SQLSafe(txtWord.Text) & "','" & _
        SQLSafe(txtMeaning.Text) & "')"
gAdoConn.Execute strSQL

PopData (strTextSearch)
txtWord.Text = ""
txtMeaning.Text = ""
Label5.BackColor = vbGreen
    txtWord.SetFocus
    Command1.Caption = "Unlock"
awal
End Sub

Private Sub cmdUpdate_Click()
If Label6.Caption = "Golek Sing Mbuk AYAR'i" Then
txtWord.Enabled = True
txtWord.SetFocus
Label6.Caption = "Ndang Di EDIT !!"
cmdUpdate.BackColor = vbRed
Label6.BackColor = vbRed
Else
Dim strUpdate As String

    strUpdate = "Update EngToMalay Set Istilah = '" & SQLSafe(txtWord) & "'," & _
    "IstilahDesc = '" & SQLSafe(txtMeaning) & "' Where Id = " & gintIdItem & ""
    gAdoConn.Execute strUpdate
    PopData (strTextSearch)
        txtWord.Enabled = False
Label6.Caption = "Golek Sing Mbuk AYAR'i"
cmdUpdate.BackColor = &H8000000F
Label6.BackColor = vbGreen
    End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Unlock" Then
cmdNew.Enabled = True
txtWord.Enabled = False
cmdDelete.Enabled = True
cmdUpdate.Enabled = True
Command1.Caption = "Lock"
Else
awal
Command1.Caption = "Unlock"
End If
End Sub

Private Sub Command2_Click()
sndPlay = sndPlaySound("I Love U.wav", 1)
Me.Hide
End Sub

Private Sub ctt_Click()
ShellExecute Me.hWnd, "open", App.Path & "\notepad.exe" _
                 , vbNullString, vbNullString, 1
End Sub

Private Sub df_Click()
Me.BackColor = System
End Sub

Private Sub dfl_Click()
Me.BackColor = &H80000000
End Sub

Private Sub Form_Load()
     On Error Resume Next
    InitConnection
    PopData (strTextSearch)
    '===========================
Picture1.Visible = False
NI.cbsize = Len(NI)
NI.hWnd = Picture1.hWnd
NI.uID = 0
NI.uID = NI.uID + 1
NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
NI.uCallbackmessage = WM_MOUSEMOVE
Picture1.Picture = Me.Icon
NI.hIcon = Picture1.Picture
NI.sZTip = "Kamus Istilah Bahasa IT" & vbNullChar
Result = Shell_NotifyIconA(NIM_ADD, NI)
 '==========MATI Keluar=================
   RemoveCancelMenuItem Me
End Sub


Private Sub hj_Click()
Me.BackColor = vbGreen
End Sub

Private Sub int_Click()
  
End Sub

Private Sub keluar_Click()
txtWord.Enabled = False
cmdNew.Enabled = False
cmdSave.Enabled = False
cmdUpdate.Enabled = False
cmdDelete.Enabled = False
Command1.Enabled = False
frmMain.login.Visible = True
frmMain.keluar.Visible = False
txtMeaning.Enabled = True
txtWord.Enabled = False
    lstView.Enabled = True
    txtSearch.Enabled = True
    Me.Height = 4200
End Sub



Private Sub kni_Click()
frmMain.BackColor = vbYellow
End Sub

Private Sub login_Click()
TransForm Form1, Form2.txtLevel.Text
    Load Form1
    Form1.Show
End Sub

Private Sub lstView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim intSelItem As Integer
    
    intSelItem = Item
    
    txtWord.Text = lstView.ListItems(intSelItem).ListSubItems(1).Text
    txtMeaning.Text = lstView.ListItems(intSelItem).ListSubItems(2).Text
    
    gintIdItem = lstView.ListItems(intSelItem).ListSubItems(3).Text

End Sub


Private Sub mnuAbout_Click()
    TransForm frmAbout, Form2.txtLevel.Text
    Load frmAbout
    frmAbout.Show
End Sub

Private Sub mnuExit_Click()
a = MsgBox("Mau Keluar..???", vbOKCancel + vbInformation)
If a = vbOK Then
Result = Shell_NotifyIconA(NIM_DELETE, NI)
Animation
Unload Me
End If
End Sub

Private Sub InitConnection()
    Dim conDBString As String

    conDBString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Kamus.mdb"
    
    Set gAdoConn = New ADODB.Connection
        gAdoConn.ConnectionString = conDBString
        gAdoConn.Open

End Sub

Private Sub PopData(strTextSearch As String)
    
    Dim lstX As ListItem
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    
    Dim intCounter As Integer
    If strTextSearch = "" Then
        strSQL = "select * from EngToMalay Order by Istilah ASC"
    Else
        strSQL = "Select * from EngToMalay Istilah " & _
        "where Istilah like '%" & strTextSearch & "%' order by Istilah asc"
        
    End If
    
    
    Set rs = New ADODB.Recordset
        rs.Open strSQL, gAdoConn, 3, 1
    
    lstView.ListItems.Clear
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            intCounter = 1
            While Not .EOF
            Set lstX = lstView.ListItems.Add(, , intCounter)
                lstX.ListSubItems.Add = Trim(!istilah)
                lstX.ListSubItems.Add = Trim(!IstilahDesc)
                lstX.ListSubItems.Add = Trim(!ID)
            intCounter = intCounter + 1
            .MoveNext
            Wend
        End If
    End With
End Sub




Private Sub mrh_Click()
Me.BackColor = vbRed
End Sub

Private Sub or_Click()
Me.BackColor = &H80FF&
End Sub





Private Sub tran_Click()
Form2.Show
End Sub

Private Sub txtSearch_Change()
    PopData (txtSearch.Text)

End Sub


Sub awal()
cmdNew.Enabled = False
cmdUpdate.Enabled = False
cmdDelete.Enabled = False
Command1.Enabled = True
    lstView.Enabled = True
    txtSearch.Enabled = True
    cmdSave.Enabled = False
    txtMeaning.Enabled = True
    txtWord.Enabled = False
End Sub


Public Sub Animation()
Dim i As Long
Dim J As Long
i = Me.ScaleHeight
J = Me.ScaleWidth


While Not i = 0
Me.Height = Me.Height - 25
i = i - 1
Wend


While Not J = 0
Me.Width = Me.Width - 25
J = J - 1
Wend
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim hasil As Long
hasil = (X And &HFF) * &H100
Select Case hasil
Case &H1E00:
frmMain.Show
Case &H4B00:
End Select
End Sub


Private Sub mnuListSkin_Click(Index As Integer)
    Select Case Index
        Case 0:
            SkinPath = App.Path & "\B-Studio.skn"
        Case 1:
            SkinPath = App.Path & "\GALAXY.SKN"
        Case 2:
            SkinPath = App.Path & "\GREEN.SKN"
        Case 3:
            SkinPath = App.Path & "\Mac.skn"
        Case 4:
            SkinPath = App.Path & "\media.skn"
        Case 5:
            SkinPath = App.Path & "\METALLIC.SKN"
        Case 6:
            SkinPath = App.Path & "\Paper.skn"
        Case 7:
            SkinPath = App.Path & "\PLASMOID.SKN"
        Case 8:
            SkinPath = App.Path & "\TopSecret.skn"
        Case 9:
            SkinPath = App.Path & "\Web-II.skn"
        Case 10:
            SkinPath = App.Path & "\WINAQUA.SKN"
        Case 11:
            SkinPath = App.Path & "\zhelezo.skn"
    End Select
    Skin1.LoadSkin SkinPath
    Skin1.ApplySkin Me.hWnd
End Sub



Private Sub Vnod_Click()
  'Menjalankan notepad
    ShellExecute Me.hWnd, "open", "Notepad.exe" _
                 , vbNullString, vbNullString, 1
    'jika Anda menginginkan lokasi
    'file secara detail, Anda bisa merubah
    'notepad.exe menjadi path lengkap yang anda inginkan
    'misal : C:\Windows\Notepad.exe
End Sub

Private Sub winword_Click()
'Menjalankan notepad
    ShellExecute Me.hWnd, "open", "winword.exe" _
                 , vbNullString, vbNullString, 1
    'jika Anda menginginkan lokasi
    'file secara detail, Anda bisa merubah
    'notepad.exe menjadi path lengkap yang anda inginkan
    'misal : C:\Windows\Notepad.exe
    
  End Sub
