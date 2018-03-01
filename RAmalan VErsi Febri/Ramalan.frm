VERSION 5.00
Begin VB.Form Frm_Ramalan 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ramalan Cinta (By-Febrian Dwi Putra)"
   ClientHeight    =   5532
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Ramalan.frx":0000
   Palette         =   "Ramalan.frx":0492
   Picture         =   "Ramalan.frx":3E4D4
   ScaleHeight     =   5532
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   3720
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2040
      TabIndex        =   6
      Top             =   3120
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RAMAL"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5640
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BARU"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   6480
      TabIndex        =   7
      Top             =   5160
      Width           =   612
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ramalan Cinta v-1.0"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SARAN :"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   2880
      TabIndex        =   11
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Kamu"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pasangan Kamu"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   6855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   75
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Persentase Cinta Kalian :"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Program Powered : By Febrian Dwi Putra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   492
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   5892
   End
   Begin VB.Image Image2 
      Height          =   492
      Left            =   0
      Picture         =   "Ramalan.frx":49B40
      Top             =   5040
      Width           =   12216
   End
   Begin VB.Image Image1 
      Height          =   492
      Left            =   0
      Picture         =   "Ramalan.frx":727AA
      Top             =   0
      Width           =   12216
   End
End
Attribute VB_Name = "Frm_Ramalan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yayat As Object
Dim arrmark(100) As String
Dim Control As Integer

Private Sub Command1_Click()

Text1.Text = Empty
Text2.Text = Empty
Label4.Caption = Empty
Label6.Caption = Empty
Label7.Caption = Empty
Text1.SetFocus

End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Then
   MsgBox ("Tolong masukkan data yang lengkap")
Else
    Call JILL
End If

End Sub
Sub JILL()

KAMU = Len(Text1.Text)
PASANGANMU = Len(Text2.Text)

HASILNYA = Val(KAMU) + Val(PASANGANMU) * 2
HASILNYA2 = HASILNYA Mod 9

Select Case HASILNYA2
Case "1"
X = "0%"
U = "Mungkin belum saatnya untuk bersama :-D"
Y = "Carilah toko obat nyamuk terdekat untuk bunuh diri"

Case "2"
X = "5%"
U = "Ada sedikit harapan tapi mendingan udahan aja dech daripada nyesel ntar :-D"
Y = "Lebih cepat lebih baik untuk cari pengganti sebelum keduluan si doi"

Case "3"
X = "10%"
U = "Si doi ingin sesuatu yang mewah"
Y = "Rajin-rajinlah menabung untuk membeli mobil mewah"

Case "4"
X = "83%"
U = "Pasangan anda ingin lebih sering dekat dengan anda"
Y = "Gunakan lem G /altecko agar kalian berdua bisa nempel terus"

Case "5"
X = "50%"
U = "Anda harus berhati-hati dengan si doi karna dia punya rencana busuk"
Y = "Segera lapor ke kepala desa setempat agar tidak kejadian"

Case "6"
X = "75%"
U = "Lebih mesra lagi dalam menjalin hubungan "
Y = "Cukup satu saja jangan banyak-banyak ya "

Case "7"
X = "90%"
U = "Wah ini baru pasangan yang sudah hampir cocok :-D"
Y = "Semoga malam ini bisa jadi malam terindah ya :-D"

Case "8"
X = "100%"
U = "Selamat menempuh hidup baru buat kalian berdua"
Y = "Ingat undangannya ya :-D Ditunggu :-D"

Case "9"
X = "200%"
U = "Cinta kalian telah melebihi batas yang ada"
Y = "periksakan diri kalian berdua ke dokter psikolog terdekat"

Case "0"
X = "99,99%"
U = "Peluang yang sangat besar untuk mencetak goal tapi sayang masih kurang beruntung 0,01%"
Y = "Tetap semangat !!!"

End Select

Label4.Caption = X
Label6.Caption = U
Label7.Caption = Y

End Sub
Private Sub Command3_Click()
If MsgBox("Tutup Aplikasi . . . . . ?", vbQuestion + vbYesNo, "komfimasi") = vbYes Then
End
End If
End Sub

