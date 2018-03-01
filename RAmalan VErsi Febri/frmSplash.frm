VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4260
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000B&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4260
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   360
      Top             =   2160
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   2760
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "v-1.0"
      BeginProperty Font 
         Name            =   "OCR-A BT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5040
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ramalan Cinta"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tunggu Dulu"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2280
      Width           =   3495
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim efek As Integer

Private Declare Function CreateEllipticRgn _
Lib "gdi32" (ByVal X1 As Long, _
ByVal Y1 As Long, ByVal X2 As Long, _
ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn _
Lib "User32" (ByVal Hwnd As Long, _
ByVal hRgn As Long, _
ByVal bRedraw As Long) As Long
Sub Splash(Fm As Form)
Fm.Show
Dim a As Integer
Dim b As Integer
Dim C As Integer
Dim d As Integer
Dim e As Integer
Dim f As Integer
Dim w As Integer
Dim X As Integer
Dim Y As Integer
Dim z As Integer
Dim current As Double
Call Fm.Move(0, 0)
w = Fm.Height: X = Fm.Width
Y = Fm.Top: z = Fm.Left
a = 0: b = 0: C = w
d = X: e = Y: f = z
Do While a < Fm.Height / 15 Or _
b < Fm.Width / 15
a = a + 30
b = b + 30
e = e + 70
f = f + 70
If a > Fm.Height / 15 Then a = a - 24
If b > Fm.Width / 15 Then b = b - 24
Call Fm.Move(f, e, d, C)
current = Timer
Do While Timer - current < 0.01
DoEvents
Loop
Call SetWindowRgn(Fm.Hwnd, _
CreateEllipticRgn(0, 0, b, a), True)
Loop
32
current = Timer
Do While Timer - current < 1
DoEvents
Loop
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
efek = efek + 5
ProgressBar1.Value = ProgressBar1.Value + 400 / 400
If efek > 500 Then
    Timer1.Enabled = False
    Screen.MousePointer = vbNormal
    Me.WindowState = 0
    Do
    Me.Left = Me.Left + 20
    Me.Move Me.Left, Me.Top
    DoEvents
    Loop Until Me.Left > Screen.Width
    Load Frm_Ramalan
    Frm_Ramalan.Show
    Unload Me
End If
End Sub

Private Sub Form_Load()
Call Splash(Me)
End Sub

