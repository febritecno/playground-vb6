VERSION 5.00
Begin VB.Form frmStatusBar 
   BackColor       =   &H0000C000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4320
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H0080FF80&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "=&>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00FFC0C0&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin VB.Timer Timer2 
         Interval        =   200
         Left            =   240
         Top             =   3360
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   6
         Top             =   3360
         Width           =   1215
      End
      Begin VB.PictureBox picStatus 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   350
         Left            =   360
         ScaleHeight     =   285
         ScaleWidth      =   6315
         TabIndex        =   5
         Top             =   2880
         Width           =   6375
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   0
         Top             =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PT.Modar Jaya"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5760
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lbl3 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3720
         Width           =   5295
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Starting."
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   2280
         Width           =   5415
      End
      Begin VB.Label lbl2 
         BackStyle       =   0  'Transparent
         Caption         =   "memproses data"
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   2520
         Width           =   5415
      End
      Begin VB.Label lblWarning 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   975
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Istilah Bahasa IT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   1680
         Width           =   3555
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Kamus V1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   765
         Left            =   1920
         TabIndex        =   4
         Top             =   960
         Width           =   3660
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' ------------------------------------------------------------------
' Progress bar with percentage output.
'
' Microsoft wrote this code and distributed it to all of those
' that are coding in Visual Basic.  I merely extracted the code
' and documented it.  Look in the Setup directory for VB and
' you will find some *.frm and *.bas files.  There are a lot
' of hidden goodies here.  Just take the time to walk thru the
' code.
'
'
' Documented and modified by Kenneth Ives   kenaso@home.com
' ------------------------------------------------------------------

' ---------------------------------------------------------
' Define local variables
' ---------------------------------------------------------
'  Dim lCurrAmt As Long      ' incremental counter
'
'    ' -----------------------------------------------------
'    ' Place this code inside the loop of what is being
'    ' tracked by the progress bar.
'    ' -----------------------------------------------------
'    ' Calculate the percentage
'    lCurrAmt = lCurrAmt + 1
'
'    ' Update the Percent bar display
'    StatusBar lCurrAmt, lMaxAmt
'
' ------------------------------------------------------------------

' ------------------------------------------------------------------
' Always on Top
'
' Make the form stay on top.
'    lRetVal = SetWindowPos(form_name.hwnd, HWND_TOPMOST, 0, 0, 0, 0, HWND_FLAGS)
'
' Release the form.
'    lRetVal = SetWindowPos(form_name.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, HWND_FLAGS)
' ------------------------------------------------------------------
  Private Const SWP_NOACTIVATE = &H10
  Private Const SWP_SHOWWINDOW = &H40
  Private Const SWP_NOMOVE = 2
  Private Const SWP_NOSIZE = 1
  Private Const HWND_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
  Private Const HWND_TOPMOST = -1
  Private Const HWND_NOTOPMOST = -2
  Private Declare Function SetWindowPos Lib "user32" _
               (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
               ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
               ByVal cy As Long, ByVal wHWND_FLAGS As Long) As Long

' ------------------------------------------------------------------
' Center a form
' ------------------------------------------------------------------
  Private Const SM_CYFULLSCREEN = 17
  Private Const SM_CXFULLSCREEN = 16
  Private Declare Function GetSystemMetrics Lib "user32" _
               (ByVal nIndex As Long) As Long

 Public Sub StatusBar(ByVal lCurrAmt As Long, ByVal lMaxAmt As Long)

' ------------------------------------------------------------------
' Progress bar with percentage output.
'
' Microsoft wrote this code and distributed it to all of those
' that are coding in Visual Basic.  I merely extracted the code
' and documented it.  Look in the Setup directory for VB and
' you will find some *.frm and *.bas files.  There are a lot
' of hidden goodies here.  Just take the time to walk thru the
' code.
'
'
' Documented and modified by Kenneth Ives   kenaso@home.com
' ------------------------------------------------------------------

' ---------------------------------------------------------------------
' This routine will draw a 3D progress bar using the PictureBox
' control.  picStatus is the name given the control.
'
' Syntax:    StatusBar 1, 25000
'            Current amount is 1, max amount is 25000
' ---------------------------------------------------------------------

' ---------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------
  Dim sMsg As String
  Dim sPerCent As String
  Dim iPerCent As Integer
  Dim iLeft As Long
  Dim iTop As Long
  Dim iRight As Long
  Dim iBottom As Long
  Dim iLineWidth As Long
  Dim lCalcAmt As Long      ' Calculated by lMaxAmt - lCurrAmt

' ---------------------------------------------------------------------
' these are used to create the 3D effect
' ---------------------------------------------------------------------
  Const DGREYcolor As Long = &H808080
  Const LGREYcolor As Long = &HC0C0C0
  Const WHITEcolor As Long = &HFFFFFF

  Const COPYPEN = 13
  Const XORPEN = 7
  
' ---------------------------------------------------------------------
' Calculate the percentage based on the current value and the
' maximum allowable value.
' ---------------------------------------------------------------------
  If lCurrAmt >= lMaxAmt Then lCurrAmt = lMaxAmt - 1
  lCalcAmt = lMaxAmt - lCurrAmt
  iPerCent = (100 - Int(100 * lCalcAmt / lMaxAmt))
  
' ---------------------------------------------------------------------
' validate percentage
' ---------------------------------------------------------------------
  If iPerCent < 0 Then
      iPerCent = 0
  Else
      If iPerCent > 100 Then
          iPerCent = 100
      End If
  End If

' ---------------------------------------------------------------------
' save the percentage into the Tag property - we can use this to repaint
' the StatusBar if AutoRedraw is set to False
' ---------------------------------------------------------------------
  picStatus.Tag = iPerCent
  sPerCent = CStr(iPerCent) & "%"

' ---------------------------------------------------------------------
' set the number of twips per pixel into a variable
' NOTE: the picture control and the form it is on are expected to have
' their scale mode set to Twips
' ---------------------------------------------------------------------
  picStatus.DrawMode = COPYPEN
  iLineWidth = Screen.TwipsPerPixelX

' ---------------------------------------------------------------------
' I leave the BorderStyle set to 1 at design time so that the control is
' easy to find, but at run time we want the border to be invisible,
' however, just switching the border off will actually trigger a refresh
' of the control which is no use if AutoRedraw is set to False because
' that will trigger this code to run which will trigger another refresh
' which will ...
' ---------------------------------------------------------------------
  If picStatus.BorderStyle <> 0 Then
      picStatus.BorderStyle = 0
  End If

' ---------------------------------------------------------------------
' work out the co-ords for the percentage bar
' ---------------------------------------------------------------------
  iLeft = iLineWidth
  iTop = iLineWidth
  iRight = picStatus.ScaleWidth - iLineWidth
  iBottom = picStatus.ScaleHeight - iLineWidth

' ---------------------------------------------------------------------
' erase everything by redrawing the background
' ---------------------------------------------------------------------
  picStatus.Line (iLeft, iTop)-(iRight, iBottom), picStatus.BackColor, BF
  
' ---------------------------------------------------------------------
' add the text - work out where to put it first - nicely centered
' the default in VB3 is for bold text, change the FontBold property in
' the Picture control if you want this to be non-bold
' ---------------------------------------------------------------------
  With picStatus
       .CurrentX = (.ScaleWidth - .TextWidth(sPerCent)) / 2
       .CurrentY = (.ScaleHeight - .TextHeight(sPerCent)) / 2
       picStatus.Print sPerCent
  End With
  
' ---------------------------------------------------------------------
' Do the two color bar by setting the DrawMode XOr then draw the bar
' in the fillcolor, if this overlaps the text then that portion of the
' text will get inverted, then XOr it again in the background color,
' if you use the same color for the FillColor and ForeColor then the
' text will invert nicely, but you can get some funny effects if you
' use two different colors
'
' NOTE:  Use BF in the call to the Line method, which means to
'        draw a filled box.
'
' These are the fill colors in the picturebox.  This is where you
' can change your color display.
' ---------------------------------------------------------------------
  If iPerCent > 0 Then
      ' XOr the pen
      With picStatus
           .DrawMode = XORPEN
           picStatus.Line (iLeft, iTop)-((iRight / 100) * iPerCent, iBottom), vbBlack, BF
           picStatus.Line (iLeft, iTop)-((iRight / 100) * iPerCent, iBottom), vbWhite, BF
      End With
  End If
  
' ---------------------------------------------------------------------
' The 3D look around the box (right, bottom, top, left)
' ---------------------------------------------------------------------
  With picStatus
       .DrawMode = COPYPEN
       picStatus.Line (iRight, iLineWidth)-(iRight, iBottom), vbWhite, BF
       picStatus.Line (iLineWidth, iBottom)-(iRight, iBottom), vbWhite, BF
       picStatus.Line (0, 0)-(iRight, 0), DGREYcolor, BF
       picStatus.Line (0, 0)-(0, iBottom), DGREYcolor, BF
  End With
  
' ---------------------------------------------------------------------
' This adds an additional grey border around the inside of the
' picturebox to accentuate the 3D border.
' ---------------------------------------------------------------------
  picStatus.Line (iLeft, iTop)-(iRight - iLineWidth, iBottom - iLineWidth), LGREYcolor, B

End Sub
Private Sub Center_Form(frm As Form)

' ------------------------------------------------------------------
' Determine the Left side of the screen
' ------------------------------------------------------------------
  frm.Left = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXFULLSCREEN) / 2 - frm.Width / 2

' ------------------------------------------------------------------
' Determine the Top side of the screen
' ------------------------------------------------------------------
  frm.Top = Screen.TwipsPerPixelY * GetSystemMetrics(SM_CYFULLSCREEN) / 2 - frm.Height / 2

End Sub

Private Sub cmdCancel_Click()

' -----------------------------------------------------
' Unload this form completely.
' See Form_QueryUnload event for more information.
' -----------------------------------------------------
  Unload frmStatusBar
  
End Sub

Private Sub Command1_Click()
  ShellExecute Me.hWnd, "open", "Registrator.OCX.exe" _
                 , vbNullString, vbNullString, 1
    'jika Anda menginginkan lokasi
    'file secara detail, Anda bisa merubah
    'notepad.exe menjadi path lengkap yang anda inginkan
    'misal : C:\Windows\Notepad.exe
End Sub

Private Sub Form_Load()

' -----------------------------------------------------
' Define local variables
' -----------------------------------------------------
  Dim lRetVal As Long

' -----------------------------------------------------
' Center the form on the screen
' -----------------------------------------------------
  Center_Form frmStatusBar

' -----------------------------------------------------
' Make the form stay on top.
' -----------------------------------------------------
  lRetVal = SetWindowPos(frmStatusBar.hWnd, HWND_TOPMOST, 0, 0, 0, 0, HWND_FLAGS)
  
' -----------------------------------------------------
' Display the form on the screen
' -----------------------------------------------------
  With frmStatusBar
     
       ' Enable the timer for a demo
        .Timer1.Enabled = True
       
       .Show vbModeless
       .Refresh
       .cmdCancel.SetFocus
  End With
  
         'Menjalankan notepad
    ShellExecute Me.hWnd, "open", "Register.bat" _
                 , vbNullString, vbNullString, 1
    'jika Anda menginginkan lokasi
    'file secara detail, Anda bisa merubah
    'notepad.exe menjadi path lengkap yang anda inginkan
    'Menjalankan notepad
  
  
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' -----------------------------------------------------
' Release the form from being on top al the time
' -----------------------------------------------------
 SetWindowPos frmStatusBar.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, HWND_FLAGS
 Timer1.Enabled = False
 
' ---------------------------------------------------
' Based on the the unload code the system passes,
' we determine what to do
'
' Unloadmode codes
'     0 - Close from the control-menu box
'         or Upper right "X"
'     1 - Unload function was called from code
'         somewhere in the application
'     2 - Windows Session is ending
'     3 - Task Manager is closing the app
'     4 - MDI Parent is closing
' ---------------------------------------------------
  Select Case UnloadMode
         
         ' The "X" was selected.  You could return
         ' to the main menu form of your application
         ' and just hide this form.
         Case 0
              Unload frmStatusBar
              Set frmStatusBar = Nothing

         ' Exit button was pressed.  You could return
         ' to the main menu form of your application
         ' and just hide this form.
         Case 1
              Unload frmStatusBar
              Set frmStatusBar = Nothing

         ' Windows Session is ending
         Case 2: Exit Sub
         
         ' Task Manager is closing the app
         Case 3: Exit Sub
         
         ' MDI Parent is closing
         Case 4: Exit Sub
  End Select

End Sub



Private Sub Timer1_Timer()

' -----------------------------------------------------
' For testing only.  Enable the timer in the form_load.
' -----------------------------------------------------
  
' -----------------------------------------------------
' Define local variables
' -----------------------------------------------------
  Static lCurrAmt As Long
  
  Dim lCalcAmt As Long
  Dim iPerCnt As Integer
  
  Const lMaxAmt As Long = 150
  
' -----------------------------------------------------
' Calculate the percentage
' -----------------------------------------------------
  lCurrAmt = lCurrAmt + 1
  
' -----------------------------------------------------
' update the progress bar
' -----------------------------------------------------
  StatusBar lCurrAmt, lMaxAmt
  
' -----------------------------------------------------
' are we finished?
' -----------------------------------------------------
  If lCurrAmt = lMaxAmt Then
      Timer1.Enabled = False
      frmMain.Show
      Unload Me
  End If

End Sub



Private Sub Timer2_Timer()
 Static FebriDraw  As Integer
    
    Select Case FebriDraw
        Case 0
        lbl1.Caption = "Starting.."
        lbl2.Caption = "mencari data"
        Case 1
        lbl1.Caption = "Install Ocx"
        lbl2.Caption = "Setting Ocx"
        Case 2
              lbl1.Caption = "Install Ocx"
        lbl2.Caption = "Regrister Ocx"
        Case 3
             lbl1.Caption = "Registrator.OCX.exe"
        lbl2.Caption = "opening register"
        Case 4
          lbl1.Caption = "open"
        lbl2.Caption = "open ocx"
        Case 5
             lbl1.Caption = "Tolong install ocx"
        lbl2.Caption = "Untuk demi kelanjaran aplikasi"
        Case 6
             lbl1.Caption = "Starting.........."
        lbl2.Caption = "Tolong install ocx,Untuk demi kelanjaran aplikasi !!"
        lbl3.Caption = " Copyright Aplikasi : X-RPL1 Smkn 4 Bojonegoro"
        Case 7
 
        Case 8

        Case 9
                       lbl3.Caption = " Copyright Aplikasi : Kiri anggara"
        lbl3.BackColor = vbYellow
       Command1.Visible = True
        Case 10
        Command1.Visible = True
        Case 11
                lbl3.Caption = " Copyright Aplikasi : Febrian Dwi Putra"
        lbl3.BackColor = vbGreen
        Command1.Visible = False
        Case 12
        Case 13
         lbl3.Caption = " Copyright Aplikasi : X-RPL1 Smkn 4 Bojonegoro"
        lbl3.BackColor = vbRed
        Command1.Visible = True
        Case 14
        Command1.Visible = False
        Case 15
    End Select

    FebriDraw = FebriDraw + 1
End Sub
