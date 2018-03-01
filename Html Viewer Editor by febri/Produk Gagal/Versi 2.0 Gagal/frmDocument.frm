VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmDocument 
   Caption         =   "Jendela Html Viewer editor"
   ClientHeight    =   5340
   ClientLeft      =   3780
   ClientTop       =   1908
   ClientWidth     =   6000
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   6000
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2772
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5532
      ExtentX         =   9758
      ExtentY         =   4890
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
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   2004
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   3535
      _Version        =   393217
      BackColor       =   12640511
      ScrollBars      =   3
      TextRTF         =   $"frmDocument.frx":1CFA
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub rtfText_SelChange()
    frmMain.tbToolBar.Buttons("Bold").Value = IIf(rtfText.SelBold, tbrPressed, tbrUnpressed)
        frmMain.tbToolBar.Buttons("Italic").Value = IIf(rtfText.SelItalic, tbrPressed, tbrUnpressed)
        frmMain.tbToolBar.Buttons("Underline").Value = IIf(rtfText.SelUnderline, tbrPressed, tbrUnpressed)
        frmMain.tbToolBar.Buttons("Align Left").Value = IIf(rtfText.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
        frmMain.tbToolBar.Buttons("Center").Value = IIf(rtfText.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
        frmMain.tbToolBar.Buttons("Align Right").Value = IIf(rtfText.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
End Sub
Private Sub Form_Resize()
On Error Resume Next
WebBrowser1.Top = 40
WebBrowser1.Left = 40
WebBrowser1.Width = Me.Width - 300
WebBrowser1.Height = Me.Height / 2 - 80

rtfText.Top = Me.Height / 2 + 100
rtfText.Left = 40
rtfText.Width = Me.Width - 300
rtfText.Height = Me.Height / 2 - 950
End Sub



Private Sub rtfText_Change()
On Error Resume Next
DoEvents
Open "C:\temp.html" For Output As #1: Print #1, rtfText.Text: Close #1
DoEvents
WebBrowser1.Navigate "C:\temp.html"
End Sub

