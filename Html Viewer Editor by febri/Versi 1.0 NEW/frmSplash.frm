VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4236
   ClientLeft      =   252
   ClientTop       =   1416
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4236
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   800
         Left            =   720
         Top             =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Loading"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   2760
         TabIndex        =   1
         Top             =   2160
         Width           =   1716
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   372
         Left            =   5520
         Top             =   3000
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H0080C0FF&
         FillStyle       =   0  'Solid
         Height          =   372
         Left            =   4560
         Top             =   3000
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   372
         Left            =   3600
         Top             =   3000
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   372
         Left            =   2640
         Top             =   3000
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   372
         Left            =   1680
         Top             =   3000
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   372
         Left            =   720
         Top             =   3000
         Visible         =   0   'False
         Width           =   972
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
 Static FebriDraw  As Integer
    
    Select Case FebriDraw
        Case 0
           Shape1.Visible = True
        Case 1
        
        Case 2
            Shape2.Visible = True
        Case 3
            Shape3.Visible = True
        Case 4
        
        Case 5
            Shape4.Visible = True
        Case 6
            Shape5.Visible = True
        Case 7
        
        Case 8
            Shape6.Visible = True
        Case 9
            
        Case 10
            Form1.Show
            Unload Me
    End Select
    
    FebriDraw = FebriDraw + 1
End Sub
