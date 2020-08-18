VERSION 5.00
Begin VB.Form frmREV 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FR 
      BackColor       =   &H00000000&
      Caption         =   "Sistema de validación de licencias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1065
      Left            =   990
      TabIndex        =   0
      Top             =   330
      Width           =   6735
      Begin VB.PictureBox picBar 
         BackColor       =   &H00404080&
         Height          =   555
         Left            =   90
         ScaleHeight     =   495
         ScaleWidth      =   6525
         TabIndex        =   2
         Top             =   420
         Width           =   6585
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "07.50"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   5730
         TabIndex        =   1
         Top             =   150
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmREV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    Me.Refresh
    FR.Left = Me.Width / 2 - FR.Width / 2
    FR.Top = Me.Height / 2 - FR.Height / 2
    Me.Refresh
    FR.Refresh
    picBar.Refresh
End Sub
