VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProces 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   4
      Height          =   1545
      Left            =   30
      Top             =   60
      Width           =   5865
   End
   Begin VB.Label lblProces 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscando discos"
      BeginProperty Font 
         Name            =   "HandelGotDLig"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   450
      Width           =   5715
   End
End
Attribute VB_Name = "frmProces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
