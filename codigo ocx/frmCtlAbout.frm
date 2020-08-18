VERSION 5.00
Begin VB.Form frmCtlAbout 
   BackColor       =   &H00000040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A cerca de..."
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   1890
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1350
      Left            =   3240
      TabIndex        =   0
      Top             =   90
      Width           =   5160
   End
   Begin VB.Image Image1 
      Height          =   2340
      Left            =   45
      Picture         =   "frmCtlAbout.frx":0000
      Top             =   90
      Width           =   3150
   End
End
Attribute VB_Name = "frmCtlAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label1.Caption = "tbrMP3 Control ActiveX" & vbCrLf & vbCrLf & _
        "Copyright (C) 2002 tbrSoft desafios digitales" & vbCrLf & vbCrLf & _
        "Escrito por Andrés Vázquez" & vbCrLf & vbCrLf & _
        "tbrsoft@hotmail.com"
End Sub

