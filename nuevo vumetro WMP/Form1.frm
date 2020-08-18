VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "STOP"
      Height          =   345
      Left            =   3300
      TabIndex        =   2
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      Height          =   345
      Left            =   2520
      TabIndex        =   1
      Top             =   5460
      Width           =   1935
   End
   Begin Proyecto1.VUMeter VU1 
      Height          =   5085
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   8969
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    VU1.DoStart
End Sub

Private Sub Command2_Click()
    VU1.DoStop
    End
End Sub
