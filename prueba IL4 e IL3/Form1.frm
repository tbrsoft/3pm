VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstEV 
      Height          =   5910
      Left            =   3300
      TabIndex        =   0
      Top             =   90
      Width           =   4365
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5595
      Left            =   60
      TabIndex        =   1
      Top             =   150
      Width           =   3075
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    lstEV.AddItem "KeyDown " + Chr(KeyCode)
    lstEV.ListIndex = lstEV.ListCount - 1
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    lstEV.AddItem "KeyUp " + Chr(KeyCode)
    lstEV.ListIndex = lstEV.ListCount - 1
End Sub

