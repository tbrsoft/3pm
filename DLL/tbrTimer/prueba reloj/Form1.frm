VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "empezar reloj"
      Height          =   465
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents T As tbrTimer.clsTimer
Attribute T.VB_VarHelpID = -1
Dim H As Long

Private Sub Command1_Click()
    H = 0
    T.Interval = 1000
    T.Enabled = True
End Sub

Private Sub Form_Load()
    Set T = New tbrTimer.clsTimer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    T.Enabled = False
End Sub

Private Sub T_Timer()
    H = H + 1
    Command1.Caption = "empezar reloj " + CStr(H)
End Sub
