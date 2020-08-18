VERSION 5.00
Begin VB.Form F1 
   BackColor       =   &H00000000&
   Caption         =   "Manejo de fallas"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   Icon            =   "F1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4320
      Left            =   150
      Picture         =   "F1.frx":0442
      ScaleHeight     =   4320
      ScaleWidth      =   3960
      TabIndex        =   0
      Top             =   330
      Width           =   3960
   End
   Begin VB.Label LB 
      BackStyle       =   0  'Transparent
      Caption         =   "Recuperar archivos externos desde su copia de seguridad."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   2
      Left            =   4860
      MouseIcon       =   "F1.frx":ABBC
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2550
      Width           =   3495
   End
   Begin VB.Label LB 
      BackStyle       =   0  'Transparent
      Caption         =   "Reparar eliminando archivos externos de 3PM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   1
      Left            =   4650
      MouseIcon       =   "F1.frx":AEC6
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1470
      Width           =   3375
   End
   Begin VB.Label LB 
      BackStyle       =   0  'Transparent
      Caption         =   "Generar un informe de errores para enviar a tbrSoft"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   0
      Left            =   4230
      MouseIcon       =   "F1.frx":B1D0
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   330
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "tbrSoft Internacional 2001 - 2007"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   60
      TabIndex        =   1
      Top             =   4710
      Width           =   8805
   End
End
Attribute VB_Name = "F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LB(0).ForeColor = vbWhite
    LB(1).ForeColor = vbWhite
    LB(2).ForeColor = vbWhite
End Sub

Private Sub LB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Index <> 0 And LB(0).ForeColor <> vbWhite Then LB(0).ForeColor = vbWhite
    If Index <> 1 And LB(1).ForeColor <> vbWhite Then LB(1).ForeColor = vbWhite
    If Index <> 2 And LB(2).ForeColor <> vbWhite Then LB(2).ForeColor = vbWhite
    
    If LB(Index).ForeColor <> vbYellow Then LB(Index).ForeColor = vbYellow
    
End Sub
