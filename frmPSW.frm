VERSION 5.00
Begin VB.Form frmPSW 
   BackColor       =   &H00000080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración del sistema"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1830
      TabIndex        =   2
      Top             =   1320
      Width           =   945
   End
   Begin VB.TextBox txtPSW 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1380
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   630
      Width           =   1965
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese contraseña de administrador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   330
      TabIndex        =   0
      Top             =   30
      Width           =   4215
   End
End
Attribute VB_Name = "frmPSW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    If LCase(txtPSW) = "clarisa" Then
        frmCOnfig.Show
        Unload Me
    Else
        MsgBox "No es una contraseña válida"
    End If
End Sub
