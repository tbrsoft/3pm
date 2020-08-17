VERSION 5.00
Begin VB.Form frmTemas 
   BackColor       =   &H00808000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Temas encontrados"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   5160
      TabIndex        =   1
      Top             =   3720
      Width           =   2205
   End
   Begin VB.ListBox lstTemas 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3840
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   330
      Width           =   4455
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Tapa de CD encontrada"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4620
      TabIndex        =   3
      Top             =   60
      Width           =   3555
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de temas encontrados"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4215
   End
   Begin VB.Image TapaCD 
      Height          =   3240
      Left            =   4650
      Picture         =   "frmTemas.frx":0000
      Stretch         =   -1  'True
      Top             =   330
      Width           =   3600
   End
End
Attribute VB_Name = "frmTemas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    DuplicarLST lstTemas, frmCOnfig.lstTemas
    frmCOnfig.TapaCD.Picture = TapaCD.Picture
    Unload Me
End Sub
