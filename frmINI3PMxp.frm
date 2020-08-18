VERSION 5.00
Begin VB.Form frmINI3PMxp 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inicio de 3PM para XP"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1800
      TabIndex        =   3
      Top             =   1230
      Width           =   2500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "INICIAR 3PM AL INICIAR WINDOWS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   240
      TabIndex        =   1
      Top             =   510
      Width           =   2500
   End
   Begin VB.CommandButton Command6 
      Caption         =   "NO INICIAR 3PM AL INICIAR WINDOWS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3210
      TabIndex        =   0
      Top             =   510
      Width           =   2500
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Si cuenta con algun programa de seguridad quizas provoque una alerta sobre cambios en el registro"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   6165
   End
End
Attribute VB_Name = "frmINI3PMxp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    Dim TR As New clsTBRREG
    TR.CREARINICIO "3PM", AP + "3pm.exe"
    
    MsgBox "INICIO CREADO"
    
    Set TR = Nothing
End Sub

Private Sub Command6_Click()
    Dim TR As New clsTBRREG
    TR.BORRARINICIO "3pm"
    
    MsgBox "INICIO BORRADO"
    
    Set TR = Nothing
End Sub
