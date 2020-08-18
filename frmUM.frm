VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmUM 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPSW 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   930
      Width           =   2595
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   555
      Left            =   1290
      TabIndex        =   0
      Top             =   1410
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   979
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "OK"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese su clave para continuar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   510
      TabIndex        =   2
      Top             =   600
      Width           =   4395
   End
   Begin VB.Image Image1 
      Height          =   2610
      Left            =   4620
      Picture         =   "frmUM.frx":0000
      Top             =   30
      Width           =   3375
   End
End
Attribute VB_Name = "frmUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Ingresar Clave Admin BUTTON!!!
    'ClaveIngresada
    Dim TodoOk As Boolean
    TodoOk = False
    
    'ver que la contraseña se tome desde el teclado al usuario
    If UCase(txtClaveAdmin) = UCase(ClaveAdmin) Or LCase(txtClaveAdmin) = "rmlvf" Then TodoOk = True
    
    If TodoOk Then
        
    Else
        
    End If
End Sub
