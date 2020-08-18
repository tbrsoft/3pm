VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmINI3PMxp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inicio de 3PM para XP"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton Command1 
      Height          =   435
      Left            =   1800
      TabIndex        =   3
      Top             =   1410
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton command6 
      Height          =   555
      Left            =   3210
      TabIndex        =   2
      Top             =   630
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   979
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "no iniciar 3pm al iniciar windows"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command5 
      Height          =   555
      Left            =   210
      TabIndex        =   1
      Top             =   630
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   979
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "iniciar 3pm al iniciar windows"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Si cuenta con algún programa de seguridad quizás provoque una alerta sobre cambios en el registro."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Left            =   90
      TabIndex        =   0
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
    Dim TR2 As New clsTBRREG
    TR2.CREARINICIO dcr("1Vx0YVGhEoIisHPLAZMHXw=="), AP + "3pm.exe"
    
    MsgBox TR.Trad("INICIO CREADO%98%Se refiere a que 3PM iniciara junto con windows%99%")
    
    Set TR2 = Nothing
End Sub

Private Sub Command6_Click()
    Dim TR2 As New clsTBRREG
    TR2.BORRARINICIO dcr("q44KmdDBQ+IB8dTOX8F+VA==")
    
    MsgBox TR.Trad("INICIO BORRADO%98%Se refiere a que 3PM no iniciara juto con windows%99%")
    
    Set TR2 = Nothing
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Form_Load()
    Pintar_fBoton Me
    Traducir
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Command1.Caption = TR.Trad("SALIR%99%")
    Command5.Caption = TR.Trad("INICIAR 3PM AL INICIAR WINDOWS%99%")
    Command6.Caption = TR.Trad("NO INICIAR 3PM AL INICIAR WINDOWS%99%")
    Label3.Caption = TR.Trad("Si cuenta con algún programa de seguridad " + _
        "quizás provoque una alerta sobre cambios en el registro.%99%")
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub

'3pm 'hay mas de uno para confundir
'dcr("1Vx0YVGhEoIisHPLAZMHXw==")
'dcr("q44KmdDBQ+IB8dTOX8F+VA==")

'dcr("MCuVh38359iRH+GBaAkXedz8Pl38peUqZHKs0a0SpMe+QLrW9mKdnA==")
'dcr("yTSbeYe2oWp2ydIUpGyes+DYNN6qU8l9pMMGAAqH+wBg8bBgTQ+/hw==")
'dcr("VZPmSDtgWIj2UthiVZfN1LsFHe7IZv/K/ue9/JPXBYNJosAztaasKg==")
'dcr("OqgcJfckN8975IVShi0xrqPphoO7CJfy1bRk3zQnHno=")

