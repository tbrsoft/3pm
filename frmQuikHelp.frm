VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmQuikHelp 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ayuda rapida de 3PM"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   6390
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   979
      fFColor         =   0
      fBColor         =   16761024
      fCapt           =   "Abrir Manual de 3PM"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16777215
   End
   Begin VB.TextBox txTeclas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5475
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmQuikHelp.frx":0000
      Top             =   840
      Width           =   7965
   End
   Begin tbrFaroButton.fBoton fBoton2 
      Height          =   555
      Left            =   1830
      TabIndex        =   2
      Top             =   6390
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   979
      fFColor         =   0
      fBColor         =   16761024
      fCapt           =   "Agregar música a 3PM"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16777215
   End
   Begin tbrFaroButton.fBoton fBoton3 
      Height          =   555
      Left            =   6420
      TabIndex        =   4
      Top             =   6360
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   979
      fFColor         =   0
      fBColor         =   16744576
      fCapt           =   "SALIR"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16777215
   End
   Begin tbrFaroButton.fBoton fBoton4 
      Cancel          =   -1  'True
      Height          =   555
      Left            =   4770
      TabIndex        =   5
      Top             =   6360
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   979
      fFColor         =   0
      fBColor         =   16744576
      fCapt           =   "Cerrar 3PM"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16777215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmQuikHelp.frx":0006
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   150
      TabIndex        =   3
      Top             =   90
      Width           =   7935
   End
End
Attribute VB_Name = "frmQuikHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fBoton1_Click()
    AbrirArchivo AP + "manual.doc", Me
End Sub

Private Sub fBoton2_Click()
    frmAddMusic.Show 1
End Sub

Private Sub fBoton3_Click()
    Unload Me
End Sub

Private Sub fBoton4_Click()
    Unload Me
    YaCerrar3PM
End Sub

Private Sub Form_Load()
    txTeclas.Text = "Teclas más usadas:" + vbCrLf + vbCrLf + _
        "  Z = ir a la izquierda o arriba dentro de un disco" + vbCrLf + _
        "  X = ir a la derecha o abajo dentro de un disco" + vbCrLf + _
        "  Q = simular inserción de moneda (sumar crédito)" + vbCrLf + _
        "  W = Cerrar 3PM" + vbCrLf + _
        "  ENTER = seleccionar disco o canción" + vbCrLf + _
        "  ESCAPE = salir de un disco" + vbCrLf + _
        "Por más teclas consulte el manual." + vbCrLf + vbCrLf + _
        "NO DUDE en enviarnos su consulta: " + vbCrLf + _
        "  * Estamos el línea de 9 a 19 hs (hora argentina) por MSN en soporte_tbrsoft@hotmail.com o ventas@tbrsoft.com " + vbCrLf + _
        "  * Por email a soporte_tecnico@tbrsoft.com o info@tbrsoft.com" + vbCrLf + _
        "  * Visite nuestro sitio web www.tbrsoft.com por información sobre representantes de tbrSoft en su país." + vbCrLf + _
        "  * Por telefono 03543-401066 o desde fuera de Argentina +54-3543-401066" + vbCrLf + vbCrLf + _
        "3PM además puede funcionar además de rockola como :" + vbCrLf + _
        "  KARAOKE" + vbCrLf + _
        "  EXPENDEDOR de múscia (VENDE POR BLUETOOTH, USB, CD, etc)" + vbCrLf + _
        "  Musica funcional" + vbCrLf + _
        "  Musica gratuita en fiestas" + vbCrLf + _
        "Y seguimos ampliando funciones sugún pedidos de nuestros cientos de clientes" + vbCrLf + vbCrLf + _
        "tbrSoft Internacional S.A." + vbCrLf + _
        "Desafios digitales 2001-2008"
        
    
End Sub
