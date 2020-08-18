VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmPLIN 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   780
      Width           =   4905
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   6060
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   780
      Width           =   4905
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   6060
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2820
      Width           =   4905
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   450
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2820
      Width           =   4905
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   405
      Left            =   4860
      TabIndex        =   0
      Top             =   2265
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   12632256
      fCapt           =   ">>"
      fEnabled        =   0   'False
      fFontN          =   "Verdana"
      fFontS          =   8
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton2 
      Height          =   405
      Left            =   10560
      TabIndex        =   1
      Top             =   2265
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   12632256
      fCapt           =   ">>"
      fEnabled        =   0   'False
      fFontN          =   "Verdana"
      fFontS          =   9
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton3 
      Height          =   345
      Left            =   5160
      TabIndex        =   4
      Top             =   4500
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   609
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir"
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   8
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton XxBoton2 
      Height          =   405
      Left            =   4890
      TabIndex        =   7
      Top             =   255
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   ">>"
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   8
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton XxBoton5 
      Height          =   405
      Left            =   10530
      TabIndex        =   8
      Top             =   255
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   12632256
      fCapt           =   ">>"
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   8
      fECol           =   5452834
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   2
      Left            =   6090
      Picture         =   "frmPLIN.frx":0000
      Top             =   150
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   3
      Left            =   510
      Picture         =   "frmPLIN.frx":18F6
      Top             =   150
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   1
      Left            =   6090
      Picture         =   "frmPLIN.frx":3113
      Top             =   2160
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   0
      Left            =   480
      Picture         =   "frmPLIN.frx":49C4
      Top             =   2160
      Width           =   4320
   End
End
Attribute VB_Name = "frmPLIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fBoton3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Traducir 'Agregado por el complemento traductor
    
    Text1.Text = TR.Trad("tbrSoft ha desarrollado el formato para karaokes 'MN1'." + vbCrLf + _
        "Este incluye pistas de audio en calidad MP3 en base a samples + instrumentos " + _
        "ejecutados por músicos. Los archivos MN1 incluyen además una o más imágenes " + _
        "y la letra sincronizada. La guía de la letra tiene un indicador de demora " + _
        "en partes instrumentales que no existe en otros sistemas%99%")
        
    Text2.Text = TR.Trad("tbrSoft le permite vender musica a través de su rockola." + vbCrLf + _
        "Agregando este complemento contará con un carrito de compras para adquirir " + _
        "la misma música que esta disponible para escucharse. El usuario podrá " + _
        "elegir la salida por bluetooth o dispositivos USB%99%")
        
    Text3.Text = TR.Trad("3PM cuenta con la posibilidad de utilizar diferentes " + _
        "origenes de musica dentro de la misma rockola. Adquiriendo este plugin " + _
        "podrá además incluir sitios FTP con multimedia de modo que no necesitará " + _
        "viajar para actualizar la música en sus equipos%99%")
        
    Text4.Text = TR.Trad("Este complemento le permite contar con un panel que " + _
        "le informa el estado de su rockola en todo momento. Podrá hacer las " + _
        "liquidaciones de dinero, podrá saber que es lo que se escucha modificar " + _
        "la configuración del sistema, cambiar la musica que se muestra, etc. " + _
        "En resumen podrá administrar la rockola sin estar presente%99%")
End Sub

Private Sub XxBoton2_Click()
    frmHabKar.Show 1
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    fBoton3.Caption = TR.Trad("SALIR%99%")
End Sub

Private Sub XxBoton5_Click()
    frmHabCart.Show 1
End Sub
