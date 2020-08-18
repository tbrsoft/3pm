VERSION 5.00
Begin VB.Form frmCLUF 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CLUF 3PM"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   1530
      TabIndex        =   1
      Top             =   4470
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   180
      Width           =   4485
   End
End
Attribute VB_Name = "frmCLUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = "CLUF - Contrato de licencia de usuario final." + vbCrLf + vbCrLf + _
        "Antes de adquirir y utilizar 3PM deberá estar de acuerdo y aceptar las" + _
        " siguientes condiciones." + vbCrLf + vbCrLf + _
        " TbrSoft de ninguna manera será responsable por el uso dado al sistema" + _
        " por los usuarios finales. La licencia para uso de 3PM será revocada" + _
        " inmediatamente si algún usuario violara las leyes vigentes (respectivas al" + _
        " país en que se utilice) respecto a los derechos de los autores de las" + _
        " composiciones reproducidas por 3PM. En todos los casos se deberá" + _
        " obtener una autorización para la reproducción de todos los ficheros mp3 " + _
        "que se incluyan." + vbCrLf + _
        " El costo estipulado por las instituciones y/o asociaciones de autores no" + _
        " es responsabilidad de tbrSoft si no de los usuarios de 3PM." + vbCrLf + _
        " La adquisición de 3PM no implica derechos de reventa de copias" + _
        " ilegales de este software ni la instalación en más de un equipo " + _
        "(salvo que la licencia adquirida asi lo indique)." + _
        " En caso de disponer de varios equipos" + _
        " deberán solicitar igual cantidad de copias de 3PM." + vbCrLf + _
        " En ningun caso podra someter a 3PM a metodos de decompilación y" + _
        " similares. El codigo fuente de este programa es propiedad de Andres Vazquez" + _
        " Flexes (Argentino, DNI n° 26.453.653) quien es titular unico de los mismos." + vbCrLf + _
        " La instalacion de 3PM y las consecuentes modificaciones" + _
        " que este software provoca en el sistema son responsabilidad exclusiva de" + _
        " quien instala este software y no de tbrSoft. tbrSoft no se hace responsable" + _
        " por las consecuencias de ningun tipo derivadas de la instalacion de 3PM."
        
End Sub
