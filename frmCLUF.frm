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
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1530
      TabIndex        =   1
      Top             =   4470
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case TeclaNewFicha
            'si ya hay 9 cargados se traga las fichas
            If CREDITOS <= MaximoFichas Then
                OnOffCAPS vbKeyScrollLock, True
                CREDITOS = CREDITOS + TemasPorCredito
                SumarContadorCreditos TemasPorCredito
                'grabar cant de creditos
                EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
                If CREDITOS >= 10 Then
                    frmINDEX.lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
                Else
                    frmINDEX.lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
                End If
            Else
                'apagar el fichero electronico
                OnOffCAPS vbKeyScrollLock, False
            End If
    End Select
End Sub

Private Sub Form_Load()
    Text1.Text = "CLUF - Contrato de licencia de usuario final." + vbCrLf + vbCrLf + _
    "Antes de adquirir y utilizar 3PM deberá estar de acuerdo y aceptar las siguientes condiciones." + vbCrLf + vbCrLf + _
    " TbrSoft de ninguna manera será responsable por el uso dado al sistema por los usuarios finales. La " + _
    "licencia para uso de 3PM será revocada inmediatamente si algún usuario violara las leyes vigentes " + _
    "(respectivas al país en que se utilice) respecto a los derechos de los autores de las composiciones " + _
    "reproducidas por 3PM. En todos los casos se deberá obtener una autorización para la reproducción de " + _
    "todos los ficheros mp3 que se incluyan." + vbCrLf + "El costo estipulado por las instituciones y/o " + _
    "asociaciones de autores no es responsabilidad de tbrSoft si no de los usuarios de 3PM." + vbCrLf + _
    "La adquisición de 3PM no implica derechos de reventa de copias ilegales de este software ni la " + _
    "instalación en más de un equipo (salvo que la licencia adquirida asi lo indique). En caso de disponer " + _
    "de varios equipos deberán solicitar igual cantidad de copias de 3PM." + vbCrLf + _
    "En ningun caso podra someter a 3PM a metodos de decompilación y similares. El codigo fuente de este " + _
    "programa es propiedad de Andres Vazquez Flexes (Argentino, DNI n° 26.453.653) quien es titular unico de " + _
    "los mismos." + vbCrLf + " La instalacion de 3PM y las consecuentes modificaciones" + _
    " que este software provoca en el sistema son responsabilidad exclusiva de" + _
    " quien instala este software y no de tbrSoft. tbrSoft no se hace responsable" + _
    " por las consecuencias de ningun tipo derivadas de la instalacion de 3PM." + vbCrLf + _
    "tbrSoft se reserva el derecho a modificar este contrarto en el futuro. Las licencias de este software" + _
    " son validas solo para un equipo, se hace referecia a equipo por su microprocesador y su placa base " + _
    "(motherboard). Po lo tanto si se reemplaza uno de estos componentes la licencia perdera valor ya que el equipo" + _
    " no sera el mismo"
        
End Sub
