VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmCLUF 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CLUF 3PM"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton Command1 
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   4470
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   873
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "OK"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   4155
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   180
      Width           =   5415
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

Private Sub Form_Activate()
    Command1.Caption = TR.Trad("OK%99%")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        
        Case TeclaCerrarSistema
            YaCerrar3PM
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = TeclaNewFicha Then
        'si ya hay 9 cargados se traga las fichas
        If CREDITOS <= MaximoFichas Then
            SetKeyState vbKeyScrollLock, True
            VarCreditos CSng(TemasPorCredito)
        Else
            'apagar el fichero electronico
            SetKeyState vbKeyScrollLock, False
        End If
    End If
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
    Traducir 'Agregado por el complemento traductor
    Dim TMP As String
    TMP = 1
    TMP = TR.Trad("CLUF - Contrato de licencia de usuario final." + vbCrLf + _
        "Antes de adquirir y utilizar 3PM deberá estar de acuerdo " + _
        "y aceptar las siguientes condiciones.%99%") + vbCrLf + vbCrLf
        
    TMP = TMP + TR.Trad(" TbrSoft de ninguna manera será responsable por el uso " + _
        "dado al sistema por los usuarios finales. La licencia para uso " + _
        "de 3PM será revocada inmediatamente si algún usuario violara las " + _
        "leyes vigentes (respectivas al país en que se utilice) respecto a " + _
        "los derechos de los autores de las composiciones reproducidas por " + _
        "3PM. En todos los casos se deberá obtener una autorización para " + _
        "la reproducción de todos los ficheros mp3 que se incluyan." + vbCrLf + _
        "El costo estipulado por las instituciones y/o asociaciones de " + _
        "autores no es responsabilidad de tbrSoft si no de los usuarios " + _
        "de 3PM.%99%") + vbCrLf
        
    TMP = TMP + TR.Trad("La adquisición de 3PM no implica derechos de reventa de copias " + _
        "ilegales de este software ni la instalación en más de un equipo. " + _
        "En caso de disponer de varios equipos deberán solicitar " + _
        "igual cantidad de copias de 3PM." + vbCrLf + _
        "En ningun caso podra someter a 3PM " + _
        "a metodos de decompilación y similares.%99%")
        
    TR.SetVars "Andres Vazquez Flexes (Argentino, DNI n° 26.453.653)"
    
    TMP = TMP + vbCrLf + TR.Trad("El codigo fuente de este " + _
        "programa es propiedad de %01% quien es titular unico de los " + _
        "mismos." + vbCrLf + _
        " La instalacion de 3PM y las consecuentes modificaciones que este " + _
        "software provoca en el sistema son responsabilidad exclusiva de " + _
        "quien instala este software y no de tbrSoft. tbrSoft no se hace " + _
        "responsable por las consecuencias de ningun tipo derivadas de la " + _
        "instalacion de 3PM.%99%") + vbCrLf
        
    TMP = TMP + TR.Trad("tbrSoft se reserva el derecho a " + _
        "modificar este contrarto en el futuro como condición " + _
        "para recibir actualizaciones." + vbCrLf + _
        "Las licencias de este software son validas solo para un " + _
        "equipo, se hace referecia a equipo por su microprocesador, su placa " + _
        "base (motherboard) y su/s disco/s rígidos. Por lo tanto si se " + _
        "reemplaza uno de estos componentes la licencia perderá valor ya " + _
        "que el equipo no sera el mismo%99%") + vbCrLf
        
    TMP = TMP + TR.Trad("La licencia de 3PM perderá valor si fuera utilizada " + _
        "fuera del país o zona habilitada para el distribuidor que le " + _
        "haya vendido a usted su licencia." + vbCrLf + _
        " El codigo fuente de 3PM no es parte de la licencia que se " + _
        "ofrece, esta es solo de uso de los " + _
        "binarios derivados del codigo compilado. Las modificaciones " + _
        "hechas al sistema se ofrecen en general sin cargo más tbrSoft se reserva " + _
        "el derecho de cobrar por " + _
        "el agregado de funcionalidades adicionales a las que obtuvo " + _
        "al adquirir este software.%99%")
        
       Text1.Text = TMP
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Command1.Caption = TR.Trad("OK%99%")
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub
