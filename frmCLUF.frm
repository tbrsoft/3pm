VERSION 5.00
Begin VB.Form frmCLUF 
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
      Width           =   2700
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
    Select Case IDIOMA
        Case "Espa�ol"
            Command1.Caption = "OK"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
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
    Select Case IDIOMA
        Case "Espa�ol"
            Text1.Text = "CLUF - Contrato de licencia de usuario final." + vbCrLf + vbCrLf + "Antes de adquirir y utilizar 3PM deber� estar de acuerdo y aceptar las siguientes condiciones." + vbCrLf + vbCrLf + _
                " TbrSoft de ninguna manera ser� responsable por el uso dado al sistema por los usuarios finales. La licencia para uso de 3PM ser� revocada inmediatamente si alg�n usuario violara las leyes vigentes " + _
                "(respectivas al pa�s en que se utilice) respecto a los derechos de los autores de las composiciones reproducidas por 3PM. En todos los casos se deber� obtener una autorizaci�n para la reproducci�n de " + _
                "todos los ficheros mp3 que se incluyan." + vbCrLf + "El costo estipulado por las instituciones y/o asociaciones de autores no es responsabilidad de tbrSoft si no de los usuarios de 3PM." + vbCrLf + _
                "La adquisici�n de 3PM no implica derechos de reventa de copias ilegales de este software ni la instalaci�n en m�s de un equipo (salvo que la licencia adquirida asi lo indique). En caso de disponer " + _
                "de varios equipos deber�n solicitar igual cantidad de copias de 3PM." + vbCrLf + "En ningun caso podra someter a 3PM a metodos de decompilaci�n y similares. El codigo fuente de este " + _
                "programa es propiedad de Andres Vazquez Flexes (Argentino, DNI n� 26.453.653) quien es titular unico de los mismos." + vbCrLf + " La instalacion de 3PM y las consecuentes modificaciones" + _
                " que este software provoca en el sistema son responsabilidad exclusiva de quien instala este software y no de tbrSoft. tbrSoft no se hace responsable" + _
                " por las consecuencias de ningun tipo derivadas de la instalacion de 3PM." + vbCrLf + "tbrSoft se reserva el derecho a modificar este contrarto en el futuro. Las licencias de este software" + _
                " son validas solo para un equipo, se hace referecia a equipo por su microprocesador, su placa base (motherboard) y su/s disco/s r�gidos." + _
                " Por lo tanto si se reemplaza uno de estos componentes la licencia perdera valor ya que el equipo" + _
                " no sera el mismo" + vbCrLf + "La licencia de 3PM perder� valor si fuera utilizada fuera del pais o zona" + _
                " habilitada para el distribuidor que le haya vendido a usted su licencia." + vbCrLf + vbCrLf + " El codigo fuente de 3PM no es parte de la licencia que se ofrece, esta es solo de uso de los binarios derivados del codigo. " + _
                "Las modificaciones hechas al sistema se ofrecen sin cargo mas tbrSoft se reserva el derecho de cobrar por el agregado de funcionalidades adicionales a las que obtuvo al adquirir este software."
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
        
End Sub
