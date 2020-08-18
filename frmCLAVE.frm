VERSION 5.00
Begin VB.Form frmCLAVE 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1890
      TabIndex        =   2
      Top             =   3090
      Width           =   2145
   End
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
      Left            =   540
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2580
      Width           =   5145
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00C0FFFF&
      Height          =   1725
      Left            =   240
      TabIndex        =   4
      Top             =   180
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SEGUN CODIGO:"
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
      Height          =   285
      Left            =   870
      TabIndex        =   3
      Top             =   2280
      Width           =   4395
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   5265
      Left            =   150
      Top             =   120
      Width           =   5805
   End
   Begin VB.Image Image1 
      Height          =   1635
      Left            =   2220
      Picture         =   "frmCLAVE.frx":0000
      Top             =   3660
      Width           =   1500
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   870
      TabIndex        =   1
      Top             =   2010
      Width           =   4395
   End
End
Attribute VB_Name = "frmCLAVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    ClaveIngresada = txtPSW
    Unload Me
End Sub

Private Sub Form_Activate()
    Select Case IDIOMA
        Case "Español"
            Label1 = "Ingrese su clave de administrador"
            Command1.Caption = "OK"
        Case "English"
            Command1.Caption = "OK"
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

Private Sub Form_Load()
    MostrarCursor True
    Label2.Caption = "SEGUN CODIGO: " + CodigoParaClaveActual
    
    tERR.Anotar "acfk"
    Dim QuedanC As Long
    QuedanC = ValidarCada - CreditosValidar
    If QuedanC > 0 Then
        'CodigoParaClaveActual busca el archivo con el numero que corresponde validar en este periodo de control
        Label3.Caption = "Ingrese a continuacion su clave para continuar utilizando 3PM. " + vbCrLf + _
            "Debe enviar la administrador el codigo: " + vbCrLf + _
            CodigoParaClaveActual + vbCrLf + _
            "Puede todavia omitir esta clave. Solo le quedan " + CStr(QuedanC) + " creditos hasta que 3PM se inhabilite"
    Else
        Label3.Caption = "De no ingresar la clave correspondiente 3PM no podra continuar. Ha llegado al limite de creditos posibles"
    End If
    
End Sub
