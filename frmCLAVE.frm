VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmCLAVE 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton Command1 
      Height          =   555
      Left            =   1740
      TabIndex        =   4
      Top             =   3000
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
      Top             =   2520
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
      ForeColor       =   &H00E0E0E0&
      Height          =   1755
      Left            =   240
      TabIndex        =   3
      Top             =   180
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Según Código"
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
      Left            =   870
      TabIndex        =   2
      Top             =   2220
      Width           =   4395
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   5895
      Left            =   150
      Top             =   120
      Width           =   5805
   End
   Begin VB.Image Image1 
      Height          =   2610
      Left            =   1530
      Picture         =   "frmCLAVE.frx":0000
      Top             =   3450
      Width           =   3375
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
      Left            =   870
      TabIndex        =   1
      Top             =   1950
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
    Label1 = TR.Trad("Ingrese su clave de administrador%99%")
    Command1.Caption = TR.Trad("OK%99%")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case TeclaCerrarSistema
            YaCerrar3PM
    End Select
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
    Traducir 'Agregado por el complemento traductor
    MostrarCursor True
    Label2.Caption = TR.Trad("Según Código: %98%Se pide que se ingrese una " + _
        "clave segun un codigo generado por el sistema. El dueño de la " + _
        "licencia tiene como generar este código%99%") + CodigoParaClaveActual
    
    tERR.Anotar "acfk"
    Dim QuedanC As Long
    QuedanC = ValidarCada - CreditosValidar
    If QuedanC > 0 Then
        'CodigoParaClaveActual busca el archivo con el numero que corresponde validar en este periodo de control
        TR.SetVars CodigoParaClaveActual, QuedanC, "3PM"
        Label3.Caption = TR.Trad("Ingrese a continuación su clave para continuar " + _
            "utilizando %03%. " + vbCrLf + _
            "Debe enviar la administrador el codigo: " + vbCrLf + _
            "%01%" + vbCrLf + _
            "Puede todavia omitir esta clave. Solo le quedan %02%" + _
            " creditos hasta que %03% se inhabilite" + _
            "%98%La variable 1 es un codigo que genera la pc y que el usuario " + _
            "debe enviar al dueño para poder seguir usando las fonola" + vbCrLf + _
            "La variable 2 son los créditos que aún quedan para seguir " + _
            "usando el equipom sin que se bloquee." + vbCrLf + _
            "la variable 3 dice 3PM%99%")
    Else
        Label3.Caption = TR.Trad("De no ingresar la clave correspondiente 3PM no " + _
            "podra continuar. Ha llegado al limite de creditos posibles%99%")
    End If
    
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Command1.Caption = TR.Trad("OK%99%")
    Label3.Caption = TR.Trad("Ingrese su clave para continuar%99%")
    Label2.Caption = TR.Trad("Según Código: %98%Se pide que se ingrese una " + _
        "clave segun un codigo generado por el sistema. El dueño de la " + _
        "licencia tiene como generar este código%99%")
    Label1.Caption = TR.Trad("Ingrese su clave para continuar%99%")
End Sub
