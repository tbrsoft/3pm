VERSION 5.00
Begin VB.Form frmCLAVE 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   4455
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
      Left            =   1140
      TabIndex        =   2
      Top             =   1350
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
      Left            =   1110
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   780
      Width           =   2115
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   3855
      Left            =   150
      Top             =   120
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   1635
      Left            =   1380
      Picture         =   "frmCLAVE.frx":0000
      Top             =   2130
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese su clave de administrador"
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
      Height          =   465
      Left            =   720
      TabIndex        =   1
      Top             =   270
      Width           =   3045
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

