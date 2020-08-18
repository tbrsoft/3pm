VERSION 5.00
Begin VB.Form frmOnlyContador 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contador de 3PM"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   1125
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   4440
      Width           =   4845
   End
   Begin VB.Label lblContador3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20264536538"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   3270
      TabIndex        =   6
      Top             =   3660
      Width           =   2100
   End
   Begin VB.Label lblContador2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20264536538"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   915
      Left            =   210
      TabIndex        =   5
      Top             =   2730
      Width           =   5160
   End
   Begin VB.Label lblContador 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20264536538"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   915
      Left            =   180
      TabIndex        =   4
      Top             =   270
      Width           =   5160
   End
   Begin VB.Label lblPESOS 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$ 888.888.888"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   180
      TabIndex        =   3
      Top             =   1200
      Width           =   5160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmOnlyContador.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   585
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   5115
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contador histórico:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   34
      Left            =   0
      TabIndex        =   1
      Top             =   2490
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contador reiniciable:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   25
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   1815
   End
End
Attribute VB_Name = "frmOnlyContador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case TeclaDER
            Unload Me
            F1.Show 1
        Case Else
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Traducir 'Agregado por el complemento traductor
    lblContador.Caption = STRceros(CONTADOR, 11) 'reini
    lblContador2.Caption = STRceros(CONTADOR2, 11) 'hist
    lblPESOS.Caption = "$ " + CStr(Round(CONTADOR * PrecioBase / TemasPorCredito, 2))
    lblContador3.Caption = STRceros(ValidarCada - CreditosValidar, 11)
    
    Label1(1).Caption = "Tecla izquierda SALIR" + vbCrLf + "Tecla derecha Ver estadisticas"
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Label1(0).Caption = TR.Trad("Si ha cambiado los precios y el valor de " + _
        "cada señal del monedero sin poner en cero este contador el valor " + _
        "en $ puede estar erroneo.%98%Antes de pòner en cero el contador reiniciable%99%")
    Label1(34).Caption = TR.Trad("Contador histórico%99%")
    Label1(25).Caption = TR.Trad("Contador reiniciable%99%")
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub
