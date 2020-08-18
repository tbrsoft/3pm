VERSION 5.00
Begin VB.Form frmREG 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de 3PM"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picXaANCHOVUM 
      AutoSize        =   -1  'True
      Height          =   675
      Left            =   11220
      ScaleHeight     =   615
      ScaleWidth      =   375
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   4230
      Picture         =   "frmREG.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4380
      Width           =   645
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   4230
      Picture         =   "frmREG.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   645
   End
   Begin VB.TextBox LBL 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2445
      Left            =   5190
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   90
      Width           =   5865
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4230
      Picture         =   "frmREG.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7470
      Width           =   645
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   4230
      Picture         =   "frmREG.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   645
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "He leido y estoy de acuerdo con el Contrato de Licencia de Usuario Final"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   465
      Left            =   4920
      TabIndex        =   5
      Top             =   7020
      Value           =   1  'Checked
      Width           =   5055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4230
      Picture         =   "frmREG.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6870
      Width           =   645
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "cerrar 3pm"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   11010
      Picture         =   "frmREG.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7230
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   4230
      Picture         =   "frmREG.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2580
      Width           =   645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Si ha solicitado ya clave gratuita o paga puede cargarla desde aquí."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   435
      Index           =   2
      Left            =   4920
      TabIndex        =   18
      Top             =   5610
      Width           =   6795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmREG.frx":1546
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   405
      Index           =   9
      Left            =   4920
      TabIndex        =   17
      Top             =   3780
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmREG.frx":15ED
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   615
      Index           =   8
      Left            =   4920
      TabIndex        =   16
      Top             =   2880
      Width           =   6915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmREG.frx":1694
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   525
      Index           =   3
      Left            =   4920
      TabIndex        =   15
      Top             =   4710
      Width           =   6825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Obtener archivo para pedir licencia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Index           =   7
      Left            =   4920
      TabIndex        =   14
      Top             =   4440
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargar archivo de licencia recibido"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   315
      Index           =   6
      Left            =   4920
      TabIndex        =   13
      Top             =   5310
      Width           =   6795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   1
      X1              =   5640
      X2              =   9750
      Y1              =   6420
      Y2              =   6420
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recuperar/Reparar (tecla Izquierda 6 veces)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   5
      Left            =   4920
      TabIndex        =   11
      Top             =   7680
      Width           =   5475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ver contrato de licencia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Index           =   4
      Left            =   4920
      TabIndex        =   10
      Top             =   6840
      Width           =   2745
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INICIAR PROGRAMA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   315
      Index           =   1
      Left            =   4920
      TabIndex        =   9
      Top             =   2580
      Width           =   7245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Abrir Manual de uso."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   315
      Index           =   0
      Left            =   4920
      TabIndex        =   8
      Top             =   3510
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   8400
      Left            =   0
      Stretch         =   -1  'True
      Top             =   30
      Width           =   4620
   End
End
Attribute VB_Name = "frmREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GGG As String
Dim LastTeclas As String
Dim CarpetaSys As String

Private Sub Check1_Click()
    Command1.Enabled = Check1
End Sub

Private Sub Command1_Click()

   'dar ingreso a la clave y la grabo solo aqui
    K.IngresaClave 'aqui se carga mLicencia !!!
        
    If K.LICENCIA = aSinCargar Then
        Dim TXTmsg As String
        TXTmsg = "3PM en version Demo tiene limite de discos, " + _
            "y trunca los temas a los 2 minutos." + vbCrLf + _
            "¿Desea ejecutar 3PM en version demo?"
        If MsgBox(TXTmsg, vbQuestion + vbYesNo, "3PM demo") = vbNo Then End
    End If
        
    If K.LICENCIA = BErronea Then
        MsgBox "Existen datos erroneos de la licencia. Si ingresa claves equivocadas o ha" + _
            " reemplazado componentes de su PC debe comunicarse con tbrSoft o su proveedor" + _
            " de 3PM para solucionar este inconveniente"
            Exit Sub
    End If
    
    If K.LICENCIA = CGratuita Then MsgBox "Clave gratuita de 3PM. "
    
    If K.LICENCIA = GFull Then MsgBox "El codigo se ha cargado correctamente. Bienvenido a 3PM "
    
    If K.LICENCIA = HSuperLicencia Then MsgBox "SUPERLICENCIA de 3PM. El codigo de SuperLicencia se ha cargado correctamente. Bienvenido a Super3PM"
    
    Unload Me
    frmINI.Show 1
                    
    tERR.Anotar "acpg"
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acph"
    Resume Next
End Sub

