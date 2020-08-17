VERSION 5.00
Begin VB.Form frmCDS 
   BackColor       =   &H000040C0&
   Caption         =   "Rockola 2003"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmCDS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Guns N' Roses"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   8280
      TabIndex        =   14
      Top             =   7050
      Width           =   3615
   End
   Begin VB.Label lblNDE 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Live era '87 - '93"
      BeginProperty Font 
         Name            =   "Belwe Bd BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   8280
      TabIndex        =   13
      Top             =   7290
      Width           =   3615
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Guns N' Roses"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   4530
      TabIndex        =   12
      Top             =   7050
      Width           =   3615
   End
   Begin VB.Label lblNDE 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Live era '87 - '93"
      BeginProperty Font 
         Name            =   "Belwe Bd BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4530
      TabIndex        =   11
      Top             =   7290
      Width           =   3615
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Guns N' Roses"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   8280
      TabIndex        =   10
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label lblNDE 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Live era '87 - '93"
      BeginProperty Font 
         Name            =   "Belwe Bd BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   8280
      TabIndex        =   9
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Shape SHsel 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   3225
      Left            =   630
      Top             =   660
      Width           =   3585
   End
   Begin VB.Image TapaCD 
      Height          =   3240
      Index           =   4
      Left            =   8280
      Picture         =   "frmCDS.frx":0442
      Stretch         =   -1  'True
      Top             =   3810
      Width           =   3600
   End
   Begin VB.Image TapaCD 
      Height          =   3240
      Index           =   3
      Left            =   4530
      Picture         =   "frmCDS.frx":CF48
      Stretch         =   -1  'True
      Top             =   3810
      Width           =   3600
   End
   Begin VB.Label lblNDE 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Live era '87 - '93"
      BeginProperty Font 
         Name            =   "Belwe Bd BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4530
      TabIndex        =   8
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Guns N' Roses"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   4530
      TabIndex        =   7
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Image TapaCD 
      Height          =   3240
      Index           =   1
      Left            =   4530
      Picture         =   "frmCDS.frx":13517
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3600
   End
   Begin VB.Image TapaCD 
      Height          =   3240
      Index           =   2
      Left            =   8310
      Picture         =   "frmCDS.frx":1716D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3600
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmCDS.frx":1948F
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   990
      TabIndex        =   3
      Top             =   8340
      Width           =   9915
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Temas"
      BeginProperty Font 
         Name            =   "Onyx BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   10920
      TabIndex        =   5
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Onyx BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   525
      Left            =   10920
      TabIndex        =   4
      Top             =   8070
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808000&
      BorderWidth     =   5
      Index           =   1
      X1              =   0
      X2              =   11850
      Y1              =   7650
      Y2              =   7650
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Créditos"
      BeginProperty Font 
         Name            =   "Onyx BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label lblCreditos 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Onyx BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   8070
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Se viene: Madres - Los Caballeros de la quema - En vivo en Obras"
      BeginProperty Font 
         Name            =   "Amerigo BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   6
      Top             =   8010
      Width           =   10725
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reproduciendo tema: Lo fragil de la locura - La Renga - La Renga. Tiempo restante 3:14"
      BeginProperty Font 
         Name            =   "Amerigo BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   7680
      Width           =   10725
   End
End
Attribute VB_Name = "frmCDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'teclas habilitadas
    'a = IZQ
    's = DER
    'l = Seleccionar
    'k = volver, escape
    'z=configuracion
    
    Select Case KeyCode
        Case vbKeyA: MoverIZQ
        Case vbKeyS: MoverDER
        'Case vbKeyL: MoverOK
        'Case vbKeyK: MoverESC
        Case vbKeyZ: frmPSW.Show
    End Select
End Sub

Private Sub Form_Load()
    AP = App.Path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    'cargar la base de datos
    Dim cnSTR As String
    cnSTR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AP + "musica.mdb;Persist Security Info=False"
    'si no pongo esto no me cuenta los recordCount
    CN.CursorLocation = adUseClient
    CN.Open cnSTR
    rsDiscos.Open "select * from tblDiscos", CN, adOpenDynamic, adLockPessimistic, adCmdText
    rsTemas.Open "select * from tblTemas", CN, adOpenDynamic, adLockPessimistic, adCmdText
    rsEstilo.Open "select * from tblTipoMusica", CN, adOpenDynamic, adLockPessimistic, adCmdText
    
    'marcar algun disco
    nDiscoElegido = 1
    'definir acion para posibles movimientos}
    EstoyHaciendo = "ViendoDiscos"
    ElegirDisco 1, SHsel
    
End Sub

Public Function Cargar4Discos(DiscoINI As Variant) As Integer
    'indica cantidad de discos mostrados
    Cargar4Discos = 0
    rsDiscos.Sort = "Interprete"
    'si no hay más no hacer nada
    If DiscoINI > rsDiscos.RecordCount Then Exit Function
    'ocultar todo
    Dim c As Integer
    c = 1
    Do While c < 5
        TapaCD(c).Visible = False
        lblGDE(c).Visible = False
        lblNDE(c).Visible = False
        c = c + 1
    Loop
    'cargar las imágenes y los datos de los 4 primeros discos
    With rsDiscos
        .Bookmark = DiscoINI
        c = 1
        Do While c < 5
            'la tapa siempre se llama tapa.jpg y esta en la carpeta indicada
            'EnCarpeta debe tener la barra final "\"
            TapaCD(c).Picture = LoadPicture(!EnCarpeta + "tapa.jpg")
            lblGDE(c) = !Interprete
            lblNDE(c) = !NombreDisco
            Cargar4Discos = c
            .MoveNext
            If .EOF Then Exit Do
            c = c + 1
        Loop
    End With
End Function
