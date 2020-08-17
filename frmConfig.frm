VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig 
   BackColor       =   &H00404000&
   Caption         =   "Configuración del sistema"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7740
      TabIndex        =   24
      Top             =   8100
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   345
      Left            =   3870
      TabIndex        =   23
      Top             =   3900
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Borrar todo"
      Height          =   345
      Left            =   3870
      TabIndex        =   22
      Top             =   3540
      Width           =   1425
   End
   Begin VB.ListBox lstTemas 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7410
      Left            =   7380
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   360
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog CmDlg 
      Left            =   2850
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCarpeta 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5220
      TabIndex        =   19
      Top             =   2730
      Width           =   405
   End
   Begin VB.CommandButton cmdModif 
      Caption         =   "Modificar"
      Height          =   345
      Left            =   5310
      TabIndex        =   18
      Top             =   3900
      Width           =   1425
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   345
      Left            =   5310
      TabIndex        =   17
      Top             =   3540
      Width           =   1425
   End
   Begin VB.TextBox txtCarpeta 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   3000
      Width           =   3705
   End
   Begin VB.ComboBox cmbEstilo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmConfig.frx":030A
      Left            =   5700
      List            =   "frmConfig.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   990
      Width           =   1485
   End
   Begin VB.ComboBox cmbAno 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmConfig.frx":030E
      Left            =   4740
      List            =   "frmConfig.frx":0310
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   990
      Width           =   885
   End
   Begin VB.TextBox txtIdDisco 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3390
      TabIndex        =   8
      Top             =   960
      Width           =   1275
   End
   Begin VB.TextBox txtNombreDisco 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3510
      TabIndex        =   6
      Top             =   2310
      Width           =   3675
   End
   Begin VB.TextBox txtInterprete 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3450
      TabIndex        =   3
      Top             =   1650
      Width           =   3735
   End
   Begin VB.ListBox lstDiscos 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   8040
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   3285
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Cambiar Nombre de Tema elegido"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   7380
      TabIndex        =   25
      Top             =   7890
      Width           =   4455
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de discos"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   20
      Top             =   90
      Width           =   3285
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Modificar Valores de Disco"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   525
      Index           =   8
      Left            =   3750
      TabIndex        =   16
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "En Carpeta"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   3510
      TabIndex        =   15
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Estilo"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6150
      TabIndex        =   12
      Top             =   720
      Width           =   705
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4740
      TabIndex        =   10
      Top             =   720
      Width           =   705
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "IdDisco"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   9
      Top             =   750
      Width           =   1185
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Disco"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   7
      Top             =   2100
      Width           =   1515
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Intérprete"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3420
      TabIndex        =   5
      Top             =   1410
      Width           =   1185
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de temas"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7470
      TabIndex        =   4
      Top             =   60
      Width           =   4215
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
      Left            =   3480
      TabIndex        =   2
      Top             =   8100
      Width           =   3615
   End
   Begin VB.Label lblGD 
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
      Left            =   3480
      TabIndex        =   1
      Top             =   7860
      Width           =   3615
   End
   Begin VB.Image TapaCD 
      Height          =   3240
      Left            =   3510
      Picture         =   "frmConfig.frx":0312
      Stretch         =   -1  'True
      Top             =   4620
      Width           =   3600
   End
End
Attribute VB_Name = "frmCOnfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCarpeta_Click()
    CmDlg.CancelError = True
    On Error GoTo Cancelo
    CmDlg.ShowOpen
    txtCarpeta = FSO.GetParentFolderName(CmDlg.FileName)
    If Right(txtCarpeta, 1) <> "\" Then txtCarpeta = txtCarpeta + "\"
    If ArchivosInFolder("mp3", txtCarpeta) = 0 Then
        MsgBox "No hay ningun archivo válido en esta carpeta"
    Else
        'mostrar el detalle de temas y tapa
        crgArchInLst frmTemas.lstTemas, "mp3", txtCarpeta
        Dim ArchTapa As String
        ArchTapa = txtCarpeta + "tapa.jpg"
        If FSO.FileExists(ArchTapa) = False Then
            'queda en blanco
            FSO.CopyFile AP + "img\tapa.jpg", txtCarpeta + "TAPA.JPG"
            frmTemas.TapaCD.Picture = LoadPicture(ArchTapa)
        End If
        frmTemas.TapaCD.Picture = LoadPicture(ArchTapa)
        frmTemas.Show 1
    End If
    
Cancelo:
    
End Sub

Private Sub cmdNuevo_Click()
    Dim t(10) As String
    t(0) = NewIdDisco: t(1) = txtInterprete: t(2) = txtNombreDisco: t(3) = cmbAno
    t(4) = cmbEstilo: t(5) = txtCarpeta
    If t(1) = "" Then MsgBox "Debe especificar un intérprete": Exit Sub
    If t(2) = "" Then MsgBox "Debe especificar el nombre del disco": Exit Sub
    If t(1) = "" Then MsgBox "Debe especificar un intérprete": Exit Sub
    'cargar el disco
    With rsDiscos
        .AddNew
        !IdDisco = t(0)
        !Interprete = t(1)
        !NombreDisco = t(2)
        !ano = t(3)
        !TipoMusica = t(4)
        !EnCarpeta = t(5)
        .Update
    End With
End Sub

Private Sub Form_Load()
    crgRSinLST rsDiscos, lstDiscos, 1, 2, -1
    'cargar los cmbs
    Dim ano As Integer
    ano = 1940
    Do While ano < 2010
        cmbAno.AddItem Trim(Str(ano))
        ano = ano + 1
    Loop
    
    crgRSinCMB rsEstilo, 0, cmbEstilo
    cmbEstilo = "(Otros)"
End Sub

Private Sub txtInterprete_Change()
    lblGD = txtInterprete
End Sub

Private Sub txtNombreDisco_Change()
    lblNDE = txtNombreDisco
End Sub
