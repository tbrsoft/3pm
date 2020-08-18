VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmIndex 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frDISCOS 
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   930
      TabIndex        =   29
      Top             =   60
      Width           =   4020
      Begin VB.PictureBox picVideo 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   570
         ScaleHeight     =   435
         ScaleWidth      =   915
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Left            =   1620
         Top             =   60
      End
      Begin VB.Timer Timer3 
         Interval        =   10000
         Left            =   2400
         Top             =   75
      End
      Begin tbr3pm.MP3Play MP3 
         Height          =   1620
         Left            =   3270
         TabIndex        =   31
         Top             =   450
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   2858
      End
      Begin VB.Image TapaCD 
         Height          =   2505
         Index           =   0
         Left            =   570
         Picture         =   "frmINDEX.frx":0000
         Stretch         =   -1  'True
         Top             =   210
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Label lblDisco 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Complete al menos la primera hoja de discos cargados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   0
         Left            =   540
         TabIndex        =   32
         Top             =   2730
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Shape lblSel 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Height          =   555
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   810
         Width           =   435
      End
   End
   Begin tbr3pm.VUMeter VU1 
      Height          =   4425
      Left            =   5550
      TabIndex        =   24
      Top             =   1110
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   7805
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Caption         =   "Touch"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2135
      Left            =   9780
      TabIndex        =   17
      Top             =   6840
      Width           =   2200
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
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
         Height          =   525
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   210
         Width           =   2030
      End
      Begin VB.CommandButton cmdPagAt 
         BackColor       =   &H00C0C0C0&
         DownPicture     =   "frmINDEX.frx":9013
         Height          =   650
         Left            =   90
         Picture         =   "frmINDEX.frx":A072
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1425
         Width           =   1000
      End
      Begin VB.CommandButton cmdDiscoAt 
         BackColor       =   &H00C0C0C0&
         DownPicture     =   "frmINDEX.frx":AB23
         Height          =   650
         Left            =   90
         Picture         =   "frmINDEX.frx":B895
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   750
         Width           =   1000
      End
      Begin VB.CommandButton cmdDiscoAd 
         BackColor       =   &H00C0C0C0&
         DownPicture     =   "frmINDEX.frx":C1D8
         Height          =   650
         Left            =   1110
         Picture         =   "frmINDEX.frx":CED5
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   750
         Width           =   1000
      End
      Begin VB.CommandButton cmdPagAd 
         BackColor       =   &H00C0C0C0&
         DownPicture     =   "frmINDEX.frx":D7AD
         Height          =   650
         Left            =   1110
         Picture         =   "frmINDEX.frx":E76E
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1425
         Width           =   1000
      End
   End
   Begin VB.ListBox lstProximos 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      IntegralHeight  =   0   'False
      Left            =   3570
      TabIndex        =   16
      Top             =   7320
      Width           =   6195
   End
   Begin VB.Frame frModoVideo 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   10290
      TabIndex        =   9
      Top             =   420
      Visible         =   0   'False
      Width           =   1715
      Begin VB.Label L 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nombre del artista - nombre del disco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1700
      End
   End
   Begin VB.Frame frTEMAS 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   10290
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   1715
      Begin VB.Label T 
         BackColor       =   &H0080FFFF&
         Caption         =   "Nombre del TEMA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1700
      End
   End
   Begin MSCommLib.MSComm COMM1 
      Left            =   1350
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   1
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.Label lblTECLAS 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "11111222223333344444"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   10530
      TabIndex        =   28
      Top             =   6495
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblNoTecla 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   10530
      TabIndex        =   27
      Top             =   6270
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblNoUSO 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   11235
      TabIndex        =   26
      Top             =   6270
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblValidar 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Validar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   6300
      TabIndex        =   25
      Top             =   7110
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label lblPuesto 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rank #888"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   30
      TabIndex        =   6
      Top             =   6840
      Width           =   1440
   End
   Begin VB.Label lblPrecios 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1 coin = 3 creditos  2 creditos = 1 tema"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   1500
      TabIndex        =   23
      Top             =   7920
      Width           =   2055
   End
   Begin VB.Label lblCreditos 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Creditos 00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   30
      TabIndex        =   2
      Top             =   7320
      Width           =   1440
   End
   Begin VB.Label lblTemaSonando 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sin Reproducción actual"
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
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Top             =   6840
      UseMnemonic     =   0   'False
      Width           =   8265
   End
   Begin VB.Label lblTEMAS 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Temas del disco elegido"
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
      Height          =   465
      Left            =   10260
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label lblModoVideo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Discos en Modo Video"
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
      Height          =   435
      Left            =   10260
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label lblPag 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pagina 88 de 88"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   1500
      TabIndex        =   7
      Top             =   7620
      Width           =   2055
   End
   Begin VB.Label lblV 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "version 8.88.888"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   8700
      Width           =   1335
   End
   Begin VB.Label lblTiempoRestante 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Falta: 00:00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   30
      TabIndex        =   1
      Top             =   7080
      Width           =   1440
   End
   Begin VB.Label LBLpORCtEMA 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   1500
      TabIndex        =   4
      Top             =   7170
      Width           =   8265
   End
   Begin VB.Label lblTOTdiscos 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Disco 188 de 188"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   1500
      TabIndex        =   3
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label lblDEMO 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Solicite la version definitiva a info@tbrsoft.com / avazquez@cpcipc.org"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1500
      TabIndex        =   8
      Top             =   8670
      UseMnemonic     =   0   'False
      Width           =   8280
   End
   Begin VB.Label lblTBR 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Software desarrollado por tbrSoft www.tbrsoft.com - info@tbrsoft.com - avazquez@cpcipc.org."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1485
      TabIndex        =   15
      Top             =   8430
      Width           =   8280
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1350
      Left            =   30
      Picture         =   "frmINDEX.frx":F17B
      Stretch         =   -1  'True
      Top             =   7620
      Width           =   1440
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ModoVideoSelTema As Boolean 'si estoy envideo
'saber si estoy eligiendo tema. Sino estoy en disco

Dim TemaElegidoModoVideo As Integer

Dim LastDiscoSel As Long
Dim DiscosEnPagina As Long

Dim VolBajando As Double 'bajando volumen para terminar tema demo
Dim LastpSeconds As Long 'comparador para bajar de a uno el volumen en demos

Dim Ancho As Long, Variacion As Long 'PARA la barra de proceso del tema
Public DuracionTema As Long 'duracion de todos los tenmas de un disco
Dim TotalTema As Long 'duracion total
Dim nDiscoSEL As Long 'del 0 al 5

Private Sub cmdDiscoAd_Click()
    Form_KeyDown TeclaDER, 0
    Command1.SetFocus
End Sub

Private Sub cmdDiscoAd_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub cmdDiscoAt_Click()
    Form_KeyDown TeclaIZQ, 0
    Command1.SetFocus
End Sub

Private Sub cmdDiscoAt_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub cmdPagAd_Click()
    Form_KeyDown TeclaPagAd, 0
    Command1.SetFocus
End Sub

Private Sub cmdPagAd_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub cmdPagAt_Click()
    Form_KeyDown TeclaPagAt, 0
    Command1.SetFocus
End Sub

Private Sub cmdPagAt_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub COMM1_OnComm()
    'identificar la señal
    Select Case CStr(COMM1.CommEvent)
        Case "3" 'cargar un credito
            Form_KeyUp 81, 0
        Case "4" 'subir volumen
            Form_KeyDown 69, 0
        Case "5" 'bajar volumen
            Form_KeyDown 68, 0
        Case "1001" 'terminar "W"
            Form_KeyDown 87, 0
    End Select
    'una vez que carga uno debe parar todo un segundo para no tomar las repeticiones
    COMM1.RThreshold = 0: COMM1.SThreshold = 0
    'se deshabilito la recepcion de señal
    
    'espero un segundo
    CC = Timer
    Do While Timer < CC + 1
        DoEvents
    Loop
    
    'habilito de vuelta
    COMM1.RThreshold = 1: COMM1.SThreshold = 1
    
    
End Sub

Private Sub Command1_Click()
    Form_KeyDown TeclaOK, 0
End Sub

Private Sub Form_Activate()
    MostrarCursor False
    'actualizar los precios
    If TemasPorCredito = 1 Then
        lblPrecios = "1 coin = " + CStr(TemasPorCredito) + " credito"
    Else
        lblPrecios = "1 coin = " + CStr(TemasPorCredito) + " creditos"
    End If
    If CreditosCuestaTema = 1 Then
        lblPrecios = lblPrecios + vbCrLf + "1 credito = 1 tema"
    Else
        lblPrecios = lblPrecios + vbCrLf + CStr(CreditosCuestaTema) + " creditos = 1 tema"
    End If
        
    If HabilitarVUMetro Then
        If VU1.inHabilitado = False Then VU1.DoStart
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'y si no es una ficha la que se esta cargando
    'aqui se regsitran las presiones de las teclas elegidas
    Dim PagNum As Long
    Select Case KeyCode
        'subir o bajar volumen
        Case 68 'es la D y baja el volumen
            If frmIndex.MP3.IsPlaying Then
                If VolumenIni <= 5 Then
                    frmIndex.MP3.Volumen = 0
                Else
                    frmIndex.MP3.Volumen = VolumenIni - 5
                End If
                VolumenIni = frmIndex.MP3.Volumen
            End If
        Case 69 'es la E y sube el volumen
            If frmIndex.MP3.IsPlaying Then
                If VolumenIni >= 95 Then
                    frmIndex.MP3.Volumen = 100
                Else
                    frmIndex.MP3.Volumen = VolumenIni + 5
                End If
                VolumenIni = frmIndex.MP3.Volumen
            End If
        Case 66 ' es la b y pasa al siguiente tema
            'si es video ocultar la pantalla de video
            If EsVideo Then picVideo.Visible = False
            EMPEZAR_SIGUIENTE
        Case TeclaPagAd
            PagNum = nDiscoGral \ (TapasMostradasH * TapasMostradasV)
            Dim PrimeroDeLaPaginaQueSigue As Long
            PrimeroDeLaPaginaQueSigue = (PagNum + 1) * (TapasMostradasH * TapasMostradasV)
            If PrimeroDeLaPaginaQueSigue < TOTAL_DISCOS Then
                If nDiscoSEL <> 0 Then UnSelDisco nDiscoSEL
                DiscosEnPagina = CargarDiscos(PrimeroDeLaPaginaQueSigue, True)
                lblTOTdiscos = "Disco " + CStr(PrimeroDeLaPaginaQueSigue + 1) + " de " + CStr(TOTAL_DISCOS)
                nDiscoSEL = 0
                nDiscoSEL = 0
                TECLAS_PRES = TECLAS_PRES + "5"
                TECLAS_PRES = Right(TECLAS_PRES, 20)
                lblTECLAS = TECLAS_PRES
            End If
        Case TeclaPagAt
            PagNum = nDiscoGral \ (TapasMostradasH * TapasMostradasV)
            If PagNum > 0 Then
                Dim PrimeroDeLaPaginaQueAnterior As Long
                PrimeroDeLaPaginaQueAnterior = (PagNum - 1) * (TapasMostradasH * TapasMostradasV)
                If nDiscoSEL <> 0 Then UnSelDisco nDiscoSEL
                DiscosEnPagina = CargarDiscos(PrimeroDeLaPaginaQueAnterior, False)
                lblTOTdiscos = "Disco " + CStr(PrimeroDeLaPaginaQueAnterior + 1) + " de " + CStr(TOTAL_DISCOS)
                TECLAS_PRES = TECLAS_PRES + "6"
                TECLAS_PRES = Right(TECLAS_PRES, 20)
                lblTECLAS = TECLAS_PRES
            End If
        Case TeclaConfig
            
            
            
             frmConfig.Show 1
        
        Case TeclaIZQ
            'BotonTouch2.Flash
            'ver si desplazo temas en modo video
            If ModoVideoSelTema Then
                If TemaElegidoModoVideo > 0 Then
                    UnSelTema TemaElegidoModoVideo
                    TemaElegidoModoVideo = TemaElegidoModoVideo - 1
                    SelTema TemaElegidoModoVideo
                    OrdenarListaTemaVideo
                End If
            Else
                'no ir a -1
                If nDiscoSEL = 0 Then
                    'ver si hay que pasar hoja o no
                    If PasarHoja Then
                        If nDiscoGral > 0 Then DiscosEnPagina = CargarDiscos(nDiscoGral - ((TapasMostradasH * TapasMostradasV)), False)
                    Else
                        'NO NO NO!!!! nDiscoGral = (TapasMostradasH * TapasMostradasV) - 1
                        'estoy en una hoja al principio y debo elegir el disco del final
                        'sel y unsel trabajan con referencias de o al total de discos por pag
                        'nDiscoGral es el numero absoluto del disco
                        'ver si existe el disco al que voy
                        If TOTAL_DISCOS > nDiscoGral + (TapasMostradasH * TapasMostradasV) - 1 Then
                            nDiscoGral = nDiscoGral + (TapasMostradasH * TapasMostradasV) - 1
                            UnSelDisco nDiscoSEL
                            SelDisco (TapasMostradasH * TapasMostradasV) - 1
                        Else
                            nDiscoGral = TOTAL_DISCOS - 1
                            UnSelDisco nDiscoSEL
                            SelDisco DiscosEnPagina - 1
                        End If
                    End If
                Else
                    nDiscoGral = nDiscoGral - 1
                    UnSelDisco nDiscoSEL
                    SelDisco nDiscoSEL - 1
                End If
                lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
            End If
            TECLAS_PRES = TECLAS_PRES + "1"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            lblTECLAS = TECLAS_PRES
        Case TeclaDER
            If ModoVideoSelTema Then
                If TemaElegidoModoVideo < UBound(MATRIZ_TEMAS) Then
                    UnSelTema TemaElegidoModoVideo
                    TemaElegidoModoVideo = TemaElegidoModoVideo + 1
                    SelTema TemaElegidoModoVideo
                    OrdenarListaTemaVideo
                End If
            Else
                If nDiscoSEL = DiscosEnPagina - 1 Then
                    'ver si hay que pasar hojas
                    If PasarHoja Then
                        If nDiscoGral + 1 < TOTAL_DISCOS Then
                            DiscosEnPagina = CargarDiscos(nDiscoGral + 1, True)
                        End If
                    Else
                        '!!!NO NO NO nDiscoGral = 0
                        'estoy en una hoja al final y debo elegir el disco del principio
                        'sel y unsel trabajan con referencias de o al total de discos por pag
                        'nDiscoGral es el numero absoluto del disco
                        nDiscoGral = nDiscoGral - DiscosEnPagina + 1
                        UnSelDisco nDiscoSEL
                        SelDisco 0
                    End If
                Else
                    If nDiscoGral + 1 < TOTAL_DISCOS Then
                        nDiscoGral = nDiscoGral + 1
                        UnSelDisco nDiscoSEL
                        SelDisco nDiscoSEL + 1
                    End If
                End If
            End If
            lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)

            TECLAS_PRES = TECLAS_PRES + "2"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            lblTECLAS = TECLAS_PRES
        Case TeclaOK
            TECLAS_PRES = TECLAS_PRES + "3"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            lblTECLAS = TECLAS_PRES
            
            If ModoVideoSelTema Then
                'si no dice salir cargar tema
                If T(TemaElegidoModoVideo) = "SALIR" Or T(TemaElegidoModoVideo) = "No hay temas" Then
                    'volver a elegir discos
                    frTEMAS.Visible = False
                    lblTEMAS.Visible = False
                    
                    frModoVideo.Height = frDISCOS.Height - lblModoVideo.Height
                    UnSelTema 0
                    ModoVideoSelTema = False
                Else
                    'ejecutar el tema
                    If CREDITOS >= CreditosCuestaTema Then
                        CREDITOS = CREDITOS - CreditosCuestaTema
                        'siempre que se ejecute un credito estaremos por debajo de maximo
                        OnOffCAPS vbKeyScrollLock, True
                        'grabar cant de creditos
                        EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
                        If CREDITOS < 10 Then frmIndex.lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
                        If CREDITOS >= 10 Then frmIndex.lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
                        
                        'grabar credito para validar
                        'creditosValidar ya se cargo en load de frmindex
                        CreditosValidar = CreditosValidar + TemasPorCredito
                        EscribirArch1Linea SYSfolder + "\radilav.cfg", CStr(CreditosValidar)
                        
                        Dim temaElegido As String
                        'lstext es una lista oculta  con datos completos
                        temaElegido = txtInLista(MATRIZ_TEMAS(TemaElegidoModoVideo), 0, ",")
                        
                        'si esta ejecutando pasa a la lista de reproducción
                        If MP3.IsPlaying Then
                            'pasar a la lista de reproducción
                            Dim NewIndLista As Long
                            NewIndLista = UBound(MATRIZ_LISTA)
                            ReDim Preserve MATRIZ_LISTA(NewIndLista + 1)
                            
                            'se graba en Matriz_Listas como path, nombre(sin .mp3)
                            MATRIZ_LISTA(NewIndLista + 1) = _
                                temaElegido + "," + _
                                FSO.GetBaseName(T(TemaElegidoModoVideo)) + _
                                " / " + FSO.GetBaseName(UbicDiscoActual)
                            CargarProximosTemas
                            'graba en reini.tbr los datos que correspondan por si se corta la luz
                            CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
                            'volver a elegir discos
                            frTEMAS.Visible = False
                            lblTEMAS.Visible = False
                            
                            frModoVideo.Height = frDISCOS.Height - lblModoVideo.Height
                            UnSelTema 0
                            ModoVideoSelTema = False
                        Else
                            'NUNCA ENTRARA AQUI, siempre esta rep video
                            'TEMA_REPRODUCIENDO y mp3.isplayin se cargan en ejecutartema
                            'paciencia
                            CORTAR_TEMA = False 'este tema va entero ya que lo eligio el usuario
                            EjecutarTema temaElegido, True
                        End If
                    End If
                End If
            Else
                'ver si hay que mostrar el frm
                'o estamos en MODO VIDEO
                If EsVideo Then
                    frModoVideo.Height = frDISCOS.Height / 4
                    OrdenarListaModoVideo
                    lblTEMAS.Top = frModoVideo.Top + frModoVideo.Height + 50
                    lblTEMAS.Left = lblModoVideo.Left
                    frTEMAS.Top = lblTEMAS.Top + lblTEMAS.Height
                    frTEMAS.Height = frDISCOS.Height - lblModoVideo.Height - frModoVideo.Height - lblTEMAS.Height - 75
                    lblTEMAS.Visible = True
                    frTEMAS.Visible = True
                    
                    'cargar los temas multimedia en t()
                    ReDim MATRIZ_TEMAS(0) 'matriz en blanco
                    'es una matriz global
                    UbicDiscoActual = txtInLista(MATRIZ_DISCOS(nDiscoGral + 1), 0, ",")
                    'encontrar todos los archivos *.mp3, *.avi, *.mpg, *.mpeg, etc
                    ReDim Preserve MATRIZ_TEMAS(0)
                    MATRIZ_TEMAS = ObtenerArchMM(UbicDiscoActual)
                    If UBound(MATRIZ_TEMAS) = 0 Then
                        T(0) = "No hay temas"
                        SelTema 0
                        ModoVideoSelTema = True
                        Exit Sub
                    End If
                    T(0) = "SALIR"
                    '----------------------------
                    'a daniel cruz le da un error como si se volviera a cargar algo que esta cargado
                    'por lo tanto tengo que poner un manejador de error aqui, unico lugar en que se carga esto
                    For Each LLL In frmIndex.T
                        If LLL.Index > 0 Then Unload LLL
                    Next
                    '----------------------------
                    For AA = 1 To UBound(MATRIZ_TEMAS)
                        Load T(AA)
                        T(AA) = FSO.GetBaseName(txtInLista(MATRIZ_TEMAS(AA), 1, ","))
                        T(AA).Top = T(AA - 1).Top + T(AA - 1).Height
                        T(AA).Left = T(AA - 1).Left
                        T(AA).Visible = True
                    Next
                    TemaElegidoModoVideo = 0
                    SelTema 0
                    ModoVideoSelTema = True
                Else
                    If lblDisco(nDiscoSEL) = "01- Los mas escuchados" Then GoTo TOP10Show
                    frmTemasDeDisco.Show 1
                End If
            End If
        Case TeclaCerrarSistema
            OnOffCAPS vbKeyCapital, False
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            MostrarCursor True
            MP3.DoClose
            If ApagarAlCierre Then APAGAR_PC
            End
        Case TeclaESC
            TECLAS_PRES = TECLAS_PRES + "4"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            lblTECLAS = TECLAS_PRES
            If ModoVideoSelTema Then
                'volver a elegir discos
                frTEMAS.Visible = False
                lblTEMAS.Visible = False
                
                frModoVideo.Height = frDISCOS.Height - lblModoVideo.Height
                UnSelTema 0
                ModoVideoSelTema = False
            End If
    End Select
    VerClaves TECLAS_PRES
    SecSinTecla = 0
    lblNoTecla = 0
    Exit Sub
TOP10Show:
    FRMTOP10.Show 1
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = TeclaNewFicha Then
        'si ya hay 9 cargados se traga las fichas
        If CREDITOS <= MaximoFichas Then
            OnOffCAPS vbKeyScrollLock, True
            CREDITOS = CREDITOS + TemasPorCredito
            SumarContadorCreditos TemasPorCredito
            'grabar cant de creditos
            EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
            If CREDITOS >= 10 Then
                lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
            Else
                lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
            End If
            'grabar credito para validar
            'creditosValidar ya se cargo en load de frmindex
            CreditosValidar = CreditosValidar + TemasPorCredito
            EscribirArch1Linea SYSfolder + "\radilav.cfg", CStr(CreditosValidar)
        Else
            'apagar el fichero electronico
            OnOffCAPS vbKeyScrollLock, False
        End If
    End If
End Sub

Private Sub Form_Load()
     RegistroDiario 'anota la fecha, hora y numero del contador
     '--------
    If TypeVersion = "SL" Then
        If FSO.FileExists(WINfolder + "\SL\indexchi.tbr") Then Image1.Picture = LoadPicture(WINfolder + "\SL\indexchi.tbr")
    End If
    '--------
    AjustarFRM Me, 12000
    'VU1.Visible = HabilitarVUMetro
    If TypeVersion = "DEMO" Then
        lblDEMO = "Este espacio sera suyo cuando adquiera la version full de 3PM"
    Else
        lblDEMO = textoUsuario
    End If
    'cargar la cantidad de tapas que corresponda
    'SE CARGAN EN ini YA ES configurable
    'TapasMostradasH = 4: TapasMostradasV = 3
    
    
    '-----------------
    If TypeVersion = "SL" Then
        If FSO.FileExists(WINfolder + "\SL\txtIDX.tbr") Then
            Set TE = FSO.OpenTextFile(WINfolder + "\SL\txtIDX.tbr", ForReading, False)
            Dim NewT As String
            NewT = TE.ReadAll
            lblTBR = NewT
            TE.Close
        Else
            lblTBR = "Software desarrollado por tbrSoft www.tbrsoft.com - info@tbrsoft.com - avazquez@cpcipc.org."
        End If
    Else
        lblTBR = "Software desarrollado por tbrSoft www.tbrsoft.com - info@tbrsoft.com - avazquez@cpcipc.org."
    End If
    '-----------------
    VU1.Width = Screen.Width
    VU1.Left = 0: VU1.Top = 0
    VU1.Height = lblTemaSonando.Top - 25
    If HabilitarVUMetro Then
        'que entre en el control
        frDISCOS.Width = VU1.Width - (VU1.AnchoBarra * 2) - 50 'Screen.Width
        frDISCOS.Left = VU1.AnchoBarra + 25 '0
    Else
        
        frDISCOS.Left = 0 ' tapa a las barras que no se usan 'VU1.Left + VU1.Width
        frDISCOS.Width = VU1.Width ' Screen.Width - VU1.Width
    End If
    frDISCOS.Top = 0
    frDISCOS.Height = lblTemaSonando.Top
    'ver si hay que mostrar el touch
    MostrarTouch = LeerConfig("MostrarTouch", "0")
    If MostrarTouch = False Then
        Frame1.Visible = False 'frame del touch
        lblTemaSonando.Width = lblTemaSonando.Width + Frame1.Width
        LBLpORCtEMA.Width = LBLpORCtEMA.Width + Frame1.Width
        lstProximos.Width = lstProximos.Width + Frame1.Width
        lblTBR.Width = lblTBR.Width + Frame1.Width
        lblDEMO.Width = lblDEMO.Width + Frame1.Width
    End If
    'frDISCOS contiene los discos a mostrar
    'se debera calcualr el tamaño de cada discos asi como cantidad horizontal y vertical
    Dim AnchoTapaDisco As Long
    Dim AltoTapaDisco As Long
    'el alto de estos incluye tambien el lbldisco
    
    AnchoTapaDisco = (frDISCOS.Width * 0.98 / TapasMostradasH)
    AltoTapaDisco = (frDISCOS.Height * 0.97 / TapasMostradasV)
    'ver cual es mayor para no permitir mucha distorsion
    'lo que se ajuste se agranda del espacio entrediscos
    Dim EspacioEntreDiscosH As Long
    Dim EspacioEntreDiscosV As Long
    EspacioEntreDiscosV = 50: EspacioEntreDiscosH = 50
    If DistorcionarTapas = False Then
        Dim DIFF As Double
        DIFF = AnchoTapaDisco - AltoTapaDisco
        If DIFF > 0 Then
            'el ancho es mas que el alto
            AnchoTapaDisco = AltoTapaDisco
            EspacioEntreDiscosH = DIFF
        Else
            'el alto es mas que el ancho
            AltoTapaDisco = AnchoTapaDisco
            EspacioEntreDiscosV = -DIFF
        End If
    End If
    
    If MostrarRotulos Then
        TapaCD(0).Width = AnchoTapaDisco
        TapaCD(0).Height = AltoTapaDisco * 0.79 '80%disco, 20% lbldisco
        lblDisco(0).Height = AltoTapaDisco * 0.19 '80%disco, 20% lbldisco
        lblDisco(0).Width = AnchoTapaDisco
    Else
        TapaCD(0).Width = AnchoTapaDisco
        TapaCD(0).Height = AltoTapaDisco
        lblDisco(0).Visible = False
    End If
    
    'ver si los rotulos van arriba o abajo
    If RotulosArriba Then
        lblDisco(0).Left = 50
        lblDisco(0).Top = 50
        TapaCD(0).Left = 50
        If MostrarRotulos Then
            TapaCD(0).Top = lblDisco(0).Top + lblDisco(0).Height + 50
        Else
            TapaCD(0).Top = 50
        End If
    Else
        TapaCD(0).Left = 50
        TapaCD(0).Top = 50
        lblDisco(0).Left = 50
        lblDisco(0).Top = TapaCD(0).Top + TapaCD(0).Height + 50
    End If
    
    Dim CantDiscos As Long
    CantDiscos = TapasMostradasH * TapasMostradasV
    'cargar la cantidad de tapas correspondientes
    c = 0
    Do While c < CantDiscos - 1 'si la primera hoja incompleta se carga completa!!
        c = c + 1
        Load TapaCD(c)
        Load lblDisco(c)
        'ya toman el tamaño del original
        If c / TapasMostradasH = c \ TapasMostradasH Then
            'es una tapa al principio de linea
            If RotulosArriba Then
                lblDisco(c).Left = 50
                lblDisco(c).Top = TapaCD(c - TapasMostradasH).Top + TapaCD(c - TapasMostradasH).Height + EspacioEntreDiscosV
                TapaCD(c).Left = 50
                If MostrarRotulos Then
                    TapaCD(c).Top = lblDisco(c).Top + lblDisco(c).Height + 50
                Else
                    TapaCD(c).Top = TapaCD(c - TapasMostradasH).Top + TapaCD(c - TapasMostradasH).Height + 50
                End If
                TapaCD(c).Visible = True
                If MostrarRotulos Then lblDisco(c).Visible = True
            Else
                TapaCD(c).Left = 50
                If MostrarRotulos Then
                    TapaCD(c).Top = lblDisco(c - TapasMostradasH).Top + lblDisco(c - TapasMostradasH).Height + EspacioEntreDiscosV
                Else
                    TapaCD(c).Top = TapaCD(c - TapasMostradasH).Top + TapaCD(c - TapasMostradasH).Height + EspacioEntreDiscosV
                End If
                lblDisco(c).Left = 50
                lblDisco(c).Top = TapaCD(c).Top + TapaCD(c).Height + 50
                TapaCD(c).Visible = True
                If MostrarRotulos Then lblDisco(c).Visible = True
            End If
        Else
            'una tapa comun que se acomoda a la derecha de la anterior
            If RotulosArriba Then
                lblDisco(c).Left = lblDisco(c - 1).Left + AnchoTapaDisco + EspacioEntreDiscosH
                lblDisco(c).Top = lblDisco(c - 1).Top
                TapaCD(c).Left = lblDisco(c).Left
                TapaCD(c).Top = TapaCD(c - 1).Top
                TapaCD(c).Visible = True
            Else
                TapaCD(c).Left = TapaCD(c - 1).Left + AnchoTapaDisco + EspacioEntreDiscosH
                TapaCD(c).Top = TapaCD(c - 1).Top
                lblDisco(c).Left = TapaCD(c).Left
                lblDisco(c).Top = lblDisco(c - 1).Top
                TapaCD(c).Visible = True
            End If
            If MostrarRotulos Then lblDisco(c).Visible = True
        End If
    Loop
    
    OnOffCAPS vbKeyScrollLock, True
    lblV = "versión " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
    lblTiempoRestante = "FALTA: " + "00:00"
    'ocultar las etiquetas
    Me.AutoRedraw = AutoReDibuj
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    'ver cuantos creditos hay
    CREDITOS = Val(LeerArch1Linea(AP + "creditos.tbr"))
    If CREDITOS >= 10 Then
        lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
    Else
        lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
    End If
    'dejar cargado el mostrados de procesos
    'Load frmini
    'cargar las variables globales

    TEMA_REPRODUCIENDO = "Sin reproducción actual"
    TEMA_SIGUIENTE = "No hay proximo tema"
    TEMAS_EN_LISTA = 0
    
    'buscar discos = todas las carpetas en AP\discos\*.*
    'y meterlos en la matriz
    MATRIZ_DISCOS() = ObtenerDir(AP + "discos")
    
    Dim CarpActual As String
    Dim pathTema As String, DuracionTema As String, nombreTEMA As String
    'mostrar proceso
    ReDim Preserve MATRIZ_TOTAL(150, 30)
    
    'ret devuelve la cantidadd de discos cargados
    DiscosEnPagina = CargarDiscos(0, True)
    'inicializar la matriz_lista (lista de reproduccion
    
    ReDim MATRIZ_LISTA(0)
    lblTOTdiscos = "Discos: " + Trim(Str(UBound(MATRIZ_DISCOS)))
    
    'si quedaron temas pendientes cargarlos
    
    Select Case ReINI
        Case "LISTA" 'solo la lista despues del tema actual
            
            If FSO.FileExists(AP + "reini.tbr") Then
                Set TE = FSO.OpenTextFile(AP + "reini.tbr", ForReading, False)
                Dim TT As String 'cada tema
                Dim z As Integer 'contador de temas en lista anterior
                z = 1
                Do While Not TE.AtEndOfStream
                    TT = TE.ReadLine
                    ReDim Preserve MATRIZ_LISTA(z)
                    MATRIZ_LISTA(z) = TT
                    z = z + 1
                Loop
                TE.Close
            End If
            EMPEZAR_SIGUIENTE
        Case "NADA"
            'no hacer nada
            'borrar la lista
            If FSO.FileExists(AP + "reini.tbr") Then FSO.DeleteFile AP + "reini.tbr", True
            Timer1.Interval = 10000
    End Select
    Unload frmINI
    
    'ver si hay validacion por creditos
    Validar = LeerConfig("Validar", "0")
    If Validar Then
        'ver si existe el archivo Creditos Validar
        If FSO.FileExists(SYSfolder + "\radilav.cfg") Then
            'leer el archivo de creditos vaildados
            CreditosValidar = CLng(LeerArch1Linea(SYSfolder + "\radilav.cfg"))
            'CodigoParaClaveActual busca el archivo con el numero que corresponde validar en este periodo de control
        Else
            EscribirArch1Linea SYSfolder + "\radilav.cfg", "0"
            CreditosValidar = 0
            CrearNuevoCodigoValidar 'graba el archivo con un numero al azar
            'lo mantiene hasta que se genera uno nuevo al terminar el periodo de control
        End If
        'ver cual es el máximo y si hay que avisar
        ValidarCada = LeerConfig("ValidarCada", "500")
        AvisarAntes = LeerConfig("AvisarAntes", "50")
        If CreditosValidar > ValidarCada - AvisarAntes Then
            'solicitar una clave
            'se podra saltear solo si todavia no llego al limite
            
            'uso el frmClave que tiene la variable publica ClaveIngresada
            Dim ClaveCorrespondiente As String
            ClaveCorrespondiente = ClaveParaValidar(CodigoParaClaveActual)
            Dim QuedanC As Long
            QuedanC = ValidarCada - CreditosValidar
            If QuedanC > 0 Then
                'CodigoParaClaveActual busca el archivo con el numero que corresponde validar en este periodo de control
                MsgBox "Ingrese a continuacion su clave para continuar utilizando 3PM. " + vbCrLf + _
                    "Debe enviar la administrador el codigo: " + vbCrLf + _
                    CodigoParaClaveActual + vbCrLf + _
                    "Puede todavia omitir esta clave. Solo le quedan " + CStr(QuedanC) + " creditos hasta que 3PM se inhabilite"
            Else
                MsgBox "De no ingresar la clave correspondiente 3PM no podra continuar. Ha llegado al limite de creditos posibles"
            End If
            
            frmCLAVE.Show 1
            If UCase(ClaveIngresada) <> UCase(ClaveCorrespondiente) Then
                If QuedanC > 0 Then
                    MsgBox "La clave es erronea!" + vbCrLf + _
                        "Le quedan " + CStr(QuedanC) + " creditos por cargar antes que se inhabilite 3PM"
                Else
                    MsgBox "No podra seguir utilizando 3PM hasta que valide con la clave correspondiente"
                    End
                End If
            Else
                'todo OK. Cargo bien la clave
                CreditosValidar = 0
                EscribirArch1Linea SYSfolder + "\radilav.cfg", "0"
                'empezar un nuevo periodo
                CrearNuevoCodigoValidar 'graba el archivo con un numero al azar
            End If
        End If
        lblValidar = "Val=" + CStr(ValidarCada) + "-Qued=" + CStr(ValidarCada - CreditosValidar) + "Actual=" + CStr(CreditosValidar) + " Codigo: " + CodigoParaClaveActual
    
    End If
    On Error GoTo NoPort1
    If COMM1.PortOpen = False Then COMM1.PortOpen = True
    On Error GoTo 0
    Exit Sub
NoPort1:
    
    On Error GoTo 0
    Exit Sub
ErrMP3:
    MsgBox Err.Description + " N°: " + Str(Err.Number)
End Sub

Public Sub SelDisco(nDisco As Long)
    lblSel.Visible = False
    lblDisco(nDisco).ForeColor = vbBlack
    'lblDISCO(nDisco).Font.Bold = True
    lblDisco(nDisco).Font.Underline = True
    lblDisco(nDisco).BackColor = vbYellow
    nDiscoSEL = nDisco
    
    lblSel.Top = TapaCD(nDiscoSEL).Top - lblSel.BorderWidth * 10
    lblSel.Left = TapaCD(nDiscoSEL).Left - lblSel.BorderWidth * 10
    lblSel.Height = TapaCD(nDiscoSEL).Height + lblSel.BorderWidth * 20
    lblSel.Width = TapaCD(nDiscoSEL).Width + lblSel.BorderWidth * 20
    lblSel.Visible = True
    lblSel.ZOrder
    lblDisco(nDisco).ZOrder
    
    'seleccionar de la lista de solo video
    L(nDiscoGral).ForeColor = vbWhite
    L(nDiscoGral).BackColor = vbBlack
    LastDiscoSel = nDiscoGral 'para saber cual desactivar en unsel
    TapaCD(nDiscoGral).BorderStyle = 1
    'imgARROW.Top = TapaCD(nDiscoGral).Top + TapaCD(nDiscoGral).Height - imgARROW.Height
    'imgArrow2.Top = TapaCD(nDiscoGral).Top + TapaCD(nDiscoGral).Height - imgArrow2.Height
    'imgARROW.Left = TapaCD(nDiscoGral).Left
    'imgArrow2.Left = TapaCD(nDiscoGral).Left + TapaCD(nDiscoGral).Width - imgARROW.Width
    'imgARROW.ZOrder
    'imgArrow2.ZOrder
    If EsVideo Then OrdenarListaModoVideo
    
End Sub

Public Sub UnSelDisco(nDisco As Long)
    lblDisco(nDisco).ForeColor = vbWhite
    'lblDISCO(nDisco).Font.Bold = False
    lblDisco(nDisco).Font.Underline = False
    lblDisco(nDisco).BackColor = vbBlack
    'seleccionar de la lista de solo video
    L(LastDiscoSel).ForeColor = vbBlack
    L(LastDiscoSel).BackColor = vbWhite
    TapaCD(LastDiscoSel).BorderStyle = 0
    
    If EsVideo Then OrdenarListaModoVideo
End Sub

Public Function CargarDiscos(numDiscoIniciar As Long, SelPrimero As Boolean) As Long
    'indicando en que disco se inicia carga ese y los seis (o lo que corresponde) que le sigen
    'devuelve el número de discos cargados
    CargarDiscos = 0
    Dim TotPags As Long
    TotPags = (TOTAL_DISCOS - 1) \ (TapasMostradasH * TapasMostradasV)
    lblPag = "Pagina " + CStr(Round(numDiscoIniciar / (TapasMostradasH * TapasMostradasV) + 1, 0)) + " de " + CStr(TotPags + 1)
    'tomar el disco que va a quedar seleccionado como numero de disco en el indice general
    If SelPrimero Then
        nDiscoGral = numDiscoIniciar
    Else
        nDiscoGral = numDiscoIniciar + ((TapasMostradasH * TapasMostradasV) - 1) 'era un 5, o sea total tapas-1
    End If
    'esconder todos los discos
    Dim NDR As Long 'numero de tapa de disco real del 0 al 5 (total de discos-1)
    
    'no hacer esto al pedo si ya estan cargadas
    Dim NDI As Long '=numdiscoiniciar de la pagina
    Dim c As Integer
    c = 1
    NDI = numDiscoIniciar
    If CargarIMGinicio Then
        If SelPrimero Then
            'si voy para adelante
            'ocultar los que ya pse
            c = 1
            Do While c <= (TapasMostradasH * TapasMostradasV)
                If NDI >= (TapasMostradasH * TapasMostradasV) Then
                    TapaCD(NDI - c).Visible = False
                    'no se cargan lbldisco, usan solo del 0 al 5
                    lblDisco(c - 1).Visible = False
                End If
                c = c + 1
            Loop
            Me.Refresh
        Else
            'sino ocultar los de adelante
            c = 1
            Do While c <= (TapasMostradasH * TapasMostradasV)
                If NDI + ((TapasMostradasH * TapasMostradasV) - 1) + c < UBound(MATRIZ_DISCOS) Then TapaCD(NDI + ((TapasMostradasH * TapasMostradasV) - 1) + c).Visible = False
                lblDisco(c - 1).Visible = False
                c = c + 1
            Loop
            'Me.Refresh
        End If
    Else
        Do While NDR < ((TapasMostradasH * TapasMostradasV))
            TapaCD(NDR).Visible = False
            lblDisco(NDR).Visible = False
            NDR = NDR + 1
        Loop
        Dim ArchTapa As String
    End If
    NDR = 0
    
    Do While NDI < numDiscoIniciar + ((TapasMostradasH * TapasMostradasV))
        'ver si existe si hay disco con este n°
        If NDI < UBound(MATRIZ_DISCOS) Then
            CargarDiscos = CargarDiscos + 1
            'ver si ya estan cargadas o se deben cargar
            If CargarIMGinicio Then
                 TapaCD(NDI).Visible = True
                TapaCD(NDI).ZOrder
            Else
                'ver si hay tapa
                ArchTapa = txtInLista(MATRIZ_DISCOS(NDI + 1), 0, ",")
                If Right(ArchTapa, 1) <> "\" Then ArchTapa = ArchTapa + "\"
                ArchTapa = ArchTapa + "tapa.jpg"
                If FSO.FileExists(ArchTapa) Then
                    TapaCD(NDR).Picture = LoadPicture(ArchTapa)
                Else
                    TapaCD(NDR).Picture = LoadPicture(AP + "tapa.jpg")
                End If
                TapaCD(NDR).Visible = True
            End If
            'poner nombre al disco
            lblDisco(NDR) = txtInLista(MATRIZ_DISCOS(NDI + 1), 1, ",")
            If MostrarRotulos Then lblDisco(NDR).Visible = True
        End If
        NDI = NDI + 1
        NDR = NDR + 1
    Loop
    If SelPrimero Then
        UnSelDisco ((TapasMostradasH * TapasMostradasV) - 1)
        SelDisco 0
    Else
        UnSelDisco 0
        SelDisco ((TapasMostradasH * TapasMostradasV) - 1)
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MostrarCursor True
    COMM1.PortOpen = False
    MP3.DoStop
    MP3.DoClose
    VU1.DoStop
End Sub

Private Sub MP3_BeginPlay()
    TotalTema = MP3.LengthInSec
    Ancho = lblTemaSonando.Width
    'EVITAR DIVISIONES POR CERO
    If TotalTema > 0 And MP3.IsPlaying Then
        Variacion = Ancho / TotalTema
        lblTiempoRestante = "TOTAL: " + MP3.Falta
    Else
        lblTiempoRestante = "FALTA: " + "00:00"
    End If
    VolBajando = MP3.Volumen
    
End Sub

Private Sub MP3_EndPlay()
    'volver a PasarHoja a su estado original
    PasarHoja = LeerConfig("PasarHoja", "1")
    VU1.Width = Screen.Width
    If HabilitarVUMetro Then
        frDISCOS.Width = VU1.Width - (VU1.AnchoBarra * 2) - 50 'Screen.Width - VU1.Width
        frDISCOS.Left = VU1.AnchoBarra + 25 ' VU1.Width
        'vu no se mueve si termina un video        'VU1.Top = 0        'VU1.Height = Me.Height
    Else
        frDISCOS.Width = VU1.Width ' Screen.Width
        frDISCOS.Left = 0
    End If
    frModoVideo.Visible = False
    lblModoVideo.Visible = False
    frTEMAS.Visible = False
    lblTEMAS.Visible = False
    ModoVideoSelTema = False
    LBLpORCtEMA.Width = Ancho
    'termino una cancion
    If EsVideo Then MP3.DoClose
    picVideo.Visible = False
    EMPEZAR_SIGUIENTE
End Sub

Private Sub MP3_Played(SecondsPlayed As Long)
    'esto pasa cada un segundo (si o si una vez por segundo
    Dim sRest As Long
    sRest = MP3.FaltaInSec
    PorcEjecutado = MP3.PercentPlay
    If PorcEjecutado > PorcentajeTEMA And CORTAR_TEMA Then
        VolBajando = VolBajando - 5 'baja 1 por segundo
        lblTemaSonando = "Cerrando " + QuitarNumeroDeTema(FSO.GetBaseName(TEMA_REPRODUCIENDO))
        If VolBajando > 0 Then
            MP3.Volumen = VolBajando
        Else
            MP3.DoStop
            'EL DOSTOP DESENCADENA UN END PLAY QUE REALIZA UN EMPEZAR SIGUINETE
            'EMPEZAR_SIGUIENTE
        End If
    End If
    lblTiempoRestante = "FALTA: " + MP3.Falta
    wi = Ancho - Variacion * (SecondsPlayed - 2)
    If wi > 0 Then LBLpORCtEMA.Width = wi
    '=====================================
    If TypeVersion = "DEMO" And SecondsPlayed > 126 And SecondsPlayed < TotalTema - 5 Then
        lblTemaSonando = "Tema Truncado. Version DEMO"
        MP3.DoStop
    End If
    'cotar tambin en el gratuito
    If TypeVersion = "DEMO2" And SecondsPlayed > 126 And SecondsPlayed < TotalTema - 5 Then
        lblTemaSonando = "Tema Truncado. Version DEMO"
        MP3.DoStop
    End If
    '=====================================
End Sub

Private Sub MSComm1_OnComm()

End Sub

Private Sub TapaCD_Click(Index As Integer)
    'nunca hay que pasar hojas
    'nDiscoGral = nDiscoGral + (Index - nDiscoSEL)
    nDiscoGral = Index 'si se cargan todas las imágenes al inicio index=nDiscoGral
    
    If nDiscoGral + 1 >= TOTAL_DISCOS Then
        MsgBox "No existe el disco elegido!!. " + vbCrLf + _
            "Carge discos desde el ADMINISTRADOR DE DISCOS en la " + vbCrLf + _
            "página de configuracion (presionando la tecla 'C')"
        Exit Sub
    End If
    
    UnSelDisco nDiscoSEL
    
    Dim PagNum As Long
    PagNum = nDiscoGral \ (TapasMostradasH * TapasMostradasV)
    nDiscoSEL = Index - (PagNum * (TapasMostradasH * TapasMostradasV))
    
    SelDisco nDiscoSEL
    
    lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
    
    'totar la tecla de enar a disco
    Form_KeyDown TeclaOK, 0
End Sub

Private Sub Timer1_Timer()
    If MP3.IsPlaying Then Exit Sub
    'controla el tiempo sin uso (sin ejecucion de temas)
    SecSinUso = SecSinUso + 10
    lblNoUSO = Trim(Str(SecSinUso))
    If SecSinUso >= EsperaMinutos Then 'esperaminutos esta en segundos
                
        SecSinUso = 0
        Dim TemasDisponibles As Long
        If TemasEnRank(1) > 50 Then
            TemasDisponibles = TemasEnRank(1) 'todos los que se escucharon
        Else
            TemasDisponibles = TemasEnRank(0) 'todos los que se escucharon
        End If
        Randomize Timer
        z = Int(Rnd * TemasDisponibles)
        z = z + 1
        CC = 0
        If FSO.FileExists(AP + "ranking.tbr") = False Then
            FSO.CreateTextFile AP + "ranking.tbr", True
            'me voy al azar ya que no hay para elegirdel rank
            GoTo AZAR
        End If
        
        Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
        Dim TT As String
        'antes de entra ver si el archivo no tiene nada
        If TE.AtEndOfStream Then GoTo AZAR
        Do While Not TE.AtEndOfStream
            CC = CC + 1
            TT = TE.ReadLine
            If CC = z Then
                Dim TemaAzar As String
                TemaAzar = txtInLista(TT, 1, ",")
                'si tuve los discos cargados en una unidad o una ubicación distinta a la que aparece
                'en el ranking, me da un error por que el archivo no existe
                If FSO.FileExists(TemaAzar) Then
                    CORTAR_TEMA = True 'este tema se eligio al azar no va entero
                    SecSinUso = 0
                    EjecutarTema TemaAzar, False
                    Exit Sub
                Else
AZAR:
                    'ejecutar algun tema de cualquier disco
                    Dim MTX10() As String: zz = 0
                    ruta = AP + "discos\"
                    Dim NombreDir As String
                    NombreDir = Dir$(ruta & "*.*", vbDirectory)
                    Do While Len(NombreDir)
                        If NombreDir = "." Or NombreDir = ".." Then
                            ' excluir las entradas "." y ".."
                        ElseIf (GetAttr(ruta & NombreDir) And vbDirectory) = 0 Then
                            ' este es un archivo normal
                        Else
                            'ver los primeros diez discos. En alguno tiene que haber temas
                            'yo se que el primero no tiene temas por que es
                            '01 - los mas escuchados
                            ReDim Preserve MTX10(zz) As String
                            MTX10(zz) = ruta & NombreDir
                            zz = zz + 1
                        End If
                        NombreDir = Dir$
                    Loop
BuscaMP3:
                    'siempre cae en el primer tema del primer directorio habilitado
                    Randomize Timer
                    Dim A As Integer, ContA As Integer
                    A = Int(Rnd * 1000) + 1
                    Dim NombreMP3 As String: zz = 0
                    Dim temaMP As String
                    Do While zz < UBound(MTX10)
                        NombreMP3 = Dir$(MTX10(zz) & "\*.mp3")
                        'si no hay ningun tema se va a la prox carpeta
                        If NombreMP3 = "" Then GoTo NextFolder
                        'da vueltas hasta encontrar un tema valido
                        Do While Len(NombreMP3)
                            temaMP = MTX10(zz) & "\" & NombreMP3
                            If FSO.FileExists(temaMP) Then
                                ContA = ContA + 1
                                If ContA >= A Then
                                    CORTAR_TEMA = True 'este tema va cortado ya que es de 3PM para que haga ruido
                                    EjecutarTema temaMP, False
                                    'solo sale cueando encuentra un tema valido
                                    SecSinUso = 0
                                    Exit Sub
                                End If
                            End If
                            NombreMP3 = Dir$
                        Loop
NextFolder:
                        zz = zz + 1
                    Loop
                End If
                Exit Do
            End If
         Loop
        'si llego aca es por que no encontro el numero sorteado al azar en la lista
        'de los mejores. Entonces elige un tema al azar
        GoTo AZAR
    End If
    
End Sub

Private Sub Timer3_Timer()
    'controla el tiempo sin uso (sin tocar teclas)
    SecSinTecla = SecSinTecla + 10
    lblNoTecla = Trim(Str(SecSinTecla))
    'no protector en video
    If EsVideo Then SecSinTecla = 0
    If SecSinTecla > EsperaTecla And EsVideo = False Then
        frmProtect.Show 1
    End If
End Sub

Public Function TemasEnRank(MasDeXVotos) As Long
    'indica cuantos temas hay en el ranking
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        FSO.CreateTextFile AP + "ranking.tbr", True
        TemasEnRankMasDeUnVoto = 0
        Exit Function
    End If
    
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    Dim TT As String
    'antes de entra ver si el archivo no tiene nada
    If TE.AtEndOfStream Then
        TemasEnRankMasDeUnVoto = 0
        Exit Function
    End If
    Dim CA As Long
    CA = 0
    Dim PuntosEste  As Long
    Do While Not TE.AtEndOfStream
        TT = TE.ReadLine
        PuntosEste = Val(txtInLista(TT, 0, ","))
        If PuntosEste > MasDeXVotos Then
            CA = CA + 1
        Else
            'todos los que siguen tienen uno (1)
            Exit Do
        End If
    Loop
    TemasEnRank = CA
End Function

Public Sub OrdenarListaModoVideo()
    'asegurarme que el disco elegido se ve en la lista
    Dim CL As Long 'contador de L
    Dim HayQueCorrerse As Long 'cuanto hay que correrse
    'para acomodar
    If L(nDiscoGral).Top > frModoVideo.Height - (L(0).Height + 25) Then
        'esta fuera de la vista para abajo
        'correr todo para abajo
        'ver cuanto hay que corresse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        HayQueCorrerse = L(nDiscoGral).Top - (frModoVideo.Height - (L(0).Height + 25))
        
        CL = 0
        Do While CL < TOTAL_DISCOS
            L(CL).Top = L(CL).Top - HayQueCorrerse
            CL = CL + 1
        Loop
    End If
    
    If L(nDiscoGral).Top < 0 Then
        'ver cuanto hay que corresse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        HayQueCorrerse = -L(nDiscoGral).Top
        
        'esta fuera de la vista para arriba
        'correr todo para arriba
        CL = 0
        Do While CL < TOTAL_DISCOS
            L(CL).Top = L(CL).Top + HayQueCorrerse
            CL = CL + 1
        Loop
    End If
    
End Sub

Public Sub SelTema(n As Integer)
    T(n).BackColor = &H0&
    T(n).ForeColor = &H80FFFF
End Sub

Public Sub UnSelTema(n As Integer)
    T(n).BackColor = &H80FFFF
    T(n).ForeColor = &H0&
End Sub

Public Sub OrdenarListaTemaVideo()
    'asegurarme que el disco elegido se ve en la lista
    Dim CL As Long 'contador de L
    Dim HayQueCorrerse As Long 'cuanto hay que correrse
    'para acomodar
    If T(TemaElegidoModoVideo).Top > frTEMAS.Height - (T(0).Height + 25) Then
        'esta fuera de la vista para abajo
        'correr todo para abajo
        'ver cuanto hay que correrse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        HayQueCorrerse = T(TemaElegidoModoVideo).Top - (frTEMAS.Height - (T(0).Height + 25))
        
        CL = 0
        Do While CL <= UBound(MATRIZ_TEMAS)
            T(CL).Top = T(CL).Top - HayQueCorrerse
            CL = CL + 1
        Loop
    End If
    
    If T(TemaElegidoModoVideo).Top < 0 Then
        'ver cuanto hay que corresse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        HayQueCorrerse = -T(TemaElegidoModoVideo).Top
        
        'esta fuera de la vista para arriba
        'correr todo para arriba
        CL = 0
        Do While CL <= UBound(MATRIZ_TEMAS)
            T(CL).Top = T(CL).Top + HayQueCorrerse
            CL = CL + 1
        Loop
    End If
    
End Sub
