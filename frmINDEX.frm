VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmINDEX 
   BackColor       =   &H00000040&
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
   Begin VB.Timer Timer1 
      Left            =   7320
      Top             =   3180
   End
   Begin MSComctlLib.Slider SLvolumen 
      Height          =   240
      Left            =   10770
      TabIndex        =   13
      Top             =   60
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   423
      _Version        =   393216
      Min             =   -10000
      Max             =   0
      SelStart        =   -1900
      Value           =   -1900
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tema actual"
      BeginProperty Font 
         Name            =   "HandelGotDLig"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   60
      TabIndex        =   16
      Top             =   7590
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proximos temas elegidos"
      BeginProperty Font 
         Name            =   "HandelGotDLig"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   1095
      Left            =   60
      TabIndex        =   15
      Top             =   7890
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "3PM"
      BeginProperty Font 
         Name            =   "HandelGotDLig"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   400
      Left            =   10200
      TabIndex        =   14
      Top             =   8610
      Width           =   1800
   End
   Begin VB.Image FlechaCD 
      Height          =   480
      Index           =   5
      Left            =   7920
      Picture         =   "frmINDEX.frx":0000
      Top             =   6510
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image FlechaCD 
      Height          =   480
      Index           =   4
      Left            =   7890
      Picture         =   "frmINDEX.frx":0442
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image FlechaCD 
      Height          =   480
      Index           =   3
      Left            =   3900
      Picture         =   "frmINDEX.frx":0884
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image FlechaCD 
      Height          =   480
      Index           =   2
      Left            =   3930
      Picture         =   "frmINDEX.frx":0CC6
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image FlechaCD 
      Height          =   480
      Index           =   1
      Left            =   -30
      Picture         =   "frmINDEX.frx":1108
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image FlechaCD 
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "frmINDEX.frx":154A
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblDisco 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   5
      Left            =   8130
      TabIndex        =   12
      Top             =   6990
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.Label lblDisco 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   4
      Left            =   8130
      TabIndex        =   11
      Top             =   3330
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.Label lblDisco 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   4170
      TabIndex        =   10
      Top             =   6990
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.Label lblDisco 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   2
      Left            =   4170
      TabIndex        =   9
      Top             =   3330
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.Label lblDisco 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   210
      TabIndex        =   8
      Top             =   6990
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.Label lblDisco 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   0
      Left            =   210
      TabIndex        =   7
      Top             =   3300
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.Image TapaCD 
      BorderStyle     =   1  'Fixed Single
      Height          =   3300
      Index           =   5
      Left            =   8370
      Stretch         =   -1  'True
      Top             =   3660
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Image TapaCD 
      BorderStyle     =   1  'Fixed Single
      Height          =   3300
      Index           =   4
      Left            =   8340
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Image TapaCD 
      BorderStyle     =   1  'Fixed Single
      Height          =   3300
      Index           =   3
      Left            =   4380
      Stretch         =   -1  'True
      Top             =   3660
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Image TapaCD 
      BorderStyle     =   1  'Fixed Single
      Height          =   3300
      Index           =   2
      Left            =   4380
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Image TapaCD 
      BorderStyle     =   1  'Fixed Single
      Height          =   3300
      Index           =   1
      Left            =   420
      Stretch         =   -1  'True
      Top             =   3660
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Image TapaCD 
      BorderStyle     =   1  'Fixed Single
      Height          =   3300
      Index           =   0
      Left            =   420
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Label lblTemaSonando 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sin Reproducción actual"
      BeginProperty Font 
         Name            =   "HandelGotDLig"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   1290
      TabIndex        =   0
      Top             =   7590
      Width           =   8925
   End
   Begin VB.Label lblIndicaciones 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Utilize las flechas para desplazarse sobre los distintos discos, para conocer el detalle de cada disco utilice en boton OK"
      BeginProperty Font 
         Name            =   "HandelGotDLig"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   6
      Top             =   7350
      Width           =   11955
   End
   Begin VB.Label lblTemasPorFicha 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Temas/ficha: 01"
      BeginProperty Font 
         Name            =   "HandelGotDLig"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   250
      Left            =   10200
      TabIndex        =   5
      Top             =   8340
      Width           =   1800
   End
   Begin VB.Label lblCreditos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Creditos: 00"
      BeginProperty Font 
         Name            =   "HandelGotDLig"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   250
      Left            =   10200
      TabIndex        =   3
      Top             =   8100
      Width           =   1800
   End
   Begin VB.Label lblTemasEnLista 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "En lista: 00"
      BeginProperty Font 
         Name            =   "HandelGotDLig"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   250
      Left            =   10200
      TabIndex        =   2
      Top             =   7830
      Width           =   1800
   End
   Begin VB.Label lblTiempoRestante 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Restante: 00:00"
      BeginProperty Font 
         Name            =   "HandelGotDLig"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   250
      Left            =   10200
      TabIndex        =   1
      Top             =   7590
      Width           =   1800
   End
   Begin VB.Label lblProximoTema 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No hay próximo tema"
      BeginProperty Font 
         Name            =   "HandelGotDLig"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   1095
      Left            =   1260
      TabIndex        =   4
      Top             =   7890
      Width           =   8940
   End
End
Attribute VB_Name = "frmINDEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nDiscoSEL As Long 'del 0 al 5
Dim nDiscoGral As Long ' del 0 a total_discos
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'aqui se regsitran las presiones de las teclas elegidas
    Select Case KeyCode
        Case vbKeyZ
            If ESTOY = 0 Then 'si estoy en los discos
                'no ir a -1
                If nDiscoSEL = 0 Then
                    If nDiscoGral > 0 Then CargarDiscos nDiscoGral - 6, False
                Else
                    nDiscoGral = nDiscoGral - 1
                    UnSelDisco nDiscoSEL
                    SelDisco nDiscoSEL - 1
                End If
            End If
        Case vbKeyX
            If ESTOY = 0 Then 'si estoy en los discos
                If nDiscoSEL = 5 Then
                    If nDiscoGral + 1 < TOTAL_DISCOS Then CargarDiscos nDiscoGral + 1, True
                Else
                    If nDiscoGral + 1 < TOTAL_DISCOS Then
                        nDiscoGral = nDiscoGral + 1
                        UnSelDisco nDiscoSEL
                        SelDisco nDiscoSEL + 1
                    End If
                End If
            End If
        Case vbKeyReturn
            If ESTOY = 0 Then
                'si estoy mostrando discos debo mostrar temas
                'se cargan los temas en una matriz con ubic archivo,nombreTema
                ReDim MATRIZ_TEMAS(0) 'matriz en blanco
                'es una matriz global
                UbicDiscoActual = txtInLista(MATRIZ_DISCOS(nDiscoGral + 1), 0, ",")
                MATRIZ_TEMAS = ObtenerArchivos(UbicDiscoActual, "*.mp3")
                ESTOY = 1 'estoy dentro de un disco
                Dim c As Integer, nombreTemas As String
                Dim pathTema As String, DuracionTema As String
                Do While c < UBound(MATRIZ_TEMAS)
                    nombreTemas = txtInLista(MATRIZ_TEMAS(c + 1), 1, ",")
                    pathTema = txtInLista(MATRIZ_TEMAS(c + 1), 0, ",")
                    ''no mostrar duracion
                    'ver cuanto dura
                    'Dim archMP3 As cPlayWMP
                    'Set archMP3 = New cPlayWMP
                    'archMP3.FileName = PathTema
                    'DuracionTema = archMP3.TiempoTotal
                    'quitar el molesto ".mp3"
                    nombreTemas = Left(nombreTemas, Len(nombreTemas) - 4)
                    frmTemasDeDisco.lstTemas.AddItem nombreTemas '+ " / " + DuracionTema
                    c = c + 1
                Loop
                'ver si hay elementos en la lista
                If frmTemasDeDisco.lstTemas.ListCount = 0 Then
                    MsgBox "No hay Temas"
                    Unload frmTemasDeDisco
                    ESTOY = 0
                    Exit Sub
                End If
                frmTemasDeDisco.lstTemas.ListIndex = 0
                frmTemasDeDisco.Show 1
            End If
        Case vbKeyEscape
            If ESTOY = 0 Then If MsgBox("salir?", vbYesNo) = vbYes Then End
            ''nunca va a detectar en estoy=1 ya que es otro formulario el que recibe la tecla
            'If ESTOY = 1 Then
            '    Unload frmTemasDeDisco
            '    ESTOY = 0
            'End If
    End Select
    'lblProximoTema = "nDiscoSel=" + Str(nDiscoSEL) + " - nDiscoGral=" + Str(nDiscoGral)
End Sub

Private Sub Form_Load()
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    'dejar cargado el mostrados de procesos
    Load frmProces
    'cargar las variables globales
    ESTOY = 0 'aparece viendo los CDS
    CREDITOS = 0
    TEMAS = 0
    TEMA_REPRODUCIENDO = "Sin reproducción actual"
    TEMA_SIGUIENTE = "No hay proximo tema"
    TEMAS_EN_LISTA = 0
    ESTOY_REPRODUCIENDO = False
    'buscar discos = todas las carpetas en AP\discos\*.*
    'y meterlos en la matriz
    MATRIZ_DISCOS() = ObtenerDir(AP + "discos")
    'leer los temas de cada carpeta de la matriz anterior
    'y generar una nueva matriz con path, duracion
    ''borrar matriz_total
    'ReDim MATRIZ_TOTAL(0, 0)
    Dim CarpActual As String
    Dim MP3 As cPlayWMP, pathTema As String, DuracionTema As String, NombreTema As String
    Set MP3 = New cPlayWMP
    'mostrar proceso
    frmProces.Show
    frmProces.pBar = 0
    frmProces.pBar.Max = UBound(MATRIZ_DISCOS) * 15 'mas o menos 15 temas por disco
    ReDim Preserve MATRIZ_TOTAL(150, 30)
    On Error GoTo ErrMP3
    For c = 1 To UBound(MATRIZ_DISCOS)
        'encontar todos los temas grabar su duracion
        CarpActual = txtInLista(MATRIZ_DISCOS(c + 1), 0, ",")
        frmProces.lblProces = "Buscando en Disco " + CarpActual
        frmProces.lblProces.Refresh
        MATRIZ_TEMAS = ObtenerArchivos(CarpActual, "*.mp3")
        For d = 1 To UBound(MATRIZ_TEMAS)
            pathTema = txtInLista(MATRIZ_TEMAS(d), 0, ",")
            NombreTema = txtInLista(MATRIZ_TEMAS(d), 1, ",")
            
            MP3.FileName = pathTema
            DuracionTema = MP3.TiempoTotal
            
            MATRIZ_TOTAL(c, d) = CarpActual + "," + NombreTema + "," + DuracionTema
            frmProces.lblProces = "Tema encontrado " + NombreTema + " = " + DuracionTema
            frmProces.lblProces.Refresh
            If frmProces.pBar + 1 = frmProces.pBar.Max Then frmProces.pBar.Max = frmProces.pBar.Max + 1
            frmProces.pBar = frmProces.pBar + 1
        Next
    Next
    'ahora cargarlos en pantalla
    'ret devuelve la cantidadd de discos cargados
    Ret = CargarDiscos(0, True)
    'inicializar la matriz_lista (lista de reproduccion
    
    ReDim MATRIZ_LISTA(0)
    Exit Sub
ErrMP3:
    MsgBox Err.Description + " N°: " + Str(Err.Number)
End Sub

Public Sub SelDisco(nDisco As Long)
    FlechaCD(nDisco).Visible = True
    lblDisco(nDisco).ForeColor = vbYellow
    lblDisco(nDisco).Font.Bold = True
    lblDisco(nDisco).Font.Underline = True
    nDiscoSEL = nDisco
    
End Sub

Public Sub UnSelDisco(nDisco As Long)
    FlechaCD(nDisco).Visible = False
    lblDisco(nDisco).ForeColor = vbWhite
    lblDisco(nDisco).Font.Bold = False
    lblDisco(nDisco).Font.Underline = False
End Sub


Public Function CargarDiscos(numDiscoIniciar As Long, SelPrimero As Boolean) As Long
    'indicando en que disco se inicia carga ese y los seis que le sigen
    'devuelve el número de discos cargados
    CargarDiscos = 0
    'tomar el disco que va a quedar seleccionado como numero de disoc en el indice general
    If SelPrimero Then
        nDiscoGral = numDiscoIniciar
    Else
        nDiscoGral = numDiscoIniciar + 5
    End If
    'esconder todos los discos
    Dim NDR 'numero de tapa de disco real del 0 al 5
    NDR = 0
    Do While NDR < 6
        TapaCD(NDR).Visible = False
        lblDisco(NDR).Visible = False
        NDR = NDR + 1
    Loop
    NDR = 0
    Dim NDI '=numdiscoiniciar
    
    NDR = 0
    NDI = numDiscoIniciar
    ''si llegue al final empiezo de vuelta
    'If NDI >= UBound(MATRIZ_DISCOS) Then NDI = 0
    
    Do While NDI < numDiscoIniciar + 6
        'ver si existe si hay disco con este n°
        If NDI < UBound(MATRIZ_DISCOS) Then
            'ver si hay tapa
            Dim ArchTapa As String
            ArchTapa = txtInLista(MATRIZ_DISCOS(NDI + 1), 0, ",")
            If Right(ArchTapa, 1) <> "\" Then ArchTapa = ArchTapa + "\"
            ArchTapa = ArchTapa + "tapa.jpg"
            If FSO.FileExists(ArchTapa) Then
                TapaCD(NDR).Picture = LoadPicture(ArchTapa)
            Else
                TapaCD(NDR).Picture = LoadPicture(AP + "tapa.jpg")
            End If
            TapaCD(NDR).Visible = True
            lblDisco(NDR) = txtInLista(MATRIZ_DISCOS(NDI + 1), 1, ",")
            lblDisco(NDR).Visible = True
        End If
        NDI = NDI + 1
        NDR = NDR + 1
    Loop
    If SelPrimero Then
        UnSelDisco 5
        SelDisco 0
    Else
        UnSelDisco 0
        SelDisco 5
    End If
    
End Function

Private Sub SLvolumen_Change()
    m_csplay.Volumen = SLvolumen
End Sub

Private Sub Timer1_Timer()
    WAIT_EMPIEZA = WAIT_EMPIEZA - 1
    
    Dim m As Long, s As Long, tt As String
    Dim sRest As Long
    'sRest = m_csplay.Duration - m_csplay.CurrentPosition
    sRest = m_csplay.SegundosRestantes
    m = sRest \ 60
    s = sRest - (m * 60)
    'corregir 2:5 por 2:05
    If s < 10 Then
        lblTiempoRestante = "Restante " + Str(m) + ":0" + Trim(Str(s))
    Else
        lblTiempoRestante = "Restante " + Str(m) + ":" + Trim(Str(s))
    End If
    
    If m = 0 And s <= 1 And WAIT_EMPIEZA < 0 Then
        Timer1.Interval = 0
        lblTiempoRestante = "Restante 0:00"
        TEMA_REPRODUCIENDO = "Sin reproduccion actual"
        ESTOY_REPRODUCIENDO = False
        'si hay algun elemento en la lista ejecutarlo
        If UBound(MATRIZ_LISTA) > 0 Then
            Dim TemaDeMatriz As String
            TemaDeMatriz = txtInLista(MATRIZ_LISTA(1), 1, ",")
            'reacomodar la matriz para quitar el primer elemento
            For c = 1 To UBound(MATRIZ_LISTA)
                
                If c < UBound(MATRIZ_LISTA) Then
                    'cuando sea cualquiera menos el ultimo
                    MATRIZ_LISTA(c) = MATRIZ_LISTA(c + 1)
                Else
                    'cuando sea el ultimo
                    'redefinir la matriz con un indice menos
                    ReDim Preserve MATRIZ_LISTA(c - 1)
                End If
            
            Next
            EjecutarTema TemaDeMatriz
            CargarProximosTemas
            'un tema no puede durar menos de 5 segundo
            'terminado se debe dar cinco vueltas para esperar si esta empezando otro
            WAIT_EMPIEZA = 5
        End If

    End If
End Sub
