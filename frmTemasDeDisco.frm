VERSION 5.00
Begin VB.Form frmTemasDeDisco 
   AutoRedraw      =   -1  'True
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
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   8985
      Left            =   150
      TabIndex        =   1
      Top             =   30
      Width           =   11805
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
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
         Height          =   1305
         Left            =   7200
         TabIndex        =   9
         Top             =   7620
         Width           =   4515
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
            Height          =   950
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmdDiscoAd 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "frmTemasDeDisco.frx":0000
            Height          =   950
            Left            =   1200
            Picture         =   "frmTemasDeDisco.frx":0CFD
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmdDiscoAt 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "frmTemasDeDisco.frx":15D5
            Height          =   950
            Left            =   120
            Picture         =   "frmTemasDeDisco.frx":2347
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SALIR"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   950
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.ListBox lstEXT 
         BackColor       =   &H00404080&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   4620
         IntegralHeight  =   0   'False
         ItemData        =   "frmTemasDeDisco.frx":2C8A
         Left            =   180
         List            =   "frmTemasDeDisco.frx":2C9D
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   3960
         Visible         =   0   'False
         Width           =   6345
      End
      Begin VB.ListBox lstTIME 
         BackColor       =   &H00404080&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   8350
         IntegralHeight  =   0   'False
         Left            =   45
         TabIndex        =   5
         Top             =   480
         Width           =   1185
      End
      Begin VB.ListBox lstTemas 
         BackColor       =   &H00404080&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   8350
         IntegralHeight  =   0   'False
         Left            =   1230
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   5865
      End
      Begin VB.Label lblNoEjecuta 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "NO HAY CREDITO PARA EJECUTAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   795
         Left            =   7200
         TabIndex        =   7
         Top             =   6840
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   4515
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "TEMAS EN ESTE DISCO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   6
         Top             =   180
         Width           =   7065
      End
      Begin VB.Label lblDataDisco 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "No hay datos adicionales del disco"
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
         Height          =   3225
         Left            =   7200
         TabIndex        =   4
         Top             =   4200
         UseMnemonic     =   0   'False
         Width           =   4500
      End
      Begin VB.Label lblDisco 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   7200
         TabIndex        =   3
         Top             =   3660
         UseMnemonic     =   0   'False
         Width           =   4545
      End
      Begin VB.Image TapaCD 
         BorderStyle     =   1  'Fixed Single
         Height          =   3300
         Left            =   7740
         Stretch         =   -1  'True
         Top             =   315
         Width           =   3465
      End
   End
End
Attribute VB_Name = "frmTemasDeDisco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoHayTemasEnDisco As Boolean
Dim DuracionTema As String

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

Private Sub Command1_Click()
    Form_KeyDown TeclaOK, 0
End Sub

Private Sub Command2_Click()
    Form_KeyDown TeclaESC, 0
End Sub

Private Sub Form_Activate()
    Me.Refresh
    Label1 = "Buscando Temas de este disco..."
    Dim ArchTapa As String
    ArchTapa = UbicDiscoActual + "\tapa.jpg"
    If FSO.FileExists(ArchTapa) Then
        TapaCD.Picture = LoadPicture(ArchTapa)
    Else
        TapaCD.Picture = LoadPicture(AP + "tapa.jpg")
    End If
    TapaCD.Refresh
    lblDISCO = FSO.GetBaseName(UbicDiscoActual)
    Dim ArchDaTa As String
    ArchDaTa = UbicDiscoActual + "data.txt"
    If FSO.FileExists(ArchDaTa) Then
        Dim A As TextStream
        Set A = FSO.OpenTextFile(ArchDaTa, ForReading, False)
        lblDataDisco = A.ReadAll
    Else
        lblDataDisco = "No hay datos adicionales de este disco"
    End If
    
    'si estoy mostrando discos debo mostrar temas
    'se cargan los temas en una matriz con ubic archivo,nombreTema
    Dim c As Integer, nombreTemas As String
    Dim pathTema As String
    lstEXT.Clear
    If NoHayTemasEnDisco Then
        lstTEMAS.AddItem "No hay temas en este disco"
        lstTEMAS.Enabled = False
        lstTIME.Enabled = False
        WriteTBRLog "No hay temas en el disco: " + UbicDiscoActual, True
        Exit Sub
    End If
    c = 1
    Do While c <= UBound(MATRIZ_TEMAS)
        pathTema = txtInLista(MATRIZ_TEMAS(c), 0, ",")
        nombreTemas = txtInLista(MATRIZ_TEMAS(c), 1, ",")
        'quitar el molesto .mp3 o lo que fuera
        nombreTemas = FSO.GetBaseName(nombreTemas)
        lstTEMAS.AddItem nombreTemas
        lstTEMAS.Refresh
        lstEXT.AddItem pathTema
        c = c + 1
    Loop
    If CargarDuracionTemas Then
        'ahora cargar las duaciones
        Dim NoCargoDuracion As Long
        NoCargoDuracion = 0
        c = 1
        Dim MP3tmp As New MP3Info
        Do While c <= UBound(MATRIZ_TEMAS)
            pathTema = lstEXT.List(c - 1)
            'si es mp3 usar el rápido, si no usar el viejo
            If UCase(Right(pathTema, 3)) = "MP3" Then
                MP3tmp.FileName = pathTema
                DuracionTema = MP3tmp.DurationSTR
            Else
                'en caso de que sea video el clsMp3 no anda!!
                'mostrar duracion VIEJO FORMATO
                DuracionTema = frmIndex.MP3.QuickLargoDeTema(pathTema)
                If DuracionTema = "N/S" Then
                    NoCargoDuracion = NoCargoDuracion + 1
                    If NoCargoDuracion > 3 Then
                        lstTIME.Visible = False
                        lstTEMAS.Left = 50
                        lstTEMAS.Width = lblNoEjecuta.Left - 50
                    End If
                End If
            End If
            lstTIME.AddItem DuracionTema
            lstTIME.Refresh
            c = c + 1
        Loop
        Set MP3tmp = Nothing
        lstTIME.Enabled = True
    End If
    lstTEMAS.Enabled = True
    lstTEMAS.ListIndex = 0
    Label1 = "Temas de este disco"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Errores
    'y si no es una ficha la que se esta cargando
    lblNoEjecuta.Visible = False
    Select Case KeyCode
        Case TeclaCerrarSistema
            OnOffCAPS vbKeyCapital, False
            MostrarCursor True
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            frmIndex.MP3.DoClose
            If ApagarAlCierre Then APAGAR_PC
            End
        
        Case TeclaESC
            TECLAS_PRES = TECLAS_PRES + "4"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
            Unload Me
        Case TeclaOK
            'ver si esta habilitado
            TECLAS_PRES = TECLAS_PRES + "3"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
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
                temaElegido = lstEXT.List(lstTEMAS.ListIndex) ' UbicDiscoActual + "\" + lstTemas + "." + EXTs(lstTemas.ListIndex)
                
                'si esta ejecutando pasa a la lista de reproducción
                If frmIndex.MP3.IsPlaying Then
                    'pasar a la lista de reproducción
                    Dim NewIndLista As Long
                    NewIndLista = UBound(MATRIZ_LISTA)
                    ReDim Preserve MATRIZ_LISTA(NewIndLista + 1)
                    'se graba en Matriz_Listas como patah, nombre(sin .mp3)
                    MATRIZ_LISTA(NewIndLista + 1) = temaElegido + "," + lstTEMAS + " / " + FSO.GetBaseName(UbicDiscoActual)
                    CargarProximosTemas
                    'graba en reini.tbr los datos que correspondan por si se corta la luz
                    CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
                Else
                    'TEMA_REPRODUCIENDO y mp3.isplayin se cargan en ejecutartema
                    'paciencia
                    lstTEMAS.Enabled = False: lstTIME.Enabled = False
                    lstTEMAS.BackColor = vbBlack: lstTIME.BackColor = vbBlack
                    lstTEMAS.ForeColor = vbYellow
                    'lstTemas.Font.Size = 22 esto hace que parezca mas de un lstbox
                    lstTEMAS.Clear: lstTIME.Clear
                    lstTEMAS.AddItem "CARGANDO TEMA"
                    lstTEMAS.AddItem "ESPERE..."
                    lstTEMAS.Refresh: lstTIME.Refresh
                    CORTAR_TEMA = False 'este tema va entero ya que lo eligio el usuario
                    EjecutarTema temaElegido, True
                End If
                'pase lo que pase me vuelvo a los discos y cierro ventana actual
                
                Unload Me
            Else
                lblNoEjecuta.Visible = True
            End If
        
        Case TeclaDER
            TECLAS_PRES = TECLAS_PRES + "2"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
            If lstTEMAS.ListIndex < lstTEMAS.ListCount - 1 Then
                lstTEMAS.ListIndex = lstTEMAS.ListIndex + 1
            Else
                lstTEMAS.ListIndex = 0
            End If
        Case TeclaIZQ
            TECLAS_PRES = TECLAS_PRES + "1"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
            If lstTEMAS.ListIndex > 0 Then
                lstTEMAS.ListIndex = lstTEMAS.ListIndex - 1
            Else
                lstTEMAS.ListIndex = lstTEMAS.ListCount - 1
            End If
        Case TeclaPagAd
            TECLAS_PRES = TECLAS_PRES + "5"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
        Case TeclaPagAt
            TECLAS_PRES = TECLAS_PRES + "6"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
    End Select
    VerClaves TECLAS_PRES
    SecSinTecla = 0
    frmIndex.lblNoTecla = 0
    
    Exit Sub
Errores:
    WriteTBRLog "Error en temasDisco_KeyDown: " + Err.Description + " (" + CStr(Err.Number) + "). Se continua...", True
    Resume Next
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = TeclaNewFicha Then
        If CREDITOS <= MaximoFichas Then
            'apagar el fichero electronico
            OnOffCAPS vbKeyScrollLock, True
            CREDITOS = CREDITOS + TemasPorCredito
            SumarContadorCreditos TemasPorCredito
            If CREDITOS >= 10 Then
                frmIndex.lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
            Else
                frmIndex.lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
            End If
        Else
            'apagar el fichero electronico
            OnOffCAPS vbKeyScrollLock, False
        End If
    End If
End Sub

Private Sub Form_Load()
    AjustarFRM Me, 12000
    Frame1.Left = Screen.Width / 2 - Frame1.Width / 2
    Frame1.Top = Screen.Height / 2 - Frame1.Height / 2

    If MostrarTouch = False Then Frame2.Visible = False        'frame del touch

    ReDim MATRIZ_TEMAS(0) 'matriz en blanco
    'es una matriz global
    UbicDiscoActual = txtInLista(MATRIZ_DISCOS(nDiscoGral + 1), 0, ",")
    
    'encontrar todos los archivos *.mp3, *.avi, *.mpg, *.mpeg, etc
    ReDim Preserve MATRIZ_TEMAS(0)
    
    MATRIZ_TEMAS = ObtenerArchMM(UbicDiscoActual)
    
    
    If UBound(MATRIZ_TEMAS) = 0 Then
        NoHayTemasEnDisco = True
    Else
        NoHayTemasEnDisco = False
    End If
    'ocultar ahora
    If CargarDuracionTemas = False Then
        lstTIME.Visible = False
        lstTEMAS.Left = 50
        lstTEMAS.Width = lblNoEjecuta.Left - 150
    End If
    
End Sub

Private Sub lstTemas_Click()
    On Local Error Resume Next
    If CargarDuracionTemas Then lstTIME.ListIndex = lstTEMAS.ListIndex
End Sub

