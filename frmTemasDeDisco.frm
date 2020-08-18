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
      TabIndex        =   0
      Top             =   -30
      Width           =   11805
      Begin VB.ListBox lstEXT 
         BackColor       =   &H00404080&
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
         ForeColor       =   &H00C0E0FF&
         Height          =   4890
         ItemData        =   "frmTemasDeDisco.frx":0000
         Left            =   5400
         List            =   "frmTemasDeDisco.frx":0013
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   3840
         Visible         =   0   'False
         Width           =   6345
      End
      Begin VB.ListBox lstTIME 
         BackColor       =   &H00404080&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   8430
         Left            =   45
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.ListBox lstTemas 
         BackColor       =   &H00404080&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   8430
         Left            =   780
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   6315
      End
      Begin VB.Label lblNoEjecuta 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "INGRESE FICHA PARA EJECUTAR MUSICA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   885
         Left            =   7200
         TabIndex        =   6
         Top             =   8010
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   4545
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "TEMAS EN ESTE DISCO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   180
         Width           =   7065
      End
      Begin VB.Label lblDataDisco 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "No hay datos adicionales del disco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3465
         Left            =   7200
         TabIndex        =   3
         Top             =   4500
         UseMnemonic     =   0   'False
         Width           =   4545
      End
      Begin VB.Label lblDisco 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   795
         Left            =   7200
         TabIndex        =   2
         Top             =   3645
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
        Dim a As TextStream
        Set a = FSO.OpenTextFile(ArchDaTa, ForReading, False)
        lblDataDisco = a.ReadAll
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
        Do While c <= UBound(MATRIZ_TEMAS)
            pathTema = lstEXT.List(c - 1)
            'mostrar duracion
            DuracionTema = frmINDEX.MP3.QuickLargoDeTema(pathTema)
            If DuracionTema = "N/S" Then
                NoCargoDuracion = NoCargoDuracion + 1
                If NoCargoDuracion > 3 Then
                    lstTIME.Visible = False
                    lstTEMAS.Left = 50
                    lstTEMAS.Width = lblNoEjecuta.Left - 50
                End If
            End If
            lstTIME.AddItem DuracionTema
            lstTIME.Refresh
            c = c + 1
        Loop
        lstTIME.Enabled = True
    End If
    lstTEMAS.Enabled = True
    lstTEMAS.ListIndex = 0
    Label1 = "Temas de este disco"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'y si no es una ficha la que se esta cargando
    lblNoEjecuta.Visible = False
    Select Case KeyCode
        Case TeclaCerrarSistema
            OnOffCAPS vbKeyCapital, False
            MostrarCursor True
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            frmINDEX.MP3.DoClose
            If ApagarAlCierre Then APAGAR_PC
            End
        Case TeclaNewFicha
            'si ya hay 9 cargados se traga las fichas
            If CREDITOS <= MaximoFichas Then
                'apagar el fichero electronico
                OnOffCAPS vbKeyScrollLock, True
                CREDITOS = CREDITOS + 1
                SumarContadorCreditos 1
                If CREDITOS >= 10 Then
                    frmINDEX.lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
                Else
                    frmINDEX.lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
                End If
                Unload Me
            Else
                'apagar el fichero electronico
                OnOffCAPS vbKeyScrollLock, False
            End If
        Case TeclaESC
            TECLAS_PRES = TECLAS_PRES + "4"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmINDEX.lblTECLAS = TECLAS_PRES
            
            Unload Me
        Case TeclaOK
            'ver si esta habilitado
            TECLAS_PRES = TECLAS_PRES + "3"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmINDEX.lblTECLAS = TECLAS_PRES
            If CREDITOS > 0 Then
                CREDITOS = CREDITOS - 1
                'siempre que se ejecute un credito estaremos por debajo de maximo
                OnOffCAPS vbKeyScrollLock, True
                'grabar cant de creditos
                EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
                If CREDITOS < 10 Then frmINDEX.lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
                If CREDITOS >= 10 Then frmINDEX.lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
                Dim temaElegido As String
                'lstext es una lista oculta  con datos completos
                temaElegido = lstEXT.List(lstTEMAS.ListIndex) ' UbicDiscoActual + "\" + lstTemas + "." + EXTs(lstTemas.ListIndex)
                
                'si esta ejecutando pasa a la lista de reproducción
                If frmINDEX.MP3.IsPlaying Then
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
            frmINDEX.lblTECLAS = TECLAS_PRES
            If lstTEMAS.ListIndex < lstTEMAS.ListCount - 1 Then
                lstTEMAS.ListIndex = lstTEMAS.ListIndex + 1
            Else
                lstTEMAS.ListIndex = 0
            End If
        Case TeclaIZQ
            TECLAS_PRES = TECLAS_PRES + "1"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmINDEX.lblTECLAS = TECLAS_PRES
            If lstTEMAS.ListIndex > 0 Then
                lstTEMAS.ListIndex = lstTEMAS.ListIndex - 1
            Else
                lstTEMAS.ListIndex = lstTEMAS.ListCount - 1
            End If
    End Select
    VerClaves TECLAS_PRES
    SecSinTecla = 0
    frmINDEX.lblNoTecla = 0
End Sub

Private Sub Form_Load()
    AjustarFRM Me, 12000
    Frame1.Left = Screen.Width / 2 - Frame1.Width / 2
    Frame1.Top = Screen.Height / 2 - Frame1.Height / 2

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
        lstTEMAS.Width = lblNoEjecuta.Left - 50
    End If
    
End Sub

Private Sub lstTemas_Click()
    On Local Error Resume Next
    If CargarDuracionTemas Then lstTIME.ListIndex = lstTEMAS.ListIndex
End Sub

