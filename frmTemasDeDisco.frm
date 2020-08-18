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
   Begin VB.Timer RelojTDD 
      Enabled         =   0   'False
      Left            =   30
      Top             =   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   8985
      Left            =   150
      TabIndex        =   1
      Top             =   30
      Width           =   11805
      Begin VB.TextBox lstAgregados 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   960
         Left            =   45
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   7290
         Width           =   7080
      End
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
         TabIndex        =   8
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
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   9
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
         Height          =   1605
         IntegralHeight  =   0   'False
         ItemData        =   "frmTemasDeDisco.frx":2C8A
         Left            =   8010
         List            =   "frmTemasDeDisco.frx":2C9D
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   4905
         Visible         =   0   'False
         Width           =   2610
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
         Height          =   7750
         IntegralHeight  =   0   'False
         Left            =   45
         TabIndex        =   4
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
         Height          =   7750
         IntegralHeight  =   0   'False
         Left            =   1260
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   480
         Width           =   5865
      End
      Begin VB.Label lblCOMOSALIR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         Caption         =   "PRESIONE ESC PARA SALIR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   30
         TabIndex        =   14
         Top             =   8280
         Width           =   7065
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7980
         TabIndex        =   13
         Top             =   3090
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblPrecios 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "1 coin = 8 creditos / 8 creditos = 1 tema / 8 creditos = 1 VIDEO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   30
         TabIndex        =   12
         Top             =   8580
         Width           =   7070
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   3
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
         TabIndex        =   2
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
Dim SegSinTecla As Long
Dim NoHayTemasEnDisco As Boolean
Dim DuracionTema As String
Dim YaInicio As Long
'0=load
'1=keyup
'hasta que no haga un keyUp no da bola a ejecutar tema!!!!!!!

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
    Form_KeyUp TeclaOK, 0
End Sub

Private Sub Command2_Click()
    Form_KeyDown TeclaESC, 0
End Sub

Private Sub Form_Activate()
    Me.Refresh
    '
    'ver los precios!!!
    LineaError = "000-0024"
    MostrarCursor False
    'actualizar los precios
    
    '---------------------
    'si es gratis no usar!
    If CreditosCuestaTema = 0 And CreditosCuestaTemaVIDEO = 0 Then
        lblPrecios = "Modo Gratuito"
    Else
        If TemasPorCredito = 1 Then
            LineaError = "000-0026"
            lblPrecios = "1 coin = " + CStr(TemasPorCredito) + " credito"
        Else
            LineaError = "000-0027"
            lblPrecios = "1 coin = " + CStr(TemasPorCredito) + " creditos"
        End If
    End If
    '-------------------------
    LineaError = "000-0028"
    If CreditosCuestaTema = 1 Then
        LineaError = "000-0029"
        lblPrecios = lblPrecios + " / " + "1 credito = 1 tema"
    Else
        If CreditosCuestaTema = 0 Then
            lblPrecios = lblPrecios + " / " + " 1 tema = GRATIS!"
        Else
            LineaError = "000-0030"
            lblPrecios = lblPrecios + " / " + CStr(CreditosCuestaTema) + " creditos = 1 tema"
        End If
    End If
    'agreagr el precio de los videos!!!
    If CreditosCuestaTemaVIDEO = 1 Then
        LineaError = "000-0029"
        lblPrecios = lblPrecios + " / " + "1 credito = 1 VIDEO"
    Else
        If CreditosCuestaTemaVIDEO = 0 Then
            lblPrecios = lblPrecios + " / " + " 1 VIDEO = GRATIS!"
        Else
            LineaError = "000-0030"
            lblPrecios = lblPrecios + " / " + CStr(CreditosCuestaTemaVIDEO) + " creditos = 1 VIDEO"
        End If
    End If
    
    'total sería
    '1 coin = 8 creditos /// " + "8 creditos = 1 tema /// 8 creditos = 1 VIDEO
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Errores
    'y si no es una ficha la que se esta cargando
    lblNoEjecuta.Visible = False
    Select Case KeyCode
        Case vbKeyF4
            If Shift = 4 Then
                Unload Me
            End If
        Case vbKeyU
            frmOnlyContador.Show 1
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
            Exit Sub
        
        Case TeclaDER
            'si esta en el modo 5 debe salir!!!
            If IsMod46Teclas = 46 Then
                If lstTemas.ListIndex < lstTemas.ListCount - 1 Then
                    lstTemas.ListIndex = lstTemas.ListIndex + 1
                Else
                    lstTemas.ListIndex = 0
                End If
                SaltarEspaciosLstTemas True
            End If
            If IsMod46Teclas = 5 Then
                'igual que el escape!!!
                TECLAS_PRES = TECLAS_PRES + "2"
                TECLAS_PRES = Right(TECLAS_PRES, 20)
                frmIndex.lblTECLAS = TECLAS_PRES
                Unload Me
                Exit Sub
            End If
            TECLAS_PRES = TECLAS_PRES + "2"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
        Case TeclaIZQ
            
            If IsMod46Teclas = 46 Then
                If lstTemas.ListIndex > 0 Then
                    lstTemas.ListIndex = lstTemas.ListIndex - 1
                Else
                    lstTemas.ListIndex = lstTemas.ListCount - 1
                End If
                SaltarEspaciosLstTemas False
            End If
            
            If IsMod46Teclas = 5 Then
                'igual que el escape!!!
                TECLAS_PRES = TECLAS_PRES + "1"
                TECLAS_PRES = Right(TECLAS_PRES, 20)
                frmIndex.lblTECLAS = TECLAS_PRES
                Unload Me
                Exit Sub
            End If
            
            TECLAS_PRES = TECLAS_PRES + "1"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
            
        Case TeclaPagAd
            If IsMod46Teclas = 5 Then
                'igual que el boton adelante!!
                If lstTemas.ListIndex < lstTemas.ListCount - 1 Then
                    lstTemas.ListIndex = lstTemas.ListIndex + 1
                Else
                    lstTemas.ListIndex = 0
                End If
                SaltarEspaciosLstTemas True
            End If
            TECLAS_PRES = TECLAS_PRES + "5"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
        Case TeclaPagAt
            If IsMod46Teclas = 5 Then
                'igual que el boton atras!!
                If lstTemas.ListIndex > 0 Then
                    lstTemas.ListIndex = lstTemas.ListIndex - 1
                Else
                    lstTemas.ListIndex = lstTemas.ListCount - 1
                End If
                SaltarEspaciosLstTemas False
            End If
            TECLAS_PRES = TECLAS_PRES + "6"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
    End Select
    SegSinTecla = 0 'protector para salir de esta frm
    VerClaves TECLAS_PRES
    SecSinTecla = 0 'preteccion global de pantalla
    frmIndex.lblNoTecla = 0
    
    Exit Sub
Errores:
    WriteTBRLog "Error en temasDisco_KeyDown: " + Err.Description + " (" + CStr(Err.Number) + "). Se continua...", True
    Resume Next
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'ver detalle mas abajo de que mierda es esto y eln el gral de este frm
    YaInicio = YaInicio + 1
    'puede no escuchar el coin!!!!!!
    'esto se pone mas abajo!!!!
    'If YaInicio <= 1 Then Exit Sub
    Select Case KeyCode
        Case TeclaNewFicha
            If CREDITOS <= MaximoFichas Then
                'apagar el fichero electronico
                OnOffCAPS vbKeyScrollLock, True
                CREDITOS = CREDITOS + TemasPorCredito
                SumarContadorCreditos TemasPorCredito
                
                ShowCredits
                
            Else
                'apagar el fichero electronico
                OnOffCAPS vbKeyScrollLock, False
            End If
        'lo pongo aca para que no se tilden apretando
        ' y se marque mil canciones
        
        'es jodido ya que se aprieta OK para entrar aqui y se suelta cuando ya se abrio!!!
        'entocnes se elige el primer tema.....que cagada
        
        'para solucionarlo pongo una variable de inicialización
        'en el key down...o se a que hasta que no apriete una tecla aqui
        'no toma la soltada de tecla (keyup)
        
        Case TeclaOK
            If YaInicio <= 1 Then Exit Sub
            'ver si esta habilitado
            TECLAS_PRES = TECLAS_PRES + "3"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
            
            'ANTES DE VER CUANTOS CREDITOS NECESITA TENGO QUE SABER SI QUIERE EJECUTAR
            'MP3 O VIDEO!!!!!!
            Dim temaElegido As String
            'lstext es una lista oculta  con datos completos
            temaElegido = lstEXT.List(lstTemas.ListIndex) ' UbicDiscoActual + "\" + lstTemas + "." + EXTs(lstTemas.ListIndex)
            
            If LCase(Right(temaElegido, 3)) = "mp3" Then
                PideVideo = False
            Else
                PideVideo = True
            End If
            
            'ver si puede pagar lo que pide!!!
            'que joyita papa!!!. Parece que supieras programar
            '--------------------------------------------------------------
            If (PideVideo = False And CREDITOS >= CreditosCuestaTema) Or _
                (PideVideo And CREDITOS >= CreditosCuestaTemaVIDEO) Then
            '--------------------------------------------------------------
                'restar lo que corresponde!!!
                If PideVideo Then
                    CREDITOS = CREDITOS - CreditosCuestaTemaVIDEO
                Else
                    CREDITOS = CREDITOS - CreditosCuestaTema
                End If
                
                'siempre que se ejecute un credito estaremos por debajo de maximo
                OnOffCAPS vbKeyScrollLock, True
                'grabar cant de creditos
                EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
                
                ShowCredits
                
                'grabar credito para validar
                'creditosValidar ya se cargo en load de frmindex
                CreditosValidar = CreditosValidar + TemasPorCredito
                EscribirArch1Linea SYSfolder + "\radilav.cfg", CStr(CreditosValidar)
                
                'si esta ejecutando pasa a la lista de reproducción
                'si esta ejecutando una prueba SACARLA!!!
                If frmIndex.MP3.IsPlaying And CORTAR_TEMA = False Then
                    'pasar a la lista de reproducción
                    Dim NewIndLista As Long
                    NewIndLista = UBound(MATRIZ_LISTA)
                    ReDim Preserve MATRIZ_LISTA(NewIndLista + 1)
                    'se graba en Matriz_Listas como patah, nombre(sin .mp3)
                    MATRIZ_LISTA(NewIndLista + 1) = temaElegido + "," + lstTemas + " / " + FSO.GetBaseName(UbicDiscoActual)
                    CargarProximosTemas
                    'graba en reini.tbr los datos que correspondan por si se corta la luz
                    CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
                    'AHORA DEBE MARCARLO COMO EJECUTADO Y SALIR PARA ELIJA OTRO
                    lstAgregados = lstAgregados + lstTemas.List(lstTemas.ListIndex) + " / "
                    
                    If BloquearMusicaElegida Then
                        lstTemas.List(lstTemas.ListIndex) = "----------"
                        lstTIME.List(lstTIME.ListIndex) = "---"
                    End If
                        
                    lstAgregados.Visible = True
                    lstTemas.Height = lstAgregados.Top - lstTemas.Top
                    lstTIME.Height = lstAgregados.Top - lstTIME.Top
                    SaltarEspaciosLstTemas True
                Else
                    'TEMA_REPRODUCIENDO y mp3.isplayin se cargan en ejecutartema
                    
                    ''ESTO SE HACIA ANTES PARA SALIR!!!!!!!!
                    ''----------------------
                    ''----------------------
                    ''paciencia
                    'lstTemas.Enabled = False: lstTIME.Enabled = False
                    'lstTemas.BackColor = vbBlack: lstTIME.BackColor = vbBlack
                    'lstTemas.ForeColor = vbYellow
                    ''lstTemas.Font.Size = 22 esto hace que parezca mas de un lstbox
                    'lstTemas.Clear: lstTIME.Clear
                    'lstTemas.AddItem "CARGANDO TEMA"
                    'lstTemas.AddItem "ESPERE..."
                    'lstTemas.Refresh: lstTIME.Refresh
                    ''----------------------
                    ''----------------------
                    'AHORA DEBE MARCARLO COMO EJECUTADO Y SALIR PARA ELIJA OTRO
                    lstAgregados = lstAgregados + lstTemas.List(lstTemas.ListIndex) + " / "
                    
                    If BloquearMusicaElegida Then
                        lstTemas.List(lstTemas.ListIndex) = "----------"
                        lstTIME.List(lstTIME.ListIndex) = "---"
                    End If
                    
                    lstAgregados.Visible = True
                    lstTemas.Height = lstAgregados.Top - lstTemas.Top
                    lstTIME.Height = lstAgregados.Top - lstTIME.Top
                    SaltarEspaciosLstTemas True
                    CORTAR_TEMA = False 'este tema va entero ya que lo eligio el usuario
                    Me.ZOrder
                    EjecutarTema temaElegido, True
                End If
                
                VerSiTocaPUB
                'dejo seguir eligiendo y no salgo!!!
                'Unload Me
            Else
                lblNoEjecuta.Visible = True
            End If
        
    End Select
    
    
End Sub

Private Sub Form_Load()
    YaInicio = 0
    If Is3pmExclusivo Then
        lstTemas.BackColor = vbBlack
        lstTIME.BackColor = vbBlack
        lstTemas.ForeColor = vbYellow
        lstTIME.ForeColor = vbYellow
        Frame1.BackColor = &H404000
        lblDataDisco.Visible = False
    End If
    If IsMod46Teclas = 5 Then
        lblCOMOSALIR = "PRESIONE FLECHA HORIZONTAL PARA SALIR"
    End If
    'esconder y mostrar cuando corresponda!!
    lstAgregados.Visible = False
    lstAgregados = ""
    lstAgregados = "ELEGIDOS" + vbCrLf
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
        lstTemas.Left = 50
        lstTemas.Width = lblNoEjecuta.Left - 150
    End If
    SegSinTecla = 0
    RelojTDD.Enabled = True
    RelojTDD.Interval = 1000
    
    Label1 = "Buscando Temas de este disco..."
    Dim ArchTapa As String
    ArchTapa = UbicDiscoActual + "\tapa.jpg"
    If FSO.FileExists(ArchTapa) Then
        TapaCD.Picture = LoadPicture(ArchTapa)
    Else
        TapaCD.Picture = LoadPicture(SYSfolder + "f61.dlw")
    End If
    TapaCD.Refresh
    lblDisco = FSO.GetBaseName(UbicDiscoActual)
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
    Dim C As Integer, nombreTemas As String
    Dim pathTema As String
    lstEXT.Clear
    If NoHayTemasEnDisco Then
        lstTemas.AddItem "No hay temas en este disco"
        lstTemas.Enabled = False
        lstTIME.Enabled = False
        WriteTBRLog "No hay temas en el disco: " + UbicDiscoActual, True
        Exit Sub
    End If
    C = 1
    Do While C <= UBound(MATRIZ_TEMAS)
        pathTema = txtInLista(MATRIZ_TEMAS(C), 0, ",")
        nombreTemas = txtInLista(MATRIZ_TEMAS(C), 1, ",")
        'quitar el molesto .mp3 o lo que fuera
        nombreTemas = FSO.GetBaseName(nombreTemas)
        lstTemas.AddItem nombreTemas
        lstTemas.Refresh
        lstEXT.AddItem pathTema
        C = C + 1
    Loop
    If CargarDuracionTemas Then
        'ahora cargar las duaciones
        Dim NoCargoDuracion As Long
        NoCargoDuracion = 0
        C = 1
        Dim MP3tmp As New MP3Info
        Do While C <= UBound(MATRIZ_TEMAS)
            pathTema = lstEXT.List(C - 1)
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
                        lstTemas.Left = 50
                        lstTemas.Width = lblNoEjecuta.Left - 50
                    End If
                End If
            End If
            lstTIME.AddItem DuracionTema
            lstTIME.Refresh
            C = C + 1
        Loop
        Set MP3tmp = Nothing
        lstTIME.Enabled = True
    End If
    lstTemas.Enabled = True
    lstTemas.ListIndex = 0
    Label1 = "Temas de este disco"
    
    
End Sub

Private Sub lstTemas_Click()
    On Local Error Resume Next
    If CargarDuracionTemas Then lstTIME.ListIndex = lstTemas.ListIndex
End Sub

Private Sub RelojTDD_Timer()
    'relojTemasDeDisco
    SegSinTecla = SegSinTecla + 1
    Label2 = SegSinTecla
    If SegSinTecla = 20 Then
        RelojTDD.Enabled = False
        Unload Me
    End If
    
End Sub
Private Sub SaltarEspaciosLstTemas(HaciaAdelante As Boolean)
    'cuando eligo un tema lo saco para que no haga macana
    'el secreto es no generar el listindex salvo que se haya encontrado...
    'uso la prop LIST() que puede ver sin tocar!!!!!!!
    Dim A As Long
    Dim CC As Long
    Dim Ahora As Long
    Ahora = lstTemas.ListIndex
    
    Dim nINI As Long, nFin As Long, StepMio As Long
    If HaciaAdelante Then
        nINI = Ahora
        nFin = lstTemas.ListCount - 1
        StepMio = 1
    Else
        nINI = Ahora
        nFin = 0
        StepMio = -1
    End If
    Dim Vueltas As Long
    Vueltas = 0
ReiniLST:
    Vueltas = Vueltas + 1
    'si da 4 vueltas es que no hay!!
    If Vueltas = 4 Then
        Unload Me
        Exit Sub
    End If
    For A = nINI To nFin Step StepMio
        If lstTemas.List(A) <> "----------" Then
            'ya esta lo encontro!!!!!!!
            'ir ahi!!!
            lstTemas.ListIndex = A
            Exit For
        Else
            'si es el ultimo......!!
            If HaciaAdelante Then
                If A = nFin Then 'este es lstTemas.ListCount - 1
                    'voy al primero
                    nINI = 0
                    GoTo ReiniLST
                End If
            Else
                If A = nFin Then 'este es 0
                    'voy al ultimo
                    nINI = lstTemas.ListCount - 1
                    GoTo ReiniLST
                End If
            End If
        End If
    Next
End Sub
