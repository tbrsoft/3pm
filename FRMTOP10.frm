VERSION 5.00
Begin VB.Form FRMTOP10 
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame FR 
      BackColor       =   &H00000080&
      Height          =   8985
      Left            =   0
      TabIndex        =   1
      Top             =   0
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
         Height          =   2055
         Left            =   9205
         TabIndex        =   6
         Top             =   6900
         Width           =   2595
         Begin VB.CommandButton cmdDiscoAt 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "FRMTOP10.frx":0000
            Height          =   870
            Left            =   120
            Picture         =   "FRMTOP10.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1110
            Width           =   1150
         End
         Begin VB.CommandButton cmdDiscoAd 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "FRMTOP10.frx":16B5
            Height          =   870
            Left            =   1350
            Picture         =   "FRMTOP10.frx":23B2
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1110
            Width           =   1150
         End
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
            Height          =   870
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   210
            Width           =   1150
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
            Height          =   870
            Left            =   1350
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   210
            Width           =   1150
         End
      End
      Begin VB.Label lblNoEjecuta 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "No hay credito para ejecutar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   645
         Left            =   9180
         TabIndex        =   4
         Top             =   6240
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Estos son los más escuchados. La mejor música elegida por ustedes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   150
         Width           =   9915
      End
      Begin VB.Label lblWAIT 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "CARGANDO TEMA     ESPERE..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   3870
         TabIndex        =   3
         Top             =   3645
         Visible         =   0   'False
         Width           =   3930
      End
      Begin VB.Image Image1 
         Height          =   1800
         Left            =   10080
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1650
      End
      Begin VB.Label lblPuestos 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   345
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   390
         Width           =   9930
      End
   End
End
Attribute VB_Name = "FRMTOP10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pANT As Integer
Dim MTXtop() As String
Dim MTXtemas() As String
Dim MTXdiscos() As String
Dim PuestoElegido As Integer

Dim MaxTop As Integer

Dim ColorUnSel As Long
Dim ColorSel As Long
Dim ForeColorTop As Long

Private Sub Command2_Click()
    Form_KeyDown TeclaESC, 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'y si no es una ficha la que se esta cargando
    lblNoEjecuta.Visible = False
    Select Case KeyCode
        Case TeclaConfig
            frmConfig.Show 1
        
        Case TeclaIZQ
            
            TECLAS_PRES = TECLAS_PRES + "1"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
            If IsMod46Teclas = 46 Then
                'ver que sea un puesto valido
                'se define como válido se tiene untexto de más de 5 caracteres
                pANT = PuestoElegido
                PuestoElegido = PuestoElegido - 1
                If PuestoElegido = -1 Then PuestoElegido = MaxTop - 1
                If Len(lblPuestos(PuestoElegido)) > 5 Then
                    lblPuestos(pANT).BackColor = ColorUnSel
                    lblPuestos(PuestoElegido).BackColor = ColorSel
                Else
                    'reacomodar puesto elegido
                    PuestoElegido = pANT
                End If
            End If
            If IsMod46Teclas = 5 Then
                Unload Me
                Exit Sub
            End If
            
        Case TeclaDER
            TECLAS_PRES = TECLAS_PRES + "2"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
            If IsMod46Teclas = 46 Then
                'ver que sea un puesto valido
                'se define como válido se tiene untexto de más de 5 caracteres
                pANT = PuestoElegido
                PuestoElegido = PuestoElegido + 1
                If PuestoElegido = MaxTop Then PuestoElegido = 0
                If Len(lblPuestos(PuestoElegido)) > 5 Then
                    'unsel el elegido
                    lblPuestos(pANT).BackColor = ColorUnSel
                    lblPuestos(PuestoElegido).BackColor = ColorSel
                Else
                    PuestoElegido = pANT
                End If
            End If
            If IsMod46Teclas = 5 Then
                Unload Me
                Exit Sub
            End If
        Case TeclaESC
            TECLAS_PRES = TECLAS_PRES + "4"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
            
            Unload Me
            Exit Sub
        Case TeclaOK
            TECLAS_PRES = TECLAS_PRES + "3"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
            'primero que pide!!!
            Dim temaElegido As String
            If PuestoElegido >= UBound(MTXtop) Then
                MsgBox "No hay tema elegido!!"
                Exit Sub
            End If
            temaElegido = MTXtop(PuestoElegido + 1)
            
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
                
                ShowCredits
                
                'si esta ejecutando pasa a la lista de reproducción
                If frmIndex.MP3.IsPlaying And CORTAR_TEMA = False Then
                    'pasar a la lista de reproducción
                    Dim NewIndLista As Long
                    NewIndLista = UBound(MATRIZ_LISTA)
                    ReDim Preserve MATRIZ_LISTA(NewIndLista + 1)
                    'se graba en Matriz_Listas como patah, nombre(sin .mp3)
                    MATRIZ_LISTA(NewIndLista + 1) = temaElegido + "," + MTXtemas(PuestoElegido + 1) + " / " + MTXdiscos(PuestoElegido + 1)
                    CargarProximosTemas
                    'graba en reini.tbr los datos que correspondan por si se corta la luz
                    CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
                Else
                    'ocultar el rank y mostrar lblWAIT
                    lblWait = "CARGANDO TEMA" + vbCrLf + "ESPERE..."
                    Dim cRank As Integer
                    cRank = 0
                    Do While cRank < MaxTop
                        lblPuestos(cRank).Visible = False
                        'lblPuestos(cRank).Refresh
                        cRank = cRank + 1
                    Loop
                    lblWait.Visible = True
                    lblWait.Refresh
                    'TEMA_REPRODUCIENDO y mp3.isplayin se cargan en ejecutartema
                    CORTAR_TEMA = False 'este tema va entero ya que lo eligio el usuario
                    EjecutarTema temaElegido, True
                End If
                
                VerSiTocaPUB
                
                'pase lo que pase me vuelvo a los discos y cierro ventana actual
                
                Unload Me
            Else
                lblNoEjecuta.Visible = True
            End If
                
        Case TeclaCerrarSistema
            OnOffCAPS vbKeyCapital, False
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            frmIndex.MP3.DoClose
            MostrarCursor True
            If ApagarAlCierre Then APAGAR_PC
            End
        Case TeclaPagAd
            TECLAS_PRES = TECLAS_PRES + "5"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
            If IsMod46Teclas = 5 Then
                'ver que sea un puesto valido
                'se define como válido se tiene untexto de más de 5 caracteres
                pANT = PuestoElegido
                PuestoElegido = PuestoElegido + 1
                If PuestoElegido = MaxTop Then PuestoElegido = 0
                If Len(lblPuestos(PuestoElegido)) > 5 Then
                    'unsel el elegido
                    lblPuestos(pANT).BackColor = ColorUnSel
                    lblPuestos(PuestoElegido).BackColor = ColorSel
                Else
                    PuestoElegido = pANT
                End If
            End If
            
        Case TeclaPagAt
            TECLAS_PRES = TECLAS_PRES + "6"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
            If IsMod46Teclas = 5 Then
                'ver que sea un puesto valido
                'se define como válido se tiene untexto de más de 5 caracteres
                pANT = PuestoElegido
                PuestoElegido = PuestoElegido - 1
                If PuestoElegido = -1 Then PuestoElegido = MaxTop - 1
                If Len(lblPuestos(PuestoElegido)) > 5 Then
                    lblPuestos(pANT).BackColor = ColorUnSel
                    lblPuestos(PuestoElegido).BackColor = ColorSel
                Else
                    'reacomodar puesto elegido
                    PuestoElegido = pANT
                End If
            End If
    End Select
    VerClaves TECLAS_PRES
    SecSinTecla = 0
    frmIndex.lblNoTecla = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = TeclaNewFicha Then
        'si ya hay 9 cargados se traga las fichas
        If CREDITOS <= MaximoFichas Then
            'apagar el fichero electronico
            OnOffCAPS vbKeyScrollLock, True
            CREDITOS = CREDITOS + TemasPorCredito
            SumarContadorCreditos TemasPorCredito
            'grabar cant de creditos
            EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
            
            ShowCredits
            
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
    
    Image1.Picture = LoadPicture(SYSfolder + "f8ya.nam")
    'si es SL cambiar
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(WINfolder + "SL\indexchi.tbr") Then
            Image1.Picture = LoadPicture(WINfolder + "SL\indexchi.tbr")
        End If
    End If
    If MostrarTouch = False Then Frame2.Visible = False        'frame del touch
    
    ColorUnSel = 1
    ColorSel = vbRed
    ForeColorTop = vbYellow
    PuestoElegido = 0
    MaxTop = 30
    
    'mostrar todos los lbls
    Dim C As Integer
    C = 0
    lblPuestos(0).BackColor = ColorUnSel
    lblPuestos(0).ForeColor = ForeColorTop
    
    Do While C < MaxTop - 1
        C = C + 1
        Load lblPuestos(C)
        If C > 0 And C < 10 Then
            lblPuestos(C).Font.Size = 12
            lblPuestos(C).Height = 300
        End If
        If C >= 10 Then
            lblPuestos(C).Font.Size = 10
            lblPuestos(C).Height = 250
        End If
        If C = 1 Or C = 10 Or C = 20 Then
            lblPuestos(C).Top = lblPuestos(C - 1).Top + lblPuestos(C - 1).Height + 150
        Else
            lblPuestos(C).Top = lblPuestos(C - 1).Top + lblPuestos(C - 1).Height
        End If
        lblPuestos(C).Width = lblPuestos(C - 1).Width
        If C = 5 Then lblPuestos(C).Width = 11650
        If C >= 20 Then
            lblPuestos(C).Font.Size = 8
            lblPuestos(C).Height = 250
            lblPuestos(C).Width = Frame2.Left - 100
        End If
        lblPuestos(C).Visible = True
        lblPuestos(C).Refresh
    Loop
    
    'leer ranking.tbr y cargar los temas que haya
    
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        FSO.CreateTextFile AP + "ranking.tbr", True
    End If
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    Dim TT As String
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    C = 0
    'INICIALIAZAR LA MATRIZ si no hay error al poner OK sin nada en el rank!!
    ReDim Preserve MTXtop(0)
    Do While Not TE.AtEndOfStream
        TT = TE.ReadLine
        ThisPTS = Val(txtInLista(TT, 0, ","))
        ThisArch = txtInLista(TT, 1, ",")
        ThisTEMA = txtInLista(TT, 2, ",")
        ThisDISCO = txtInLista(TT, 3, ",")
            
        If C = MaxTop Then Exit Do
        'si elarchivo no existe no se debe cargar
        If FSO.FileExists(ThisArch) Then
            lblPuestos(C).UseMnemonic = False
            lblPuestos(C) = " " + Trim(Str(C + 1)) + "º " + _
            QuitarNumeroDeTema(ThisTEMA) + " / " + ThisDISCO + " [" + Trim(Str(ThisPTS)) + " pts]"
            lblPuestos(C).Refresh
            
            C = C + 1
            ReDim Preserve MTXtop(C)
            MTXtop(C) = ThisArch
            ReDim Preserve MTXtemas(C)
            MTXtemas(C) = ThisTEMA
            ReDim Preserve MTXdiscos(C)
            MTXdiscos(C) = ThisDISCO
        End If
    Loop
    TE.Close
    'élegir el primero
    lblPuestos(0).BackColor = ColorSel
    AjustarFRM Me, 12000
    FR.Left = Screen.Width / 2 - FR.Width / 2
    FR.Top = Screen.Height / 2 - FR.Height / 2
    
End Sub

Private Sub cmdDiscoAt_Click()
    Form_KeyDown TeclaIZQ, 0
    Command1.SetFocus
End Sub

Private Sub cmdDiscoAt_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub cmdDiscoAd_Click()
    Form_KeyDown TeclaDER, 0
    Command1.SetFocus
End Sub

Private Sub cmdDiscoAd_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub Command1_Click()
    Form_KeyDown TeclaOK, 0
End Sub

