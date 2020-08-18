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
      Left            =   90
      TabIndex        =   0
      Top             =   -45
      Width           =   11775
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
         Height          =   1935
         Left            =   9450
         TabIndex        =   5
         Top             =   6960
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   2265
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
         TabIndex        =   4
         Top             =   3645
         Visible         =   0   'False
         Width           =   3930
      End
      Begin VB.Image Image1 
         Height          =   1800
         Left            =   10080
         Picture         =   "FRMTOP10.frx":0000
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1650
      End
      Begin VB.Label lblPuestos 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
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
         Left            =   90
         TabIndex        =   3
         Top             =   180
         Width           =   9930
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Top 3PM"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   960
         Index           =   0
         Left            =   10125
         TabIndex        =   2
         Top             =   1980
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Lo mejor..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Index           =   1
         Left            =   10125
         TabIndex        =   1
         Top             =   2880
         Width           =   1575
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'y si no es una ficha la que se esta cargando
    lblNoEjecuta.Visible = False
    Select Case KeyCode
        Case TeclaConfig
            frmConfig.Show 1
        Case TeclaNewFicha
            'si ya hay 9 cargados se traga las fichas
            If CREDITOS <= MaximoFichas Then
                'apagar el fichero electronico
                OnOffCAPS vbKeyScrollLock, True
                CREDITOS = CREDITOS + 1
                SumarContadorCreditos 1
                'grabar cant de creditos
                EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
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
        Case TeclaIZQ
            TECLAS_PRES = TECLAS_PRES + "1"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmINDEX.lblTECLAS = TECLAS_PRES
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
        Case TeclaDER
            TECLAS_PRES = TECLAS_PRES + "2"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmINDEX.lblTECLAS = TECLAS_PRES
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
        Case TeclaESC
            TECLAS_PRES = TECLAS_PRES + "4"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmINDEX.lblTECLAS = TECLAS_PRES
            
            Unload Me
        Case TeclaOK
            TECLAS_PRES = TECLAS_PRES + "3"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmINDEX.lblTECLAS = TECLAS_PRES
            'ver si esta habilitado
            If CREDITOS > 0 Then
                CREDITOS = CREDITOS - 1
                'siempre que se ejecute un credito estaremos por debajo de maximo
                OnOffCAPS vbKeyScrollLock, True
                If CREDITOS < 10 Then frmINDEX.lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
                If CREDITOS >= 10 Then frmINDEX.lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
                Dim temaElegido As String
                temaElegido = MTXtop(PuestoElegido + 1)
                
                'si esta ejecutando pasa a la lista de reproducción
                If frmINDEX.MP3.IsPlaying Then
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
                'pase lo que pase me vuelvo a los discos y cierro ventana actual
                
                Unload Me
            Else
                lblNoEjecuta.Visible = True
            End If
                
        Case TeclaCerrarSistema
            OnOffCAPS vbKeyCapital, False
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            frmINDEX.MP3.DoClose
            MostrarCursor True
            If ApagarAlCierre Then APAGAR_PC
            
            End
    End Select
    VerClaves TECLAS_PRES
    SecSinTecla = 0
    frmINDEX.lblNoTecla = 0
End Sub

Private Sub Form_Load()
    ColorUnSel = 1
    ColorSel = vbRed
    ForeColorTop = vbYellow
    PuestoElegido = 0
    ''el maximo depende de la definicion de pantalla
    'If Screen.Width > 9000 Then MaxTop = 26 '640x480
    'If Screen.Width > 11400 Then MaxTop = 26 '800x600
    'If Screen.Width > 14760 Then MaxTop = 26 '1024x768
    'If Screen.Width > 18600 Then MaxTop = 26 '1280x1024
    'corregirdo lo anterior, siempre son 26
    MaxTop = 31
    
    'mostrar todos los lbls
    Dim c As Integer
    c = 0
    lblPuestos(0).BackColor = ColorUnSel
    lblPuestos(0).ForeColor = ForeColorTop
    
    Do While c < MaxTop - 1
        c = c + 1
        Load lblPuestos(c)
        If c > 0 And c < 10 Then
            lblPuestos(c).Font.Size = 12
            lblPuestos(c).Height = 300
        End If
        If c >= 10 Then
            lblPuestos(c).Font.Size = 10
            lblPuestos(c).Height = 250
        End If
        If c = 1 Or c = 10 Or c = 20 Then
            lblPuestos(c).Top = lblPuestos(c - 1).Top + lblPuestos(c - 1).Height + 150
        Else
            lblPuestos(c).Top = lblPuestos(c - 1).Top + lblPuestos(c - 1).Height
        End If
        lblPuestos(c).Visible = True
        lblPuestos(c).Refresh
    Loop
    
    'leer ranking.tbr y cargar los temas que haya
    Dim TE As TextStream
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        FSO.CreateTextFile AP + "ranking.tbr", True
    End If
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    Dim TT As String
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    c = 0
    Do While Not TE.AtEndOfStream
        TT = TE.ReadLine
        ThisPTS = Val(txtInLista(TT, 0, ","))
        ThisArch = txtInLista(TT, 1, ",")
        ThisTEMA = txtInLista(TT, 2, ",")
        ThisDISCO = txtInLista(TT, 3, ",")
            
        If c = MaxTop Then Exit Do
        'si elarchivo no existe no se debe cargar
        If FSO.FileExists(ThisArch) Then
            lblPuestos(c).UseMnemonic = False
            lblPuestos(c) = " " + Trim(Str(c + 1)) + "º " + _
            QuitarNumeroDeTema(ThisTEMA) + " / " + ThisDISCO + " [" + Trim(Str(ThisPTS)) + " pts]"
            lblPuestos(c).Refresh
            
            c = c + 1
            ReDim Preserve MTXtop(c)
            MTXtop(c) = ThisArch
            ReDim Preserve MTXtemas(c)
            MTXtemas(c) = ThisTEMA
            ReDim Preserve MTXdiscos(c)
            MTXdiscos(c) = ThisDISCO
        End If
    Loop
    'élegir el primero
    lblPuestos(0).BackColor = ColorSel
    AjustarFRM Me, 12000
    FR.Left = Screen.Width / 2 - FR.Width / 2
    FR.Top = Screen.Height / 2 - FR.Height / 2
    
End Sub

