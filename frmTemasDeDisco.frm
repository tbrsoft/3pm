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
         Left            =   30
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   7050
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
         TabIndex        =   9
         Top             =   7620
         Width           =   4515
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OK"
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
            TabIndex        =   6
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
         Height          =   1605
         IntegralHeight  =   0   'False
         ItemData        =   "frmTemasDeDisco.frx":2C8A
         Left            =   8010
         List            =   "frmTemasDeDisco.frx":2C9D
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   4905
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.ListBox lstTIME 
         BackColor       =   &H00404080&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   6555
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
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   6555
         IntegralHeight  =   0   'False
         Left            =   1260
         Sorted          =   -1  'True
         TabIndex        =   0
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
         TabIndex        =   15
         Top             =   8040
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
         TabIndex        =   14
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
         Height          =   495
         Left            =   30
         TabIndex        =   13
         Top             =   8340
         Width           =   7065
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
Dim TeclaBajo As Long 'codigo de la tecla que se detecto en KDown para usar en KeyUP
Dim SegSinTecla As Long
Dim NoHayTemasEnDisco As Boolean
Dim DuracionTema As String
Dim YaInicio As Long
'0=load
'1=keyup
'hasta que no haga un keyUp no da bola a ejecutar tema!!!!!!!

Private Sub cmdDiscoAd_Click()
    If IsMod46Teclas = 46 Then
        'el evento set focus no puee ponerse si es que el formulario salio
        ' o sea si ya n existe
        'en el caso de 5 teclas las teclas del costado SALEN!
        'y el set focus da error
        Form_KeyDown TeclaDER, 0
        'entonces solo lo hago si estoy el el modo 4 o6. NO EN EL 5!
        Command1.SetFocus
    End If
    If IsMod46Teclas = 5 Then
        'el evento set focus no puee ponerse si es que el formulario salio
        ' o sea si ya n existe
        'en el caso de 5 teclas las teclas del costado SALEN!
        'y el set focus da error
        Form_KeyDown TeclaPagAd, 0
        'entonces solo lo hago si estoy el el modo 4 o6. NO EN EL 5!
        'Command1.SetFocus
    End If
End Sub

Private Sub cmdDiscoAd_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub cmdDiscoAt_Click()
    'si esta en el modo de 5 teclas se debe simular los botones de pag adel y atras
    'ya que estas son las que realmente pasan los discos
    
    If IsMod46Teclas = 46 Then
        'el evento set focus no puee ponerse si es que el formulario salio
        ' o sea si ya n existe
        'en el caso de 5 teclas las teclas del costado SALEN!
        'y el set focus da error
        Form_KeyDown TeclaIZQ, 0
        'entonces solo lo hago si estoy el el modo 4 o6. NO EN EL 5!
        Command1.SetFocus
    End If
    
    If IsMod46Teclas = 5 Then
        'el evento set focus no puee ponerse si es que el formulario salio
        ' o sea si ya n existe
        'en el caso de 5 teclas las teclas del costado SALEN!
        'y el set focus da error
        Form_KeyDown TeclaPagAt, 0
        'entonces solo lo hago si estoy el el modo 4 o6. NO EN EL 5!
        'Command1.SetFocus
    End If
    
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
    tERR.Anotar "000-0024"
    MostrarCursor False
    'actualizar los precios
    
    '---------------------
    'si es gratis no usar!
    'actualizar los precios
    '---------------------
    'si es gratis no usar!
    If CreditosCuestaTema(0) = 0 Then
        lblPrecios = "Musica Gratis"
    Else
        lblPrecios = "1 cancion = " + CStr(FormatCurrency(PrecioBase * CreditosCuestaTema(0), , , , vbFalse))
        If CreditosCuestaTema(1) > 0 Then
        lblPrecios = lblPrecios + " / 2 canciones = " + CStr(FormatCurrency(PrecioBase * CreditosCuestaTema(1), , , , vbFalse))
        End If
        
        If CreditosCuestaTema(2) > 0 Then
            lblPrecios = lblPrecios + " / 3 canciones = " + CStr(FormatCurrency(PrecioBase * CreditosCuestaTema(2), , , , vbFalse))
        End If
    End If
    
    'si es gratis no usar!
    If CreditosCuestaTemaVIDEO(0) = 0 Then
        lblPrecios = lblPrecios + " / Videos Gratis"
    Else
        lblPrecios = lblPrecios + " / 1 video = " + CStr(FormatCurrency(CreditosCuestaTemaVIDEO(0) * PrecioBase, , , , vbFalse))
        
        If CreditosCuestaTemaVIDEO(1) > 0 Then
            lblPrecios = lblPrecios + " / 2 videos = " + CStr(FormatCurrency(CreditosCuestaTemaVIDEO(1) * PrecioBase, , , , vbFalse))
        End If
        
        If CreditosCuestaTemaVIDEO(2) > 0 Then
            lblPrecios = lblPrecios + " / 3 videos = " + CStr(FormatCurrency(PrecioBase * CreditosCuestaTemaVIDEO(2), , , , vbFalse))
        End If
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Errores
    'y si no es una ficha la que se esta cargando
    lblNoEjecuta.Visible = False
    
    'la verdadera tecla debe mostrar si es una tecla del teclado numerico
    Dim RealKeyCode As Integer
    
    'de manera predeterminada son el mismo
    'salvo los casos que se especifican
    RealKeyCode = KeyCode
    
    If IsKeyPad(Me) Then
        'lasa falla reconocida por microsoft es de la tecla enter
        'sea cual sea sale el keycode 13 por mas que sea la del keypad
        'que es el 108
        If KeyCode = 13 Then
            RealKeyCode = 108
        End If
        'ademas si esta apretado el BLOQ NUM
    End If
    TeclaBajo = RealKeyCode 'en Kdown anda mejor   ¿¿¿¿¿¿¿¿¿¿porque?????
    'lblCOMOSALIR = CStr(KeyCode) + "-" + CStr(RealKeyCode)
    
    Select Case RealKeyCode
        Case TeclaNewFicha
            If FindParam3PM("to") = "kd" Then
                LTE 1
                If CREDITOS <= MaximoFichas Then
                    'apagar el fichero electronico
                    SetKeyState vbKeyScrollLock, True
                    VarCreditos CSng(TemasPorCredito)
                Else
                    'apagar el fichero electronico
                    SetKeyState vbKeyScrollLock, False
                End If
            End If
        Case TeclaNewFicha2
            If FindParam3PM("to2") = "kd" Then
                LTE 2
                If CREDITOS <= MaximoFichas Then
                    'apagar el fichero electronico
                    SetKeyState vbKeyScrollLock, True
                    VarCreditos CSng(CreditosBilletes)
                Else
                    'apagar el fichero electronico
                    SetKeyState vbKeyScrollLock, False
                End If
            End If
        Case vbKeyF4
            If Shift = 4 Then
                Unload Me
            End If
        Case TeclaShowContador
            frmOnlyContador.Show 1
        Case TeclaCerrarSistema
            SetKeyState vbKeyCapital, False
            MostrarCursor True
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            frmIndex.MP3.DoClose 99
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
                If lstTEMAS.ListIndex < lstTEMAS.ListCount - 1 Then
                    lstTEMAS.ListIndex = lstTEMAS.ListIndex + 1
                Else
                    lstTEMAS.ListIndex = 0
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
                If lstTEMAS.ListIndex > 0 Then
                    lstTEMAS.ListIndex = lstTEMAS.ListIndex - 1
                Else
                    lstTEMAS.ListIndex = lstTEMAS.ListCount - 1
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
            'si esta en 5 teclas y ademas eligio la flecha de direccion abajo o arriba
            'falla por que el listbox recibe y se mueve!
            'KeyCode 40 = Flecha abajo
            'KeyCode 38 = Flecha Arriba
            If IsMod46Teclas = 5 Then
                'igual que el boton adelante!!
                If lstTEMAS.ListIndex < lstTEMAS.ListCount - 1 Then
                    'si es 40 ya va a bajar de todas formas!
                    'se evita la duplicacion!
                    If TeclaPagAd <> 40 Then lstTEMAS.ListIndex = lstTEMAS.ListIndex + 1
                Else
                    'en este caso si es 40 ira al 0 y luego llegari el movimiento hacia abajo
                    'porque primero yo y despues el funcionamiento propio del lstBox ¿?
                    'NO LO SE!. Se queda si la vuelta
                    If TeclaPagAd <> 40 Then lstTEMAS.ListIndex = 0
                End If
                SaltarEspaciosLstTemas True
            End If
            TECLAS_PRES = TECLAS_PRES + "5"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            frmIndex.lblTECLAS = TECLAS_PRES
        Case TeclaPagAt
            If IsMod46Teclas = 5 Then
                'igual que el boton atras!!
                If lstTEMAS.ListIndex > 0 Then
                    'ver comentario TeclaPagAd!!!
                    '!!!!!!!!!!!!!
                    If TeclaPagAt <> 38 Then lstTEMAS.ListIndex = lstTEMAS.ListIndex - 1
                Else
                    'en este caso si es 40 ira al 0 y luego llegari el movimiento hacia abajo
                    'porque primero yo y despues el funcionamiento propio del lstBox ¿?
                    'NO LO SE!. Se queda si la vuelta
                    If TeclaPagAt <> 38 Then lstTEMAS.ListIndex = lstTEMAS.ListCount - 1
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
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acps"
    Resume Next
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Local Error GoTo FallaKD
    
    'y si no es una ficha la que se esta cargando
    'aqui se regsitran las presiones de las teclas elegidas
    tERR.Anotar "000-0033", KeyCode
    
    'la verdadera tecla debe mostrar si es una tecla del teclado numerico
    Dim RealKeyCode As Integer
    'ver si es o no numpad
    If IsKeyPad(Me) Then
        'la falla reconocida por microsoft es de la tecla enter
        'sea cual sea sale el keycode 13 por mas que sea la del keypad
        'que es el 108
        RealKeyCode = KeyCode
        If KeyCode = 13 Then RealKeyCode = 108
        'ademas si esta apretado el BLOQ NUM
    Else
        'de manera predeterminada son el mismo
        'salvo los casos que se especifican
        RealKeyCode = KeyCode
    End If
    
    If TeclaBajo = 108 Then RealKeyCode = 108
      
    'ver detalle mas abajo de que mierda es esto y en el gral de este frm
    YaInicio = YaInicio + 1
    'puede no escuchar el coin!!!!!!
    'esto se pone mas abajo!!!!
    'If YaInicio <= 1 Then Exit Sub
        
    
    Select Case RealKeyCode
        Case TeclaNewFicha
            If FindParam3PM("to") = "999999" Then
                LTE 1
                If CREDITOS <= MaximoFichas Then
                    'apagar el fichero electronico
                    SetKeyState vbKeyScrollLock, True
                    VarCreditos CSng(TemasPorCredito)
                Else
                    'apagar el fichero electronico
                    SetKeyState vbKeyScrollLock, False
                End If
            End If
        Case TeclaNewFicha2
            If FindParam3PM("to2") = "999999" Then
                LTE 2
                If CREDITOS <= MaximoFichas Then
                    'apagar el fichero electronico
                    SetKeyState vbKeyScrollLock, True
                    VarCreditos CSng(CreditosBilletes)
                Else
                    'apagar el fichero electronico
                    SetKeyState vbKeyScrollLock, False
                End If
            End If
        
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
            temaElegido = lstEXT.List(lstTEMAS.ListIndex) ' UbicDiscoActual + "\" + lstTemas + "." + EXTs(lstTemas.ListIndex)
            
            If LCase(Right(temaElegido, 3)) = "mp3" Or LCase(Right(temaElegido, 3)) = "wma" Then '''Or LCase(Right(temaElegido, 3)) = "mp4" Then
                PideVideo = False
            Else
                PideVideo = True
            End If
            
            'ver si puede pagar lo que pide!!!
            'que joyita papa!!!. Parece que supieras programar
'            '--------------------------------------------------------------
'            If (PideVideo = False And CREDITOS >= CreditosCuestaTema(0)) Or _
'                (PideVideo And CREDITOS >= CreditosCuestaTemaVIDEO(0)) Then
'            '--------------------------------------------------------------
            '--------------------------------------------------------------
            If (PideVideo = False And CREDITOS >= PrecNowAudio) Or _
                (PideVideo And CREDITOS >= PrecNowVideo) Then
            '--------------------------------------------------------------

                'siempre que se ejecute un credito estaremos por debajo de maximo
                SetKeyState vbKeyScrollLock, True
                                
                'restar lo que corresponde!!!
                If PideVideo Then
                    VarCreditos -PrecNowVideo
                Else
                    'VarCreditos -CreditosCuestaTema(0)
                    VarCreditos -PrecNowAudio
                End If
                                
                'si esta ejecutando pasa a la lista de reproducción
                'si esta ejecutando una prueba SACARLA!!!
                'el 99 pregunta si cualquier cosa se esta ejecutando!!
                If (frmIndex.MP3.IsPlaying(0) Or frmIndex.MP3.IsPlaying(1)) And CORTAR_TEMA = False Then
                    'pasar a la lista de reproducción
                    Dim NewIndLista As Long
                    NewIndLista = UBound(MATRIZ_LISTA)
                    ReDim Preserve MATRIZ_LISTA(NewIndLista + 1)
                    'se graba en Matriz_Listas como patah, nombre(sin .mp3)
                    MATRIZ_LISTA(NewIndLista + 1) = temaElegido + "," + lstTEMAS + " / " + FSO.GetBaseName(UbicDiscoActual)
                    CargarProximosTemas
                    'graba en reini.tbr los datos que correspondan por si se corta la luz
                    CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
                    'AHORA DEBE MARCARLO COMO EJECUTADO Y SALIR PARA ELIJA OTRO
                    lstAgregados = lstAgregados + lstTEMAS.List(lstTEMAS.ListIndex) + " / "
                    
                    If BloquearMusicaElegida Then
                        lstTEMAS.List(lstTEMAS.ListIndex) = "----------"
                        lstTIME.List(lstTIME.ListIndex) = "---"
                    End If
                        
                    lstAgregados.Visible = True
                    lstTEMAS.Height = lstAgregados.Top - lstTEMAS.Top
                    lstTIME.Height = lstAgregados.Top - lstTIME.Top
                    SaltarEspaciosLstTemas True
                    
                    If OutTemasWhenSel Then Unload Me
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
                    lstAgregados = lstAgregados + lstTEMAS.List(lstTEMAS.ListIndex) + " / "
                    
                    If BloquearMusicaElegida Then
                        lstTEMAS.List(lstTEMAS.ListIndex) = "----------"
                        lstTIME.List(lstTIME.ListIndex) = "---"
                    End If
                    
                    lstAgregados.Visible = True
                    lstTEMAS.Height = lstAgregados.Top - lstTEMAS.Top
                    lstTIME.Height = lstAgregados.Top - lstTIME.Top
                    SaltarEspaciosLstTemas True
                    CORTAR_TEMA = False 'este tema va entero ya que lo eligio el usuario
                    Me.ZOrder
                    EjecutarTema temaElegido, True
                    'si es un video y sale en el monitor de la PC _
                        salir para verlo!!!
                    If Salida2 = False Then Unload Me
                End If
                
                VerSiTocaPUB
                'dejo seguir eligiendo y no salgo!!!
                'SI ESTA CONFIGURADO ASI (6.5)!!!
                If OutTemasWhenSel Then Unload Me
            Else
                lblNoEjecuta.Visible = True
            End If
    End Select
    
Exit Sub

FallaKD:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acpt"
    Resume Next
    
    
End Sub

Private Sub Form_Load()
    YaInicio = 0
    If Is3pmExclusivo Then
        lstTEMAS.BackColor = vbBlack
        lstTIME.BackColor = vbBlack
        lstTEMAS.ForeColor = vbYellow
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

    'es una matriz global
    UbicDiscoActual = txtInLista(MATRIZ_DISCOS(nDiscoGral), 0, ",")
    
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
    Dim c As Integer, nombreTemas As String
    Dim pathTema As String
    lstEXT.Clear
    lstTEMAS.Clear
    If NoHayTemasEnDisco Then
        lstTEMAS.AddItem "No hay temas en este disco"
        lstTEMAS.Enabled = False
        lstTIME.Enabled = False
        tERR.AppendLog "No hay temas en el disco: " + UbicDiscoActual + ".acpu"
        Exit Sub
    End If
    c = 1
    Dim EXT As String
    Do While c <= UBound(MATRIZ_TEMAS)
        pathTema = txtInLista(MATRIZ_TEMAS(c), 0, "#")
        nombreTemas = txtInLista(MATRIZ_TEMAS(c), 1, "#")
        EXT = LCase(txtInLista(nombreTemas, 1, "."))
        'quitar el molesto .mp3 o lo que fuera
        Select Case LCase(EXT)
            Case "mp3"
                EXT = " (mp3-Musica)"
'            Case "mp4"
'                EXT = " (mp4-Musica)"
            Case "wma"
                EXT = " (wma-Musica)"
            Case "mpeg", "mpg", "avi", "wmv"
                EXT = " (" + LCase(EXT) + "-Video)"
            Case "dat"
                EXT = " (dat-VCD-Video)"
        End Select
        nombreTemas = FSO.GetBaseName(nombreTemas) + EXT
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
            'XXXX no se si podra leeer la duracion del mp4 igual que el mp3
            If UCase(Right(pathTema, 3)) = "MP3" Then '''Or UCase(Right(pathTema, 3)) = "MP4" Then
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

Private Sub lstTemas_Click()
    On Local Error Resume Next
    If CargarDuracionTemas Then lstTIME.ListIndex = lstTEMAS.ListIndex
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
    Ahora = lstTEMAS.ListIndex
    
    Dim nINI As Long, nFin As Long, StepMio As Long
    If HaciaAdelante Then
        nINI = Ahora
        nFin = lstTEMAS.ListCount - 1
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
        If lstTEMAS.List(A) <> "----------" Then
            'ya esta lo encontro!!!!!!!
            'ir ahi!!!
            lstTEMAS.ListIndex = A
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
                    nINI = lstTEMAS.ListCount - 1
                    GoTo ReiniLST
                End If
            End If
        End If
    Next
End Sub
