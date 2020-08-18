VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmINI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10530
   Icon            =   "frmINI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox ts3INI 
      Height          =   465
      Left            =   2760
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   1425
      Left            =   1680
      TabIndex        =   0
      Top             =   2850
      Width           =   7590
      Begin tbrFaroButton.fBoton pBAR 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   990
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   582
         fFColor         =   12632319
         fBColor         =   0
         fCapt           =   ""
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   16777215
      End
      Begin tbrFaroButton.fBoton XxBoton1 
         Height          =   390
         Left            =   90
         TabIndex        =   3
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   688
         fFColor         =   6553600
         fBColor         =   0
         fCapt           =   ""
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   16711680
      End
      Begin VB.Label lblINI 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Contando: 00"
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
         Height          =   495
         Left            =   90
         TabIndex        =   2
         Top             =   480
         Width           =   7410
      End
      Begin VB.Label VVV 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "versión"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   1185
      End
   End
   Begin VB.Image Image1 
      Height          =   1305
      Left            =   3510
      Top             =   1350
      Width           =   1440
   End
End
Attribute VB_Name = "frmINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Traducir 'Agregado por el complemento traductor
    
    On Error GoTo MiErr
    
    my_MEM.SetMomento "0088"
    
    tERR.Anotar "acmy"
    MostrarCursor False
    
    
    If MDCN2 > 0 Then
        VVV.ForeColor = vbYellow
    End If
    
    'VVV = "3PM v " + CStr(App.Major) + "." + STRceros(App.Minor, 2) + "." + STRceros(App.Revision, 3)
    VVV = dcr("S/7qOI2TQKjGUTi9076V2/a/tnd6I+PA") + CStr(App.Major) + "." + STRceros(App.Minor, 2) + "." + STRceros(App.Revision, 3)
    
    VVV.Left = Frame1.Width / 2 - VVV.Width / 2
    lblINI.Width = Frame1.Width - 300
    lblINI.Left = 150
    Frame1.BorderStyle = 0
    '----------------------------------------
    
    LCs3 = LeerConfig("UsarS3", "0")
    tERR.Anotar "sVU01-s3", LCs3
    'no se activa escuchar por el puerto si no esta configurado
    If LCs3 = "1" Then
        my_MEM.SetMomento "0089"
        tERR.Anotar "faaa"
        Set s3 = New tbrSKS3.clsTbrSKS3
        
        'si hay que indicar otros puertos es aca !!!
        'el lpt comun es
        's3.setPorts &H378, &H379, &H37A ' 888 889 890
        'ejemplo de lpt2
        's3.setPorts &H278, &H279, &H27A ' 632 633 634
        'ejemplo de pci+usb como en pc que mosse llevo a USA
        's3.setPorts &HB050, &HB051, &HB052 'sera 45136, 45137 y 45138
        
        Dim Ports(2) As Integer
        Ports(0) = CInt("&H" + LeerConfig("LptPort0", "378"))
        Ports(1) = CInt("&H" + LeerConfig("LptPort1", "379"))
        Ports(2) = CInt("&H" + LeerConfig("LptPort2", "37A"))
        
        s3.setPorts Ports(0), Ports(1), Ports(2)
        s3.INIT
        
        '*************************
        s3.HwndMsg = ts3INI.HWND
        
        'tERR.AppendLog "S3_1:" + CStr(txtS3.HWND)
        
        s3.ReIniCounters 'son los mios
        tERR.Anotar "faab"
        s3.Prender
        s3.SetInterval CLng(LeerConfig("FrecTecladoTBR", "50"))
        
        tERR.Anotar "faac"
        s3.Prender
        
        esperar 1
        
        s3.ReIniContLuis
        tERR.Anotar "faad"
        esperar 1
        
        'obtener el numero de placa
        NP = CLng(s3.GetnPlaca(SYSfolder + "prec.dll"))
        tERR.Anotar "faae" + CStr(NP)
        If NP = -1 Then
            'mLog "No se podido comenzar la prueba. Quizas la interfase no este conectada o sea una versión solo botones"
            'Exit Sub
        End If
        
        'VER LICENCIA!!!!!
        Dim TimOut As Single  'TimeOut
        TimOut = 2
        
        Dim J As Long, RET As Long, cRet As Long
        cRet = 0
        For J = 1 To 10
            RET = s3.AddCont(J Mod 4, TimOut)
            tERR.Anotar "faaf", J, RET
            If RET = 2 Then 'time out
                tERR.Anotar "***** TIME OUT - CONT:" + CStr(J) + " (timeout!) " + s3.GetResLicSTR
            End If
            
            If RET = 1 Then 'llego mal!!! poner en cero
                tERR.Anotar "***** MAL CONT:" + CStr(J) + " (bad!) " + s3.GetResLicSTR
                'reinicio todo!!!
                s3.ReIniContLuis
                esperar 1
            End If
            
            If RET = 0 Then
                tERR.Anotar "CONT:" + CStr(J) + " (ok!) " + s3.GetResLicSTR
                cRet = cRet + 1 'veces que esta ok
            End If
            
            If J > 2 And cRet = 0 Then
                Wueltas = 0
                Exit For
            End If
        Next J
        
        tERR.Anotar "faag", cRet
        '*************************
        my_MEM.SetMomento "0090 " + CStr(Wueltas)
        Wueltas = cRet
        If Wueltas < 8 Then
            tERR.AppendLog "Fin i2H" + CStr(Wueltas) + "." + CStr(J)
        Else
            s3.ToTimer2 True
            tERR.AppendSinHist CStr(Wueltas) + "_2100_H_" + CStr(NP)
            'para que la licencia sea valida va a buscar la lista de placas validas para este software
            K.IngresaClave dcr("q44KmdDBQ+IB8dTOX8F+VA=="), False
            
        End If
        
    End If
    tERR.Anotar "eaap"
    
    '--------
    'cargar los previstos
    my_MEM.SetMomento "0081"
    tERR.Anotar "acmz", K.sabseee(dcr("q44KmdDBQ+IB8dTOX8F+VA=="))
    'ver si existe la personalizada
    ', la del skin es:
    IMF = ExtraData.getDef.getImagePath("iniciasys")
    
    tERR.Anotar "acmz2", IMF
    If K.sabseee(dcr("1Vx0YVGhEoIisHPLAZMHXw==")) >= Supsabseee Then
        If fso.FileExists(GPF("iisl67")) Then
            tERR.Anotar "acmz3"
            Image1.Picture = LoadPicture(GPF("iisl67"))
            frmVIDEO.picBigImg = LoadPicture(GPF("iisl67"))
        Else
            tERR.Anotar "acmz3A"
            Image1.Picture = LoadPicture(IMF)
            frmVIDEO.picBigImg = LoadPicture(IMF)
        End If
    Else
        tERR.Anotar "acmz3B"
        Image1.Picture = LoadPicture(IMF)
        frmVIDEO.picBigImg = LoadPicture(IMF)
    End If
    
    my_MEM.SetMomento "0091"
    tERR.Anotar "acmz4"
    Image1.Left = Screen.Width / 2 - Image1.Width / 2
    Image1.Top = 300 'Me.Height / 2 - Image1.Height / 2
    'Frame1.Width = 5500
    Frame1.Left = Screen.Width / 2 - Frame1.Width / 2
    tERR.Anotar "acmz5", Screen.Height, Image1.Height, Image1.Top, Frame1.Left
    Frame1.Top = Image1.Top + Image1.Height + 300
    Frame1.Height = Screen.Height - Image1.Height - 1600
    PBar.Left = lblINI.Left
    tERR.Anotar "acmz6"
    XxBoton1.Left = PBar.Left - 15
    XxBoton1.Width = lblINI.Width
    tERR.Anotar "acmz7"
    frmVIDEO.picBigImg.Top = frmVIDEO.Height / 2 - frmVIDEO.picBigImg.Height / 2
    frmVIDEO.picBigImg.Left = frmVIDEO.Width / 2 - frmVIDEO.picBigImg.Width / 2
    frmVIDEO.picBigImg.Visible = True
    frmVIDEO.picBigImg.Refresh
    
    tERR.Anotar "000A-00903"
    lblINI.Caption = TR.Trad("Inicializando 3PM...%98%Al arranque del sistema%99%")
    lblINI.Refresh
    PBar.Width = 0
    
    tERR.Anotar "acna"
    
    Me.Show
    Me.Refresh
    
    my_MEM.SetMomento "0092"
    '////////////////////////////////////////////////////////
    'leer el archivo de configuracion GPF("config")
    TeclaDER = Val(LeerConfig("TeclaDerecha", "88"))
    TeclaIZQ = Val(LeerConfig("TeclaIzquierda", "90"))
    TeclaPagAd = Val(LeerConfig("TeclaPagAd", "77"))
    TeclaPagAt = Val(LeerConfig("TeclaPagAt", "78"))
    TeclaOK = Val(LeerConfig("TeclaOK", "13"))
    TeclaCancionVIP = Val(LeerConfig("TeclaCancionVIP", "89"))
    teclaSumValidar = Val(LeerConfig("teclaSumValidar", "80"))
    TeclaCarrito = Val(LeerConfig("TeclaCarrito", "79"))
    TeclaESC = Val(LeerConfig("TeclaESC", "27"))
    TeclaNewFicha = Val(LeerConfig("TeclaNuevaFicha", "81"))
    TeclaNewFicha2 = Val(LeerConfig("TeclaNuevaFicha2", "83"))
    TeclaConfig = Val(LeerConfig("TeclaConfig", "67"))
    TeclaCerrarSistema = Val(LeerConfig("TeclaCerrarSistema", "87"))
    tERR.Anotar "acnb"
    TeclaShowContador = Val(LeerConfig("TeclaShowContador", "85")) 'U
    TeclaPutCeroContador = Val(LeerConfig("TeclaPutCeroContador", "86")) 'V
    TeclaFF = Val(LeerConfig("TeclaFF", "74")) 'J
    TeclaBajaVolumen = Val(LeerConfig("TeclaBajaVolumen", "68")) 'D
    TeclaSubeVolumen = Val(LeerConfig("TeclaSubeVolumen", "69")) 'E
    TeclaNextMusic = Val(LeerConfig("TeclaNextMusic", "66")) 'B
    lblINI.Caption = TR.Trad("Inicializando 3PM...%99%") + "01"
    lblINI.Refresh
    PBar.Width = lblINI.Width * 0.15
    '////////////////////////////////////////////////////////
    TeclaDERx2 = Val(LeerConfig("TeclaDerechax2", "2"))
    TeclaIZQx2 = Val(LeerConfig("TeclaIzquierdax2", "1"))
    TeclaPagAdx2 = Val(LeerConfig("TeclaPagAdx2", "3"))
    TeclaPagAtx2 = Val(LeerConfig("TeclaPagAtx2", "4"))
    TeclaOKx2 = Val(LeerConfig("TeclaOKx2", "5"))
    TeclaCancionVIPx2 = Val(LeerConfig("TeclaCancionVIPx2", "17"))
    teclaSumValidarX2 = Val(LeerConfig("teclaSumValidarX2", "18"))
    TeclaCarritox2 = Val(LeerConfig("TeclaCarritox2", "16"))
    TeclaESCx2 = Val(LeerConfig("TeclaESCx2", "7"))
    TeclaNewFichax2 = Val(LeerConfig("TeclaNuevaFichax2", "22"))
    TeclaNewFicha2x2 = Val(LeerConfig("TeclaNuevaFicha2x2", "23"))
    TeclaConfigx2 = Val(LeerConfig("TeclaConfigx2", "8"))
    TeclaCerrarSistemax2 = Val(LeerConfig("TeclaCerrarSistemax2", "9"))
    tERR.Anotar "acnbx2"
    TeclaShowContadorx2 = Val(LeerConfig("TeclaShowContadorx2", "10")) 'U
    TeclaPutCeroContadorx2 = Val(LeerConfig("TeclaPutCeroContadorx2", "11")) 'V
    TeclaFFx2 = Val(LeerConfig("TeclaFFx2", "12")) 'J
    TeclaBajaVolumenx2 = Val(LeerConfig("TeclaBajaVolumenx2", "13")) 'D
    TeclaSubeVolumenx2 = Val(LeerConfig("TeclaSubeVolumenx2", "14")) 'E
    TeclaNextMusicx2 = Val(LeerConfig("TeclaNextMusicx2", "15")) 'B
    '////////////////////////////////////////////////////////
    ShowCreditsMode = Val(LeerConfig("ShowCreditsMode", "0"))
    ApagarAlCierre = LeerConfig("ApagarAlCierre", "0")
    'puede ser 46 o 5 por ahora
    IsMod46Teclas = CLng(LeerConfig("IsMod46Teclas", "46"))
    
    lblINI.Caption = TR.Trad("Inicializando 3PM...%99%") + "02"
    lblINI.Refresh
    PBar.Width = lblINI.Width * 0.45
    
    MaximoFichas = Val(LeerConfig("MaximoFichas", "40"))
    EsperaMinutos = Val(LeerConfig("EsperaMinutos", "900"))
    'Valores de ReIni FULL=tema ejecutando y lista LISTA=solo lista NADA=arranca de cero
    ReINI = LeerConfig("ReINI", "LISTA")
    tERR.Anotar "acnd"
    VolumenIni = CLng(LeerConfig("Volumen", "50"))
    tERR.Anotar "acnd2", VolumenIni
    VolumenIni2 = CLng(LeerConfig("Volumen2", "50"))
    tERR.Anotar "acnd3", VolumenIni2
    EsperaTecla = Val(LeerConfig("EsperaTecla", "900"))
    tERR.Anotar "acnd4", EsperaTecla
    PorcentajeTEMA = Val(LeerConfig("PorcentajeTema", "60"))
    tERR.Anotar "acnd5", PorcentajeTEMA
    FASTini = LeerConfig("FastIni", "1")
    tERR.Anotar "acnd6", FASTini
    HabilitarVUMetro = LeerConfig("HabilitarVUMetro", "0")
    LoadTapaIni = LeerConfig("LoadTapaIni", "0")
    tERR.Anotar "acnd7", HabilitarVUMetro
    vidFullScreen = LeerConfig("VidFullScreen", "1")
    QuitaBarraSup = LeerConfig("QuitaBarraSup", "0")
    QuitaBarraInf = LeerConfig("QuitaBarraInf", "0")
    
    tERR.Anotar "acnd8", vidFullScreen
    Salida2 = LeerConfig("Salida2", "0")
    tERR.Anotar "acnd9", Salida2
    NoVumVID = LeerConfig("NoVumVid", "1")
    tERR.Anotar "acnd10", NoVumVID
    OutTemasWhenSel = LeerConfig("OutTemasWhenSel", "0")
    tERR.Anotar "acnd11", OutTemasWhenSel
    BloquearMusicaElegida = LeerConfig("BloquearMusicaElegida", "1")
    tERR.Anotar "acnd12", BloquearMusicaElegida
    TapasMostradasH = Val(LeerConfig("DiscosH", "3"))
    tERR.Anotar "acnd13", TapasMostradasH
    TapasMostradasV = Val(LeerConfig("DiscosV", "2"))
    tERR.Anotar "acnd14", TapasMostradasV
    PasarHoja = LeerConfig("Pasarhoja", "1")
    tERR.Anotar "acnd15", PasarHoja
    DistorcionarTapas = LeerConfig("DistorcionarTapas", "0")
    tERR.Anotar "acnd16", DistorcionarTapas
    Protector = LeerConfig("Protector", "1")
    tERR.Anotar "acne"
    CreditForTestMusic = CLng(LeerConfig("CreditForTestMusic", "0"))
    MaxListaTestMusic = CLng(LeerConfig("MaxListaTestMusic", "0"))
    MaxMuestrasToAddCredit = CLng(LeerConfig("MaxMuestrasToAddCredit", "0"))
    
    lblINI.Caption = TR.Trad("Inicializando 3PM...%99%") + "03"
    lblINI.Refresh
    PBar.Width = lblINI.Width * 0.67
    
    CargarDuracionTemas = LeerConfig("CargarDuracionTemas", "0")
    MostrarRotulos = LeerConfig("MostrarRotulos", "1")
    RotulosArriba = LeerConfig("RotulosArriba", "0")
    DuracionProtect = LeerConfig("DuracionProtect", "180")
    TemasPorCredito = LeerConfig("TemasPorCredito", "1")
    CreditosBilletes = LeerConfig("CreditosBilletes", "10")
    PrecioBase = CSng(LeerConfig("PrecioBase", "0,50"))
    
    'ver si hay que mostrar el touch
    MostrarTouch = LeerConfig("MostrarTouch", "1")
    
    CreditosCuestaTema(0) = LeerConfig("CreditosCuestaTema", "1")
    CreditosCuestaTema(1) = LeerConfig("CreditosCuestaTema2", "2")
    CreditosCuestaTema(2) = LeerConfig("CreditosCuestaTema3", "3")
    'upManu
    CreditosXaVipMusica = LeerConfig("CreditosXaVipMusica", "0") 'predeterminado desactivado
    PrecNowVIP = CreditosXaVipMusica
    
    CreditosCuestaTemaVIDEO(0) = LeerConfig("CreditosCuestaTemaVIDEO", "2")
    CreditosCuestaTemaVIDEO(1) = LeerConfig("CreditosCuestaTemaVIDEO2", "3")
    CreditosCuestaTemaVIDEO(2) = LeerConfig("CreditosCuestaTemaVIDEO3", "4")
    
    my_MEM.SetMomento "0093"
    
    'ver cuantos creditos hay
    CREDITOS = 0
    
    If fso.FileExists(GPF("creditosactuales")) Then
        VarCreditos CSng(LeerArch1Linea(GPF("creditosactuales"))), False, False, False
    Else
        VarCreditos 0, False, False, False
    End If
    tERR.Anotar "acfb", CREDITOS
    
    ActionLedOn = LeerConfig("ActionLedOn", "0")
    ActionLedINIhs = LeerConfig("ActionLedINIhs", "0")
    ActionLedFINhs = LeerConfig("ActionLedFINhs", "24")
    ActionLedMuchoCredito = LeerConfig("ActionLedMuchoCredito", "6") 'predeterminado se enciende el scroll
    ActionLedPocoCredito = LeerConfig("ActionLedPocoCredito", "5")
    ActionLedPalying = LeerConfig("ActionLedPalying", "3") 'predertminado el caps significa que hay musica
    ActionLedNoPlaying = LeerConfig("ActionLedNoPlaying", "4")
    ActionLedPalyingVip = LeerConfig("ActionLedPalyingVip", "1") 'PUEDE JODER EL NUMLOCK A LAS SEÑALES DEL TECLADO!!
    ActionLedNoPlayVip = LeerConfig("ActionLedNoPlayVip", "2")
    
    GrabaKar = LeerConfig("GrabaKar", "0")
    KbpsKar = LeerConfig("KbpsKar", "128")
    GrabaKarQuick = LeerConfig("GrabaKarQuick", "1")
    'solo se gasta memoria si se va a usar!
    If GrabaKar > 0 Then
        Set TW10 = New tbrWRII.tbrWR2
        TW10.SetFileLog AP + "logWII.log"
        tERR.AppendSinHist "tbrWII:INICIA"
        TW10.Dispositivo = 0 'elegir solo para que pueda hacer el log ok
        'registro para ver que placas tiene
        TW10.LogDispositivos
        TW10.LogLineas
    End If
    
    'apagar todos e ir viendo que hacer
    LedEvent "APAGAR"
    'ver si hay algun led para avisar del monedero
    If MaximoFichas > 0 And CREDITOS > MaximoFichas Then
        LedEvent "ActionLedMuchoCredito"
    Else
        'apagar el fichero electronico
        LedEvent "ActionLedPocoCredito"
    End If
    
    'inicializar los precios (se hace en el vacreditos)
    'en este caso no se suma ni al contador ni a la validacion
    
    Select Case MDCN2
        Case 0 'sin crack!!
            '"Cargue los datos de su empresa aqui"
            textoUsuario = LeerConfig("TextoUsuario", dcr("5GjFL+wevJFr9DU1lnH/7nSaQCLVZw2otsEccz+r5aC76h8ofTCt3JRLLAfYLpMX"))
        Case 1 'crack en dic 09
            ''www.tbrsoft.com + _ + "Cargue los datos de su empresa aqui"
            textoUsuario = dcr("IbsGyqFDkye5yuiUGhHbZuivOdq7zrNEfHLeIVdNTQg=") + vbCrLf + _
                LeerConfig("TextoUsuario", dcr("5GjFL+wevJFr9DU1lnH/7nSaQCLVZw2otsEccz+r5aC76h8ofTCt3JRLLAfYLpMX"))
                
        Case 2 'crack en ene 10
            'www.tbrsoft.com
            textoUsuario = dcr("IbsGyqFDkye5yuiUGhHbZuivOdq7zrNEfHLeIVdNTQg=")
            
            'rockolas@peru.com
            'textoUsuario = dcr("1Cb7nxQ9JnbzmNaUZr8iUueU4qQn8Te62N17nMcETKg=")
            
        Case 3 'crack en feb 10
            'musica y video gratuitos !!!
            textoUsuario = dcr("ExRo0fQ3SYU9tOSZ3T0CyLvKLe9WYKG0ortTb8Xr6MmjuG7ShLwmdA==")
        
        Case 4 'crack en marzo 09
            'info@tbrsoft.com
            textoUsuario = dcr("lDheXOYXh8HqejJ6XrxpIF+PISeSC0OOmXedRCZwBHc=")
    End Select
    
    
    textoUsuario = Replace(textoUsuario, Chr(5), vbCrLf)
    
    lblINI.Caption = TR.Trad("Inicializando 3PM...%99%") + "04"
    lblINI.Refresh
    PBar.Width = lblINI.Width * 0.78
    
    tERR.Anotar "acnf"
    'publicidad
    'la cargo si o si para que si despues entra a la conficuracion ya este cargada
    PUBs.CargarPUBs
    'inicializar publicidades si corresponde
    PUBs.HabilitarPublicidadesMp3Vid = LeerConfig("MostrarPub", "0")
    PUBs.HabilitarPublicidadesVMute = LeerConfig("MostrarPUBMute", "0")
    PUBs.SonarPublicidadesCada = LeerConfig("PubliCada", "5")
    PUBs.HabilitarPublicidadesIMG = LeerConfig("MostrarPubIMG", "0")
    PUBs.SonarPublicidadesIMGCada = LeerConfig("PubliIMGCada", "10")
    
    'cargar variables de claves
    'archivo de claves
    If fso.FileExists(GPF("sequeda32")) = False Then
        Set TE = fso.CreateTextFile(GPF("sequeda32"), True)
        TE.WriteLine "Config:12345612345612345612"
        TE.WriteLine "Close:45612345612345612345"
        TE.WriteLine "Credit:1234441234441234561"
        TE.Close
    End If
    
    tERR.Anotar "acnh"
    Set TE = fso.OpenTextFile(GPF("sequeda32"), ForReading, False)
    'config/close/credit es el orden del archivo
    ClaveConfig = txtInLista(TE.ReadLine, 1, ":")
    ClaveClose = txtInLista(TE.ReadLine, 1, ":")
    ClaveCredit = txtInLista(TE.ReadLine, 1, ":")
    TE.Close
    
    lblINI.Caption = TR.Trad("Inicializando 3PM...%99%") + "05"
    lblINI.Refresh
    PBar.Width = lblINI.Width * 0.84
    
    tERR.Anotar "acni"
    'ver si ya estaba cargado
    If App.PrevInstance Then
        MsgBox TR.Trad("No se pueden abrir dos instancias " + _
            "de 3pm%98%Instancia se refiere cada copia del programa en ejecución" + vbCrLf + _
            "Esto se hace por que si 3PM esta cargado en memoria no tiene " + _
            "sentido que se cargue de nuevo. Puede haber fallas%99%"): End
    End If
    
    my_MEM.SetMomento "0094"
    'ASEGURARSE QUE EXISTA la carpeta del ranking y la imagen que le corresponde
    If fso.FolderExists(AP + "discos") = False Then
        fso.CreateFolder AP + "discos"
    End If
    
    tERR.Anotar "acnj"
    
    'ver si es superlicencia y usa otra tapa predeterminada
    IMF = GetTpPred
    
    If fso.FileExists(IMF) = False Then
        MsgBox TR.Trad("No se encuentra el archivo de imagen de las " + _
            "portadas predeterminadas!. " + vbCrLf + _
            "La instalacion de 3PM no es corecta!%98%" + _
            "Con portadas se refiere a la tapa de los discos%99%")
        End
    End If
    
    'carpeta del protector
    If fso.FolderExists(AP + "fotos") = False Then
        fso.CreateFolder AP + "fotos"
    End If
    
    TECLAS_PRES = "11111222223333344444" 'arranca con 20 pulsaciones
    
    
    '===================ORDENAR EL RANKING================================
    my_MEM.SetMomento "0096"
    lblINI.Caption = TR.Trad("Inicializando 3PM...%99%") + "06"
    lblINI.Refresh
    PBar.Width = lblINI.Width * 0.95
    
    'ordenar el ranking
    srtRNK
    'borrale si esta crakeado
    DelFrmRank
    '==================================================================
    my_MEM.SetMomento "0098"
    'se inicializa el contador para que la variable CONTADOR tenga el
    'valor de todas las fichas cargadas
    'si este es cero esta en los primeros usos entonces mostrar el CLUF
    tERR.Anotar "acns"
    SumarContadorCreditos 0
    SumarContadorCarrito 0
    
    Dim MtxTmpOrigenes() As String
    Dim Origenes As String
    Origenes = LeerArch1Linea(GPF("origs"))
    PartOrigenes = Split(Origenes, "*")

    Dim H As Long
    
    'la carga real de origenes es el for que viene despues, esto solo en el caso excepcional que se graben karaokes
    If GrabaKar > 0 Then
        'asegurarse que haya una carpeta u origen donde caiga todo esto!
        FolKarSave = AP + "Karaokes grabados"
        If fso.FolderExists(FolKarSave) = False Then fso.CreateFolder FolKarSave
        'si recien empeiza el dia puede ser que se vaya a grabar pero aun no este la carpeta, asegurarme que se cree!!
        FolKarSaveNAU = FolKarSave + "\" + "GRABACIONES " + _
            STRceros(Day(Date), 2) + "-" + STRceros(Month(Date), 2) + "-" + STRceros(Year(Date), 4)
        If fso.FolderExists(FolKarSaveNAU) = False Then fso.CreateFolder FolKarSaveNAU
        
        tERR.Anotar "acfc3k", FolKarSave, FolKarSaveNAU
                
        Origenes = FolKarSave + "*" + LeerArch1Linea(GPF("origs"))
        PartOrigenes = Split(Origenes, "*")
    End If
    my_MEM.SetMomento "0099"
    Dim ResumenIniDiscos As String
    ResumenIniDiscos = ""
    For H = 0 To UBound(PartOrigenes)
        tERR.Anotar "acfc3", PartOrigenes(H)
        
        'ver los discos del origene elegido
        lblINI.Caption = TR.Trad("ESPERE. Buscando...%99%") + PartOrigenes(H)
        lblINI.Refresh
        PBar.Width = (lblINI.Width * H / 100) Mod lblINI.Width
        
        MtxTmpOrigenes() = ObtenerDir(PartOrigenes(H))
        ResumenIniDiscos = ResumenIniDiscos + PartOrigenes(H) + ": " + CStr(UBound(MtxTmpOrigenes)) + vbCrLf
        'ver los discos del origene elegido
        lblINI.Caption = TR.Trad("Ordenando...%99%") + PartOrigenes(H)
        lblINI.Refresh
        PBar.Width = (lblINI.Width * H / 100) Mod lblINI.Width
        
        'acumular a la matriz general
        SumarMatriz MATRIZ_DISCOS, MtxTmpOrigenes
    Next H
    tERR.AppendSinHist "HINIDSC:" + vbCrLf + ResumenIniDiscos
    '*******************************************************************
    my_MEM.SetMomento "0100"
    '*******************************************************************
    
    'ver que hay de discos nuevo e inicializar lo que corresponda
    'en cada disco debe hacer un archivo que indique en que fecha se agrego
    'de esta forma se cuanto se escucha en promedio cada disco y si un disco no se ha escuchado
    'para al automatizar el ingreso de musica tambien se haga con el egreso de musica
    
    Dim ArchDaTa As String
    For H = 1 To UBound(MATRIZ_DISCOS)
        ArchDaTa = txtInLista(MATRIZ_DISCOS(H), 0, ",")
        'If ArchTapa = "_RANK_" Then GoTo TAPADEF
        If Right(ArchDaTa, 1) <> "\" Then ArchDaTa = ArchDaTa + "\"
        ArchDaTa = ArchDaTa + "2h.dt"
        
        If fso.FileExists(ArchDaTa) = False Then
            tERR.Anotar "acfc3m", ArchDaTa, "SI"
            Dim TxR As TextStream
            Set TxR = fso.CreateTextFile(ArchDaTa, True)
                Dim dateCreate As Long
                dateCreate = CLng(Date)
                TxR.Write "create " + CStr(dateCreate)
                'VER SI HACE FALTA MAS DEJE ACA!!!
                'MM889
            TxR.Close
        Else
            tERR.Anotar "acfc3m", ArchDaTa, "NO"
        End If
    Next H
    'my_MEM.SetMomento "0101"
    'ver si hay que cargar las imagenes al inicio!!!
    Dim ArchTapa As String
    
    'Algunas imagenes son necesarias
    'esas las cargo en LOP2
    'Algunas las necesito son
    
    Dim F6 As String
    '*******************************************************************
    '1: Tapa de ranking predeterminada (si es SL puede ser una personal)

    F6 = "tddp323"
    If K.sabseee(dcr("1Vx0YVGhEoIisHPLAZMHXw==")) >= Supsabseee Then
        If fso.FileExists(GPF(F6)) Then
            IMF = GPF(F6)
        Else
            IMF = ExtraData.getDef.getImagePath("taparanking")
        End If
    Else
        IMF = ExtraData.getDef.getImagePath("taparanking")
    End If
    tERR.Anotar "acfc3n", IMF
    LOP.AddImage IMF, True 'si o si se carga
    '*******************************************************************
    
    '*******************************************************************
    '2: Tapa de DISCOS predeterminada (si es SL puede ser una personal)
    F6 = "tddp322"
    IMF = GetTpPred
    tERR.Anotar "acfc3p", IMF
    LOP.AddImage IMF, True 'si o si se carga
    my_MEM.SetMomento "0103"
    '*******************************************************************
    'TODOS LOS DEMAS!!!!
    For H = 1 To UBound(MATRIZ_DISCOS)
        ArchTapa = txtInLista(MATRIZ_DISCOS(H), 0, ",")
        'If ArchTapa = "_RANK_" Then GoTo TAPADEF
        If Right(ArchTapa, 1) <> "\" Then ArchTapa = ArchTapa + "\"
        ArchTapa = ArchTapa + "tapa.jpg"
        
        lblINI.Caption = TR.Trad("Buscando...%99%") + ArchTapa
        lblINI.Refresh
        PBar.Width = (lblINI.Width * H / 100) Mod lblINI.Width
        
        'solo la cargo si existe y ademas tiene el tamaño que tiene que tener
        If fso.FileExists(ArchTapa) Then
            'si la tapa es demasiado grande
            tERR.Anotar "acfc3q", ArchTapa
            If FileLen(ArchTapa) > TamanoTapaPermitido * 1024 Then
                tERR.Anotar "acgf2", NDR, ArchTapa, CStr(FileLen(ArchTapa))
            Else
                LOP.AddImage ArchTapa, LoadTapaIni
            End If
        End If
    Next H
    my_MEM.SetMomento "0104"
    PBar.Visible = False
    XxBoton1.Visible = False
    
    my_MEM.SetMomento "0084"
    
    lblINI.Caption = TR.Trad("Abriendo 3PM ... %99%")
    lblINI.Refresh
    
    
    frmIndex.Show 1
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acnt"
    Resume Next
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    tERR.Anotar "ACNC19"
    VVV.Caption = TR.Trad("versión%98%Se refiere a la version numerica del software%99%")
End Sub

Private Sub ts3INI_Change()
    tERR.Anotar "ACNC20", ts3INI.tExt
End Sub
