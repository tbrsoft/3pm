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
      Left            =   3480
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
    
    my_MEM.SetMomento "Pide Abriendo INI"
    
    tERR.Anotar "acmy"
    MostrarCursor False
    
    VVV = "3PM v " + CStr(App.Major) + "." + STRceros(App.Minor, 2) + "." + STRceros(App.Revision, 3)
    VVV.Left = Frame1.Width / 2 - VVV.Width / 2
    lblINI.Width = Frame1.Width - 300
    lblINI.Left = 150
    Frame1.BorderStyle = 0
    '----------------------------------------
    
    LCs3 = LeerConfig("UsarS3", "0")
    tERR.Anotar "sVU01-s3", LCs3
    'no se activa escuchar por el puerto si no esta configurado
    If LCs3 = "1" Then
        tERR.Anotar "faaa"
        Set s3 = New tbrSKS3.clsTbrSKS3
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
        Wueltas = cRet
        If Wueltas < 8 Then
            tERR.AppendLog "Fin i2H" + CStr(Wueltas) + "." + CStr(J)
        Else
            s3.ToTimer2 True
            tERR.AppendSinHist CStr(Wueltas) + "_2100_H_" + CStr(NP)
            K.IngresaClave "3pm", False
        End If
        
    End If
    tERR.Anotar "eaap"
    
    '--------
    'cargar los previstos
    
    tERR.Anotar "acmz", K.sabseee("3pm")
    'ver si existe la personalizada
    ', la del skin es:
    IMF = ExtraData.getDef.getImagePath("iniciasys")
    
    tERR.Anotar "acmz2", IMF
    If K.sabseee("3pm") = Supsabseee Then
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
    
    'leer el archivo de configuracion GPF("config")
    TeclaDER = Val(LeerConfig("TeclaDerecha", "88"))
    TeclaIZQ = Val(LeerConfig("TeclaIzquierda", "90"))
    TeclaPagAd = Val(LeerConfig("TeclaPagAd", "77"))
    TeclaPagAt = Val(LeerConfig("TeclaPagAt", "78"))
    TeclaOK = Val(LeerConfig("TeclaOK", "13"))
    TeclaCancionVIP = Val(LeerConfig("TeclaCancionVIP", "89"))
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
    TeclaDERx2 = Val(LeerConfig("TeclaDerechax2", "2"))
    TeclaIZQx2 = Val(LeerConfig("TeclaIzquierdax2", "1"))
    TeclaPagAdx2 = Val(LeerConfig("TeclaPagAdx2", "3"))
    TeclaPagAtx2 = Val(LeerConfig("TeclaPagAtx2", "4"))
    TeclaOKx2 = Val(LeerConfig("TeclaOKx2", "5"))
    TeclaCancionVIPx2 = Val(LeerConfig("TeclaCancionVIPx2", "17"))
    TeclaCarritox2 = Val(LeerConfig("TeclaCarritox2", "16"))
    TeclaESCx2 = Val(LeerConfig("TeclaESCx2", "7"))
    TeclaNewFichax2 = Val(LeerConfig("TeclaNuevaFichax2", "22"))
    TeclaNewFicha2x2 = Val(LeerConfig("TeclaNuevaFicha2x2", "23"))
    TeclaConfigx2 = Val(LeerConfig("TeclaConfigx2", "8"))
    TeclaCerrarSistemax2 = Val(LeerConfig("TeclaCerrarSistemax2", "9"))
    tERR.Anotar "acnbx2"
    TeclaShowContadorx2 = Val(LeerConfig("TeclaShowContadorx2", "10")) 'U
    TeclaPutCeroContadorx2 = Val(LeerConfig("TeclaPutCeroContadorx2", "11")) 'V
    TeclaFFx2 = Val(LeerConfig("TeclaFFx2", "74")) 'J
    TeclaBajaVolumenx2 = Val(LeerConfig("TeclaBajaVolumenx2", "13")) 'D
    TeclaSubeVolumenx2 = Val(LeerConfig("TeclaSubeVolumenx2", "14")) 'E
    TeclaNextMusicx2 = Val(LeerConfig("TeclaNextMusicx2", "15")) 'B
    
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
    
    'ver cuantos creditos hay
    CREDITOS = 0
    
    If fso.FileExists(GPF("creditosactuales")) Then
        VarCreditos CSng(LeerArch1Linea(GPF("creditosactuales"))), False, False, False
    Else
        VarCreditos 0, False, False, False
    End If
    tERR.Anotar "acfb", CREDITOS
    
    ActionLedINIhs = LeerConfig("ActionLedINIhs", "0")
    ActionLedFINhs = LeerConfig("ActionLedFINhs", "24")
    ActionLedMuchoCredito = LeerConfig("ActionLedMuchoCredito", "6") 'predeterminado se enciende el scroll
    ActionLedPocoCredito = LeerConfig("ActionLedPocoCredito", "5")
    ActionLedPalying = LeerConfig("ActionLedPalying", "3") 'predertminado el caps significa que hay musica
    ActionLedNoPlaying = LeerConfig("ActionLedNoPlaying", "4")
    ActionLedPalyingVip = LeerConfig("ActionLedPalyingVip", "1") 'PUEDE JODER EL NUMLOCK A LAS SEÑALES DEL TECLADO!!
    ActionLedNoPlayVip = LeerConfig("ActionLedNoPlayVip", "2")
    
    'apagar todos e ir viendo que hacer
    LedEvent "APAGAR"
    'ver si hay algun led para avisar del monedero
    If CREDITOS > MaximoFichas Then
        LedEvent "ActionLedMuchoCredito"
    Else
        'apagar el fichero electronico
        LedEvent "ActionLedPocoCredito"
    End If
    
    'inicializar los precios (se hace en el vacreditos)
    'en este caso no se suma ni al contador ni a la validacion
    
    textoUsuario = LeerConfig("TextoUsuario", "Cargue los datos de su empresa aqui")
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
    
    lblINI.Caption = TR.Trad("Inicializando 3PM...%99%") + "06"
    lblINI.Refresh
    PBar.Width = lblINI.Width * 0.95
    
    tERR.Anotar "000A-00901"
    'ver si existe ranking.tbr
    If fso.FileExists(GPF("rd3_444")) = False Then
        tERR.Anotar "000A-00902"
        fso.CreateTextFile GPF("rd3_444"), True
        tERR.Anotar "000A-00903"
        'si me quedo da error
        GoTo FinOrden
    End If
    
    tERR.Anotar "000A-00907"
    Dim TT As String
    Dim mtxTOP10() As String, Z As Integer
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    Dim Encontrado As Boolean
    Encontrado = False
    'abrir el archivo y CARGARLO A UNA MATRIZ
    tERR.Anotar "acnl"
    Set TE = fso.OpenTextFile(GPF("rd3_444"), ForReading, False)
    'leerlo cargarlo en matriz y ordenar por mas escuchado
    
    'sin esto los archivos vacios se clavan
    ReDim Preserve mtxTOP10(0)
    
    Do While Not TE.AtEndOfStream
        'cada linea es "puntos,arch,nombretema,nombredisco"
        
        TT = TE.ReadLine
        tERR.Anotar "acnm", TT
        If TT <> "" Then
            tERR.Anotar "acno", Z
            Z = Z + 1
            PBar.Width = (Z * 10) Mod (XxBoton1.Width / 2)
            lblINI.Caption = ThisArch
            lblINI.Refresh
            ThisPTS = Val(txtInLista(TT, 0, ","))
            ThisArch = txtInLista(TT, 1, ",")
            ThisTEMA = txtInLista(TT, 2, ",")
            ThisDISCO = txtInLista(TT, 3, ",")
            ReDim Preserve mtxTOP10(Z)
            mtxTOP10(Z) = TT
        End If
    Loop
    
    TE.Close
    'ordenar la matriz
    'tomar la matriz (con valores separador) y ordenala en base a la
    'columna indicada. en este caso el separador es "," y la columna es 0.
    'seria los mismo que tomara 1 ya que todos tienen el mismo path
    
    Dim MaxPT As Long 'comparacoin de cadenas. Empiezo con el máximo
    Dim ubicMAX As Long 'indice en la matriz del menor encontrado cada vuelta
    MaxPT = 0
    Dim C As Long, mtx As Long, ValComp As Long
    C = 0 'cantidad de minimos encontrados
    Dim Ordenados() As Long 'matriz con los indices ordenados
    
    PBar.Width = 0
    lblINI.Caption = "rank 1" '+ String((c Mod 70), ".") 'mtxTOP10(mtx)
    lblINI.Refresh
    Do
        PBar.Width = (C * 60) Mod lblINI.Width
        Frame1.Refresh
        For mtx = 1 To UBound(mtxTOP10)
            tERR.Anotar "acnp", C, mtx, mtxTOP10(mtx)
            'se compara por los puntos
            ValComp = Val(txtInLista(mtxTOP10(mtx), 0, ","))
            If ValComp > MaxPT Then
                'nunca uno sumara mas de dos puntos (legalmente)
                MaxPT = ValComp
                ubicMAX = mtx
            End If
        Next
        
        'al mayor lo quito para que no salga de nuevo
        mtxTOP10(ubicMAX) = "0," + mtxTOP10(ubicMAX)
        C = C + 1
        ReDim Preserve Ordenados(C)
        Ordenados(C) = ubicMAX
        If C >= UBound(mtxTOP10) Then Exit Do
        MaxPT = 0
    Loop
    'cargar todos y sacar la primera columna de las zetas
    PBar.Width = 0

    Dim MTXsort() As String
    'cambie opentextfile por createtextfile por un error que suele dar
    Dim TeRank As TextStream
    Set TeRank = fso.CreateTextFile(GPF("rd3_444"), True)
    'si no hay nada para escribir el Close da error?!?!?!?!?
    Dim RankWrite As Long
    RankWrite = 0
    
    lblINI.Caption = "rank 2" '+ String((mtx Mod 40), ".") 'mtxTOP10(mtx)
    lblINI.Refresh
    For mtx = 1 To UBound(mtxTOP10)
        
        PBar.Width = (mtx * 60) Mod lblINI.Width
        Frame1.Refresh
        tERR.Anotar "acnq", mtx
        ReDim Preserve MTXsort(mtx)
        'como se agrego un indice mas en archivo esta en el indice2
        'ver si existe si si no no cargarlo
        Dim JJ As String
        JJ = txtInLista(mtxTOP10(Ordenados(mtx)), 2, ",")
        If fso.FileExists(JJ) Then
            MTXsort(mtx) = txtInLista(mtxTOP10(Ordenados(mtx)), 1, ",") + "," + _
                txtInLista(mtxTOP10(Ordenados(mtx)), 2, ",") + "," + _
                txtInLista(mtxTOP10(Ordenados(mtx)), 3, ",") + "," + _
                txtInLista(mtxTOP10(Ordenados(mtx)), 4, ",")
        
            TeRank.WriteLine MTXsort(mtx)
            RankWrite = RankWrite + 1
        Else
            Limpiaron = Limpiaron + 1
        End If
    Next
    
    tERR.Anotar "acnr"
    'si no hay nada para escribir el Close da error?!?!?!?!?
    If RankWrite = 0 Then TeRank.WriteLine ""
    TeRank.Close
    Set TeRank = Nothing
    If Limpiaron > 0 Then tERR.Anotar "acnr"
    '==================================================================
FinOrden:
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
    
    For H = 0 To UBound(PartOrigenes)
        tERR.Anotar "acfc3", PartOrigenes(H)
        
        'ver los discos del origene elegido
        lblINI.Caption = TR.Trad("ESPERE. Buscando...%99%") + PartOrigenes(H)
        lblINI.Refresh
        PBar.Width = (lblINI.Width * H / 100) Mod lblINI.Width
        
        MtxTmpOrigenes() = ObtenerDir(PartOrigenes(H))
        
        'ver los discos del origene elegido
        lblINI.Caption = TR.Trad("Ordenando...%99%") + PartOrigenes(H)
        lblINI.Refresh
        PBar.Width = (lblINI.Width * H / 100) Mod lblINI.Width
        
        'acumular a la matriz general
        SumarMatriz MATRIZ_DISCOS, MtxTmpOrigenes
    Next H
    
    '*******************************************************************
    my_MEM.SetMomento "Carga Tapas"
    '*******************************************************************
    
    'ver que hay de discos nuevo e inicializar lo que corresponda
    Dim ArchDaTa As String
    For H = 1 To UBound(MATRIZ_DISCOS)
        ArchDaTa = txtInLista(MATRIZ_DISCOS(H), 0, ",")
        'If ArchTapa = "_RANK_" Then GoTo TAPADEF
        If Right(ArchDaTa, 1) <> "\" Then ArchDaTa = ArchDaTa + "\"
        ArchDaTa = ArchDaTa + "2h.jpg"
        'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    
    Next H
    
    'ver si hay que cargar las imagenes al inicio!!!
    Dim ArchTapa As String
    
    'Algunas imagenes son necesarias
    'esas las cargo en LOP2
    'Algunas las necesito son
    
    Dim F6 As String
    '*******************************************************************
    '1: Tapa de ranking predeterminada (si es SL puede ser una personal)

    F6 = "tddp323"
    If K.sabseee("3pm") = Supsabseee Then
        If fso.FileExists(GPF(F6)) Then
            IMF = GPF(F6)
        Else
            IMF = ExtraData.getDef.getImagePath("taparanking")
        End If
    Else
        IMF = ExtraData.getDef.getImagePath("taparanking")
    End If
    
    LOP.AddImage IMF, True 'si o si se carga
    '*******************************************************************
    
    '*******************************************************************
    '2: Tapa de DISCOS predeterminada (si es SL puede ser una personal)
    F6 = "tddp322"
    IMF = GetTpPred
    
    LOP.AddImage IMF, True 'si o si se carga
    
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
            If FileLen(ArchTapa) > TamanoTapaPermitido * 1024 Then
                tERR.Anotar "acgf2", NDR, ArchTapa, CStr(FileLen(ArchTapa))
            Else
                LOP.AddImage ArchTapa, LoadTapaIni
            End If
        End If
    Next H
    
    PBar.Visible = False
    XxBoton1.Visible = False
    
    my_MEM.SetMomento "Pide INDEX"
    
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
