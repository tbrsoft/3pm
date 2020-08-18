VERSION 5.00
Begin VB.Form frmINI 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmINI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ListBox lblPROCES 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1590
      IntegralHeight  =   0   'False
      Left            =   1320
      TabIndex        =   4
      Top             =   7320
      Width           =   9015
   End
   Begin VB.Label VVV 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   450
      Left            =   1380
      TabIndex        =   3
      Top             =   6420
      Width           =   120
   End
   Begin VB.Label lblTipoLIC 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciando SUPERLICENCIA"
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
      Height          =   285
      Left            =   6930
      TabIndex        =   2
      Top             =   150
      Width           =   3270
   End
   Begin VB.Label pBar 
      BackColor       =   &H00C0FFFF&
      Height          =   90
      Left            =   1380
      TabIndex        =   1
      Top             =   7140
      Width           =   435
   End
   Begin VB.Label lblINI 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Contando Discos: 00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   1350
      TabIndex        =   0
      Top             =   6930
      Width           =   8970
   End
   Begin VB.Image Image1 
      Height          =   6825
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   90
      Width           =   9000
   End
End
Attribute VB_Name = "frmINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    On Error GoTo MiErr
    tERR.Anotar "acmy"
    MostrarCursor False
    
    VVV = "3PM v " + Trim(CStr(App.Major)) + "." + Trim(CStr(App.Minor)) + "." + Trim(CStr(App.Revision))
    '--------
    'cargar los previstos

    tERR.Anotar "acmz", K.LICENCIA
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(GPF("iisl67")) Then
            Image1.Picture = LoadPicture(GPF("iisl67"))
            frmVIDEO.picBigImg = LoadPicture(GPF("iisl67"))
        Else
            Image1.Picture = LoadPicture(GPF("extr233_52"))
            frmVIDEO.picBigImg = LoadPicture(GPF("extr233_52"))
        End If
    Else
        Image1.Picture = LoadPicture(GPF("extr233_52"))
        frmVIDEO.picBigImg = LoadPicture(GPF("extr233_52"))
    End If
    frmVIDEO.picBigImg.Top = frmVIDEO.Height / 2 - frmVIDEO.picBigImg.Height / 2
    frmVIDEO.picBigImg.Left = frmVIDEO.Width / 2 - frmVIDEO.picBigImg.Width / 2
    frmVIDEO.picBigImg.Visible = True
    frmVIDEO.picBigImg.Refresh
    '--------
    Select Case K.LICENCIA
        Case HSuperLicencia
            lblTipoLIC = "Iniciando SUPERLICENCIA"
        Case GFull
            lblTipoLIC = "Iniciando Licencia Full"
        Case CGratuita
            lblTipoLIC = "Iniciando Demo gratuito"
        Case aSinCargar
            lblTipoLIC = "Iniciando Demo 3PM"
    End Select
    lblTipoLIC.Refresh
    tERR.Anotar "acna"
    AjustarFRM Me, 12000
    'leer el archivo de configuracion GPF("config")
    CargarIMGinicio = LeerConfig("CargarImagenInicio", "1")
    AutoReDibuj = LeerConfig("AutoReDraw", "1")
    TeclaDER = Val(LeerConfig("TeclaDerecha", "88"))
    TeclaIZQ = Val(LeerConfig("TeclaIzquierda", "90"))
    TeclaPagAd = Val(LeerConfig("TeclaPagAd", "77"))
    TeclaPagAt = Val(LeerConfig("TeclaPagAt", "78"))
    TeclaOK = Val(LeerConfig("TeclaOK", "13"))
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
    
    TeclaDERx2 = Val(LeerConfig("TeclaDerechax2", "2"))
    TeclaIZQx2 = Val(LeerConfig("TeclaIzquierdax2", "1"))
    TeclaPagAdx2 = Val(LeerConfig("TeclaPagAdx2", "3"))
    TeclaPagAtx2 = Val(LeerConfig("TeclaPagAtx2", "4"))
    TeclaOKx2 = Val(LeerConfig("TeclaOKx2", "5"))
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
    HabilitarVUMetro = LeerConfig("HabilitarVUMetro", "1")
    tERR.Anotar "acnd7", HabilitarVUMetro
    vidFullScreen = LeerConfig("VidFullScreen", "1")
    tERR.Anotar "acnd8", vidFullScreen
    Salida2 = LeerConfig("Salida2", "0")
    tERR.Anotar "acnd9", Salida2
    NoVumVID = LeerConfig("NoVumVid", "0")
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
    CargarDuracionTemas = LeerConfig("CargarDuracionTemas", "0")
    MostrarRotulos = LeerConfig("MostrarRotulos", "1")
    RotulosArriba = LeerConfig("RotulosArriba", "0")
    DuracionProtect = LeerConfig("DuracionProtect", "180")
    RankToPeople = LeerConfig("RankToPeople", "1")
    TemasPorCredito = LeerConfig("TemasPorCredito", "1")
    CreditosBilletes = LeerConfig("CreditosBilletes", "10")
    PrecioBase = CSng(LeerConfig("PrecioBase", "0,50"))
    
    'ver si hay que mostrar el touch
    MostrarTouch = LeerConfig("MostrarTouch", "0")
    
    CreditosCuestaTema(0) = LeerConfig("CreditosCuestaTema", "1")
    CreditosCuestaTema(1) = LeerConfig("CreditosCuestaTema2", "2")
    CreditosCuestaTema(2) = LeerConfig("CreditosCuestaTema3", "3")
    
    CreditosCuestaTemaVIDEO(0) = LeerConfig("CreditosCuestaTemaVIDEO", "2")
    CreditosCuestaTemaVIDEO(1) = LeerConfig("CreditosCuestaTemaVIDEO2", "3")
    CreditosCuestaTemaVIDEO(2) = LeerConfig("CreditosCuestaTemaVIDEO3", "4")
    
    'ver cuantos creditos hay
    CREDITOS = 0
    
    If FSO.FileExists(GPF("creditosactuales")) Then
        VarCreditos CSng(LeerArch1Linea(GPF("creditosactuales"))), False, False, False
    Else
        VarCreditos 0, False, False, False
    End If
    tERR.Anotar "acfb", CREDITOS
    
    'inicializar los precios (se hace en el vacreditos)
    'en este caso no se suma ni al contador ni a la validacion
    
    textoUsuario = LeerConfig("TextoUsuario", "Cargue los datos de su empresa aqui")
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
    
    tERR.Anotar "acng"
    IDIOMA = LeerConfig("Idioma", "Español")
    
    'cargar variables de claves
    'archivo de claves
    If FSO.FileExists(GPF("sequeda32")) = False Then
        Set TE = FSO.CreateTextFile(GPF("sequeda32"), True)
        TE.WriteLine "Config:12345612345612345612"
        TE.WriteLine "Close:45612345612345612345"
        TE.WriteLine "Credit:1234441234441234561"
        TE.Close
    End If
    tERR.Anotar "acnh"
    Set TE = FSO.OpenTextFile(GPF("sequeda32"), ForReading, False)
    'config/close/credit es el orden del archivo
    ClaveConfig = txtInLista(TE.ReadLine, 1, ":")
    ClaveClose = txtInLista(TE.ReadLine, 1, ":")
    ClaveCredit = txtInLista(TE.ReadLine, 1, ":")
    TE.Close
    
    Me.Show
    Me.Refresh
    
    tERR.Anotar "acni"
    'ver si ya estaba cargado
    If App.PrevInstance Then MsgBox "No se pueden abrir dos instancias de 3pm": End
    
    'ASEGURARSE QUE EXISTA la carpeta del ranking y la imagen que le corresponde
    If FSO.FolderExists(AP + "discos") = False Then
        FSO.CreateFolder AP + "discos"
    End If
    tERR.Anotar "acni2"
    If FSO.FolderExists(AP + "discos\01- Los mas escuchados") = False Then
        FSO.CreateFolder AP + "discos\01- Los mas escuchados"
     End If
     tERR.Anotar "acni3", HabilitarVUMetro, NoVumVID
    'siempre copiarlo, si es el SL con prioridad
    If FSO.FileExists(GPF("233_54_b")) Then
        'aqui hay un error de acceso denegado si es de solo lectura!!!!!
        'se corrige así.
        tERR.Anotar "acni4"
        FSO.CopyFile GPF("233_54_b"), AP + "discos\01- Los mas escuchados\tapa.jpg", True
    Else
        If FSO.FileExists(GPF("extr233_54")) Then
            tERR.Anotar "acni5"
            'aqui hay un error de acceso denegado si es de solo lectura!!!!!
            'se corrige así.
            FSO.CopyFile GPF("extr233_54"), AP + "discos\01- Los mas escuchados\tapa.jpg", True
        Else
            tERR.Anotar "acni6"
            MsgBox "No se encuentra el archivo de imagen del Ranking!. La instalacion de 3PM no es corecta!"
            End
        End If
    End If
    
    tERR.Anotar "acnj"
    If FSO.FileExists(GPF("extr233_61")) = False Then
        MsgBox "No se encuentra el archivo de imagen de las portadas predeterminadas!. La instalacion de 3PM no es corecta!"
        End
    End If
    'carpeta del protector
    If FSO.FolderExists(AP + "fotos") = False Then
        FSO.CreateFolder AP + "fotos"
    End If
    
    TECLAS_PRES = "11111222223333344444" 'arranca con 20 pulsaciones
    
    
    '===================ORDENAR EL RANKING================================
    
    tERR.Anotar "000A-00901"
    'ver si existe ranking.tbr
    If FSO.FileExists(GPF("rd3_444")) = False Then
        tERR.Anotar "000A-00902"
        FSO.CreateTextFile GPF("rd3_444"), True
        tERR.Anotar "000A-00903"
        'si me quedo da error
        GoTo FinOrden
    End If
    tERR.Anotar "acnk"
    tERR.Anotar "000A-00903"
    lblINI.Caption = "Inicializando 3PM..."
    tERR.Anotar "000A-00904"
    lblINI.Refresh
    tERR.Anotar "000A-00905"
    PBar.Width = 0
    tERR.Anotar "000A-00906"
    PBar.Refresh
    tERR.Anotar "000A-00907"
    Dim TT As String
    Dim mtxTOP10() As String, z As Integer
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    Dim Encontrado As Boolean
    Encontrado = False
    'abrir el archivo y CARGARLO A UNA MATRIZ
    tERR.Anotar "acnl"
    Set TE = FSO.OpenTextFile(GPF("rd3_444"), ForReading, False)
    'leerlo cargarlo en matriz y ordenar por mas escuchado
    
    'sin esto los archivos vacios se clavan
    ReDim Preserve mtxTOP10(0)
    
    Do While Not TE.AtEndOfStream
        'cada linea es "puntos,arch,nombretema,nombredisco"
        
        TT = TE.ReadLine
        tERR.Anotar "acnm", TT
        If TT <> "" Then
            tERR.Anotar "acno", z
            z = z + 1
            PBar.Width = (z * 10) Mod (lblPROCES.Width / 2)
            'If PBar.Width > lblPROCES.Width Then PBar.Width = 100
            ThisPTS = Val(txtInLista(TT, 0, ","))
            ThisArch = txtInLista(TT, 1, ",")
            ThisTEMA = txtInLista(TT, 2, ",")
            ThisDISCO = txtInLista(TT, 3, ",")
            ReDim Preserve mtxTOP10(z)
            mtxTOP10(z) = TT
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
    Dim c As Long, mtx As Long, ValComp As Long
    c = 0 'cantidad de minimos encontrados
    Dim Ordenados() As Long 'matriz con los indices ordenados
    
    PBar.Width = 0
    PBar.Refresh
    
    Do
        PBar.Width = c * 10
        
        For mtx = 1 To UBound(mtxTOP10)
            tERR.Anotar "acnp", c, mtx, mtxTOP10(mtx)
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
        c = c + 1
        ReDim Preserve Ordenados(c)
        Ordenados(c) = ubicMAX
        If c >= UBound(mtxTOP10) Then Exit Do
        MaxPT = 0
    Loop
    'cargar todos y sacar la primera columna de las zetas
    PBar.Width = 0
    PBar.Refresh
    Dim MTXsort() As String
    'cambie opentextfile por createtextfile por un error que suele dar
    Dim TeRank As TextStream
    Set TeRank = FSO.CreateTextFile(GPF("rd3_444"), True)
    'si no hay nada para escribir el Close da error?!?!?!?!?
    Dim RankWrite As Long
    RankWrite = 0
    
    For mtx = 1 To UBound(mtxTOP10)
        tERR.Anotar "acnq", mtx
        ReDim Preserve MTXsort(mtx)
        'como se agrego un indice mas en archivo esta en el indice2
        'ver si existe si si no no cargarlo
        If Dir(txtInLista(mtxTOP10(Ordenados(mtx)), 2, ",")) <> "" Then
            MTXsort(mtx) = txtInLista(mtxTOP10(Ordenados(mtx)), 1, ",") + "," + _
                txtInLista(mtxTOP10(Ordenados(mtx)), 2, ",") + "," + _
                txtInLista(mtxTOP10(Ordenados(mtx)), 3, ",") + "," + _
                txtInLista(mtxTOP10(Ordenados(mtx)), 4, ",")
        
            TeRank.WriteLine MTXsort(mtx)
            PBar.Width = mtx * 10
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
    frmIndex.Show 1
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acnt"
    Resume Next
End Sub

Private Sub lblPROCES_Click()
    If lblPROCES.ListIndex = -1 Then Exit Sub
    lblPROCES.ListIndex = lblPROCES.ListCount - 1
End Sub

