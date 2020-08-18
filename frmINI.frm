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
   Begin VB.Label VVV 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   1560
      TabIndex        =   4
      Top             =   30
      Width           =   4875
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   7020
      TabIndex        =   3
      Top             =   7080
      Width           =   3270
   End
   Begin VB.Label pBar 
      BackColor       =   &H000000FF&
      Height          =   90
      Left            =   1320
      TabIndex        =   2
      Top             =   8160
      Width           =   435
   End
   Begin VB.Label lblINI 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Contando Discos: 00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1350
      TabIndex        =   1
      Top             =   7440
      Width           =   8970
   End
   Begin VB.Label lblProces 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buscando discos"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   7710
      UseMnemonic     =   0   'False
      Width           =   9030
   End
   Begin VB.Image Image1 
      Height          =   6825
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   540
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
    
    VVV = "3PM v " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
    '--------
    'cargar los previstos
    
    
    tERR.Anotar "acmz", K.LICENCIA
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(WINfolder + "SL\imgbig.tbr") Then
            Image1.Picture = LoadPicture(WINfolder + "SL\imgbig.tbr")
            frmVIDEO.picBigImg = LoadPicture(WINfolder + "SL\imgbig.tbr")
        Else
            Image1.Picture = LoadPicture(SYSfolder + "f52.dlw")
            frmVIDEO.picBigImg = LoadPicture(SYSfolder + "f52.dlw")
        End If
    Else
        Image1.Picture = LoadPicture(SYSfolder + "f52.dlw")
        frmVIDEO.picBigImg = LoadPicture(SYSfolder + "f52.dlw")
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
    'leer el archivo de configuracion SYSfolder + "3pmcfg.tbr"
    CargarIMGinicio = LeerConfig("CargarImagenInicio", "1")
    AutoReDibuj = LeerConfig("AutoReDraw", "1")
    TeclaDER = Val(LeerConfig("TeclaDerecha", "88"))
    TeclaIZQ = Val(LeerConfig("TeclaIzquierda", "90"))
    TeclaPagAd = Val(LeerConfig("TeclaPagAd", "77"))
    TeclaPagAt = Val(LeerConfig("TeclaPagAt", "78"))
    TeclaOK = Val(LeerConfig("TeclaOK", "13"))
    TeclaESC = Val(LeerConfig("TeclaESC", "27"))
    TeclaNewFicha = Val(LeerConfig("TeclaNuevaFicha", "81"))
    TeclaConfig = Val(LeerConfig("TeclaConfig", "67"))
    TeclaCerrarSistema = Val(LeerConfig("TeclaCerrarSistema", "87"))
    tERR.Anotar "acnb"
    TeclaShowContador = Val(LeerConfig("TeclaShowContador", "85")) 'U
    TeclaPutCeroContador = Val(LeerConfig("TeclaPutCeroContador", "86")) 'V
    TeclaFF = Val(LeerConfig("TeclaFF", "74")) 'J
    TeclaBajaVolumen = Val(LeerConfig("TeclaBajaVolumen", "68")) 'D
    TeclaSubeVolumen = Val(LeerConfig("TeclaSubeVolumen", "69")) 'E
    TeclaNextMusic = Val(LeerConfig("TeclaNextMusic", "66")) 'B
    
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
    CreditosCuestaTema(0) = LeerConfig("CreditosCuestaTema", "1")
    CreditosCuestaTema(1) = LeerConfig("CreditosCuestaTema2", "2")
    CreditosCuestaTema(2) = LeerConfig("CreditosCuestaTema3", "3")
    'inicializar los precios
    PrecNowAudio = CreditosCuestaTema(0)
    CreditosCuestaTemaVIDEO(0) = LeerConfig("CreditosCuestaTemaVIDEO", "2")
    CreditosCuestaTemaVIDEO(1) = LeerConfig("CreditosCuestaTemaVIDEO2", "3")
    CreditosCuestaTemaVIDEO(2) = LeerConfig("CreditosCuestaTemaVIDEO3", "4")
    PrecNowVideo = CreditosCuestaTemaVIDEO(0)
    textoUsuario = LeerConfig("TextoUsuario", "Cargue los datos de su empresa aqui")
    tERR.Anotar "acnf"
    'publicidad
    'inicializar publicidades si corresponde
    MostrarPUB = LeerConfig("MostrarPub", "0")
    PubliCada = LeerConfig("PubliCada", "5")
    IDIOMA = LeerConfig("Idioma", "Espa�ol")
    PUBs.HabilitarPublicidadesMp3Vid = MostrarPUB
    PUBs.SonarPublicidadesCada = PubliCada
    
    MostrarPUBIMG = LeerConfig("MostrarPubIMG", "0")
    PubliIMGCada = LeerConfig("PubliIMGCada", "10")
    PUBs.HabilitarPublicidadesIMG = MostrarPUBIMG
    PUBs.SonarPublicidadesIMGCada = PubliIMGCada
    tERR.Anotar "acng"
    'la cargo si o si para que si despues entra a la conficuracion ya este cargada
    PUBs.CargarPUBs
    
    'cargar variables de claves
    'archivo de claves
    If FSO.FileExists(WINfolder + "sevalc.dll") = False Then
        Set TE = FSO.CreateTextFile(WINfolder + "sevalc.dll", True)
        TE.WriteLine "Config:12345612345612345612"
        TE.WriteLine "Close:45612345612345612345"
        TE.WriteLine "Credit:1234441234441234561"
        TE.Close
    End If
    tERR.Anotar "acnh"
    Set TE = FSO.OpenTextFile(WINfolder + "sevalc.dll", ForReading, False)
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
    If FSO.FileExists(SYSfolder + "f9yaSL.nam") Then
        'aqui hay un error de acceso denegado si es de solo lectura!!!!!
        'se corrige as�.
        tERR.Anotar "acni4"
        FSO.CopyFile SYSfolder + "f9yaSL.nam", AP + "discos\01- Los mas escuchados\tapa.jpg", True
    Else
        If FSO.FileExists(SYSfolder + "f54.dlw") Then
            tERR.Anotar "acni5"
            'aqui hay un error de acceso denegado si es de solo lectura!!!!!
            'se corrige as�.
            FSO.CopyFile SYSfolder + "f54.dlw", AP + "discos\01- Los mas escuchados\tapa.jpg", True
        Else
            tERR.Anotar "acni6"
            MsgBox "No se encuentra el archivo de imagen del Ranking!. La instalacion de 3PM no es corecta!"
            End
        End If
    End If
    
    tERR.Anotar "acnj"
    If FSO.FileExists(SYSfolder + "f61.dlw") = False Then
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
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        tERR.Anotar "000A-00902"
        FSO.CreateTextFile AP + "ranking.tbr", True
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
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
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
            PBar.Width = z * 10
            If PBar.Width > lblProces.Width Then PBar.Width = 100
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
    
    Dim MaxPT As Long 'comparacoin de cadenas. Empiezo con el m�ximo
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
    Set TeRank = FSO.CreateTextFile(AP + "ranking.tbr", True)
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
            'WriteTBRLog "limpiado del Rank: " + _
            '    txtInLista(mtxTOP10(Ordenados(mtx)), 2, ","), False
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
