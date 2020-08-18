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
      Top             =   8070
      Width           =   435
   End
   Begin VB.Label lblINI 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Contando Discos: 00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1350
      TabIndex        =   1
      Top             =   7470
      Width           =   4620
   End
   Begin VB.Label lblProces 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buscando discos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Width           =   6660
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
    MostrarCursor False
    On Local Error GoTo NoINI
    'VVV = "v " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
    '--------
    'cargar los previstos
    Image1.Picture = LoadPicture(SYSfolder + "f52.dlw")
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(WINfolder + "SL\imgbig.tbr") Then Image1.Picture = LoadPicture(WINfolder + "SL\imgbig.tbr")
    End If
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
    VolumenIni = CLng(LeerConfig("Volumen", "50"))
    VolumenIni2 = CLng(LeerConfig("Volumen2", "50"))
    EsperaTecla = Val(LeerConfig("EsperaTecla", "900"))
    PorcentajeTEMA = Val(LeerConfig("PorcentajeTema", "60"))
    FASTini = LeerConfig("FastIni", "1")
    HabilitarVUMetro = LeerConfig("HabilitarVUMetro", "1")
    vidFullScreen = LeerConfig("VidFullScreen", "1")
    Salida2 = LeerConfig("Salida2", "0")
    NoVumVID = LeerConfig("NoVumVid", "0")
    OutTemasWhenSel = LeerConfig("OutTemasWhenSel", "0")
    BloquearMusicaElegida = LeerConfig("BloquearMusicaElegida", "1")
    TapasMostradasH = Val(LeerConfig("DiscosH", "3"))
    TapasMostradasV = Val(LeerConfig("DiscosV", "2"))
    PasarHoja = LeerConfig("Pasarhoja", "1")
    DistorcionarTapas = LeerConfig("DistorcionarTapas", "0")
    Protector = LeerConfig("Protector", "1")
    CargarDuracionTemas = LeerConfig("CargarDuracionTemas", "0")
    MostrarRotulos = LeerConfig("MostrarRotulos", "1")
    RotulosArriba = LeerConfig("RotulosArriba", "0")
    DuracionProtect = LeerConfig("DuracionProtect", "180")
    RankToPeople = LeerConfig("RankToPeople", "1")
    TemasPorCredito = LeerConfig("TemasPorCredito", "1")
    CreditosCuestaTema = LeerConfig("CreditosCuestaTema", "1")
    CreditosCuestaTemaVIDEO = LeerConfig("CreditosCuestaTemaVIDEO", "2")
    textoUsuario = LeerConfig("TextoUsuario", "Cargue los datos de su empresa aqui")
    
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
    Set TE = FSO.OpenTextFile(WINfolder + "sevalc.dll", ForReading, False)
    'config/close/credit es el orden del archivo
    ClaveConfig = txtInLista(TE.ReadLine, 1, ":")
    ClaveClose = txtInLista(TE.ReadLine, 1, ":")
    ClaveCredit = txtInLista(TE.ReadLine, 1, ":")
    TE.Close
    
    Me.Show
    Me.Refresh
    
    
    'ver si ya estaba cargado
    If App.PrevInstance Then MsgBox "No se pueden abrir dos instancias de 3pm": End
        
    'ASEGURARSE QUE EXISTA la carpeta del ranking y la imagen que le corresponde
    If FSO.FolderExists(AP + "discos") = False Then
        FSO.CreateFolder AP + "discos"
    End If
    If FSO.FolderExists(AP + "discos\01- Los mas escuchados") = False Then
        FSO.CreateFolder AP + "discos\01- Los mas escuchados"
     End If
    'siempre copiarlo, si es el SL con prioridad
    If FSO.FileExists(SYSfolder + "f9yaSL.nam") Then
        'aqui hay un error de acceso denegado si es de solo lectura!!!!!
        'se corrige as�.
        FSO.CopyFile SYSfolder + "f9yaSL.nam", AP + "discos\01- Los mas escuchados\tapa.jpg", True
    Else
        If FSO.FileExists(SYSfolder + "f54.dlw") Then
            'aqui hay un error de acceso denegado si es de solo lectura!!!!!
            'se corrige as�.
            FSO.CopyFile SYSfolder + "f54.dlw", AP + "discos\01- Los mas escuchados\tapa.jpg", True
        Else
            MsgBox "No se encuentra el archivo de imagen del Ranking!. La instalacion de 3PM no es corecta!"
            End
        End If
    End If
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
    On Error GoTo notop
    CaminoError "000A-00901"
    'ver si existe ranking.tbr
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        CaminoError "000A-00902"
        FSO.CreateTextFile AP + "ranking.tbr", True
        CaminoError "000A-00903"
        'si me quedo da error
        GoTo FinOrden
    End If
    CaminoError "000A-00903"
    lblINI.Caption = "Inicializando 3PM..."
    CaminoError "000A-00904"
    lblINI.Refresh
    CaminoError "000A-00905"
    PBar.Width = 0
    CaminoError "000A-00906"
    PBar.Refresh
    CaminoError "000A-00907"
    Dim TT As String
    Dim mtxTOP10() As String, z As Integer
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    Dim Encontrado As Boolean
    Encontrado = False
    'abrir el archivo y CARGARLO A UNA MATRIZ
    CaminoError "000A-00908"
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    'leerlo cargarlo en matriz y ordenar por mas escuchado
    CaminoError "000A-00909"
    'sin esto los archivos vacios se clavan
    ReDim Preserve mtxTOP10(0)
    CaminoError "000A-00910"
    Do While Not TE.AtEndOfStream
        'cada linea es "puntos,arch,nombretema,nombredisco"
        CaminoError "000A-00911"
        TT = TE.ReadLine
        CaminoError "000A-00912"
        If TT <> "" Then
            CaminoError "000A-00913"
            z = z + 1
            CaminoError "000A-00914"
            PBar.Width = z * 10
            If PBar.Width > lblProces.Width Then PBar.Width = 100
            CaminoError "000A-00915"
            ThisPTS = Val(txtInLista(TT, 0, ","))
            CaminoError "000A-00916"
            ThisArch = txtInLista(TT, 1, ",")
            CaminoError "000A-00917"
            ThisTEMA = txtInLista(TT, 2, ",")
            CaminoError "000A-00918"
            ThisDISCO = txtInLista(TT, 3, ",")
            CaminoError "000A-00919"
            ReDim Preserve mtxTOP10(z)
            CaminoError "000A-00920"
            mtxTOP10(z) = TT
        End If
    Loop
    CaminoError "000A-00921"
    TE.Close
    'ordenar la matriz
    'tomar la matriz (con valores separador) y ordenala en base a la
    'columna indicada. en este caso el separador es "," y la columna es 0.
    'seria los mismo que tomara 1 ya que todos tienen el mismo path
    CaminoError "000A-00922"
    Dim MaxPT As Long 'comparacoin de cadenas. Empiezo con el m�ximo
    Dim ubicMAX As Long 'indice en la matriz del menor encontrado cada vuelta
    MaxPT = 0
    Dim c As Long, mtx As Long, ValComp As Long
    c = 0 'cantidad de minimos encontrados
    Dim Ordenados() As Long 'matriz con los indices ordenados
    CaminoError "000A-00923"
    PBar.Width = 0
    PBar.Refresh
    CaminoError "000A-00924"
    Do
        PBar.Width = c * 10
        CaminoError "000A-00925"
        For mtx = 1 To UBound(mtxTOP10)
            'se compara por los puntos
            CaminoError "000A-00926"
            ValComp = Val(txtInLista(mtxTOP10(mtx), 0, ","))
            CaminoError "000A-00927"
            If ValComp > MaxPT Then
                'nunca uno sumara mas de dos puntos (legalmente)
                CaminoError "000A-00928"
                MaxPT = ValComp
                ubicMAX = mtx
            End If
        Next
        CaminoError "000A-00929"
        'al mayor lo quito para que no salga de nuevo
        mtxTOP10(ubicMAX) = "0," + mtxTOP10(ubicMAX)
        c = c + 1
        CaminoError "000A-00930"
        ReDim Preserve Ordenados(c)
        Ordenados(c) = ubicMAX
        CaminoError "000A-00931"
        If c >= UBound(mtxTOP10) Then Exit Do
        MaxPT = 0
    Loop
    'cargar todos y sacar la primera columna de las zetas
    CaminoError "000A-00932"
    PBar.Width = 0
    PBar.Refresh
    CaminoError "000A-00933"
    Dim MTXsort() As String
    'cambie opentextfile por createtextfile por un error que suele dar
    Dim TeRank As TextStream
    Set TeRank = FSO.CreateTextFile(AP + "ranking.tbr", True)
    'si no hay nada para escribir el Close da error?!?!?!?!?
    Dim RankWrite As Long
    RankWrite = 0
    CaminoError "000A-00934"
    For mtx = 1 To UBound(mtxTOP10)
        CaminoError "000A-00935"
        ReDim Preserve MTXsort(mtx)
        'como se agrego un indice mas en archivo esta en el indice2
        'ver si existe si si no no cargarlo
        CaminoError "000A-00936"
        If Dir(txtInLista(mtxTOP10(Ordenados(mtx)), 2, ",")) <> "" Then
            MTXsort(mtx) = txtInLista(mtxTOP10(Ordenados(mtx)), 1, ",") + "," + _
                txtInLista(mtxTOP10(Ordenados(mtx)), 2, ",") + "," + _
                txtInLista(mtxTOP10(Ordenados(mtx)), 3, ",") + "," + _
                txtInLista(mtxTOP10(Ordenados(mtx)), 4, ",")
            CaminoError "000A-00938"
            TeRank.WriteLine MTXsort(mtx)
            PBar.Width = mtx * 10
            RankWrite = RankWrite + 1
        Else
            CaminoError "000A-00937"
            WriteTBRLog "limpiado del Rank: " + _
                txtInLista(mtxTOP10(Ordenados(mtx)), 2, ","), False
            Limpiaron = Limpiaron + 1
        End If
    Next
    CaminoError "000A-00939"
    'si no hay nada para escribir el Close da error?!?!?!?!?
    If RankWrite = 0 Then TeRank.WriteLine ""
    CaminoError "000A-00981"
    TeRank.Close
    CaminoError "000A-00940"
    Set TeRank = Nothing
    If Limpiaron > 0 Then WriteTBRLog "Se limpiaron " + CStr(Limpiaron) + " temas", True
    '==================================================================
FinOrden:
    'se inicializa el contador para que la variable CONTADOR tenga el
    'valor de todas las fichas cargadas
    'si este es cero esta en los primeros usos entonces mostrar el CLUF
    CaminoError "000A-00941"
    SumarContadorCreditos 0
    frmIndex.Show 1
    Exit Sub
notop:
    WriteTBRLog "frmINI - Ranking ordenar. " + vbCrLf + Err.Description, True
    Resume Next
    Exit Sub
NoINI:
    WriteTBRLog "frmINI - LOAD. " + vbCrLf + Err.Description, True
    Resume Next
End Sub
