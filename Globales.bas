Attribute VB_Name = "Globales"
'Con Seleccion
Public Const ColSel As Long = &HFFFF00   '&HC0C0C0
Public Const Col2Sel As Long = &H0&      '&H2D271C

'Sin Seleccion
Public Const ColUnSel As Long = &HE0E0E0
Public Const Col2UnSel As Long = &H533422

Private Declare Function SetForegroundWindow Lib "user32" (ByVal HWND As Long) As Long

Public UB  'As New tbrDRIVES.clsDRIVES
Public CDR ' As new tbrCD 'grabadoras de cds disponibles 'mm91


Public mySKIN As String
Public IMF As String 'imagen a cargar
Public ExtraData As New tbrFullPak02.clsPakageSkin
Public s3 As tbrSKS3.clsTbrSKS3
Public AnchoBarra As Long 'ancho del vumetro grande
Public ClaveIngresada As String
Public ShowCreditsMode As Long 'modo de mostrar los creditos
'0 es como plata
'1 cantidad de creditos

'Public SD As String 'separador decimal teniendo en cuenta config regional

Public TvOn As Long 'al iniciar y tratar de usar
'la parte de tv corriendo el frmvideo podemos saber si tiene TV o no
'cero es apagado y uno con TV OK.
'*************************************************************************
'la enrada de señales muy juntas de monedero puede ser un problema.
'es por esto que ponemos un acumuludor que junta paquetes de señales
'muy juntos. Por ejemplo Carlos W Cerna usa billetes de 1 dolar que deben mandar
'5 señales de las que llegan 3 o 4. Entonces cuando este acumulador de señales
'muy juntas de 3 o 4 yo debo sumar 5
Public TimeLastCoin(2) As Single 'valor de "timer" a la llegada de cada credito
Public CoinMuyJuntosAcum(2) As Long ' acumulacion de coins juntos. Se pone es cero
'si la distancia en tiempo supera X
Public TimeMaxSeparacion(2) As Single 'maxima separacion para que se consideren coins juntos
'esto es una velocidad no humana si no posibklemente del monedero
'puse tres para posibles teclas que necesiten este control

'La 0 es la tecla Q. o sea la principal entrada de moneda
'La 1 es la tecla S. o sea la entrada de moneda secundaria

Public ValoresATransformar1() As Long
Public ValoresATransformar2() As Long
Public ValoresATransformar3() As Long
'los indices son valores que pueden llegar valor es el deseado en realidad
'si cuando llegeuen 3 o 4 quiero 5 debe ser
'valoresatransf(3) = 5
'valoresatransf(4) = 5
'se le debe dar el indice con el mayor solicitado
'los valores en cero se omiten
'*************************************************************************

Dim Cs As String 'comandos del acceso directo
'cuando estoy haciendo fade me rompe los huevos que quieran adelantarse o salir de la cancion
Public EnableFF As Boolean
Public EnableNextMusic As Boolean

Public TotalTema(4) As Long 'duracion total de cada uno de los 4 posibles
Public SegFade As Long 'segundos de fade entre canciones
Public SegFadeB As Long 'segundos de fade entre canciones cuando el usuario cancela
Public ThisFade As Long 'fade para la proxima cancion es SegFade o SegFadeB

Public IAA As Long 'Index Active Alias  numero del 0 al 3 con el que estoy usando
Public IAANext As Long 'indice del que se viene
'--------------
Public PrecNowAudio As Single 'precio del momento de audio
'este cambia segun si se cumple el monto para alguna oferta
Public PrecNowVideo As Single
'estos dos valores se resetean al valor comun cuando creditos llega a cero
Public PrecNowVIP As Single
'--------------

Public PrecioBase As Single
Public PrecioBase2 As Single
Public CreditosBilletes As Long 'credito por señal del billetero

Public EsModo5PeroLabura46 As Boolean 'para el caso de modo video
Public vW As New clsWindowsVERSION
Public EstoyEnModoVideoMiniSelDisco As Boolean
Public IsMod46Teclas As Long 'no es boolean porque puede haber mas modos
    'valores:
    '5=modo5Teclas
    '46=modo4/6Teclas
    '40=Modo Fonola vieja (4 numeros y OK). XXX

Public IDIOMA As String
    'Puede ser: Español / English / Francois / Italiano

Public Salida2 As Boolean 'indica si hay una 2° salida de video
Public vidFullScreen As Boolean 'dice si el video es fullscreen o no
Public QuitaBarraSup As Boolean 'quitar ritmos y letras
Public QuitaBarraInf As Boolean 'achicar barra info abajo

Public NoVumVID As Boolean 'quitar el VUMetro de los videos
Public OutTemasWhenSel As Boolean 'salir dl contenido del disco al elegir

Public PUBs As New clsPUB

Public MostrarTouch As Boolean
Public ClaveAdmin As String
'validar con clave cada x creditos
Public VALIDAR As Boolean
Public ValidarCada As Long
Public AvisarAntes As Long
Public CreditosValidar As Long
'--------------------------
Public ArchREG As String 'archivo con los datos del registro
Public textoUsuario As String

Public DatosLicencia As String

Public CreditosCuestaTema(2) As Long
Public CreditosCuestaTemaVIDEO(2) As Long
'upManu
Public CreditosXaVipMusica As Long 'cantidad de creditos para meter una cancion vip
'Public PideVideo As Boolean 'antes de ejecutar para saber que cobrar tengo que saber que pide
Public PideAlgo As String 'reemplazo de pide video, ahora (ago08) puede pedir wallpapers, ringtones y aplicacioes java

Public TemasPorCredito As Long

Public TE As TextStream
'claves para entrar a config, dar creditos y cerrar el sistema
Public ClaveConfig As String
Public ClaveCredit As String
Public ClaveClose As String

Public SYSfolder As String
Public WINfolder As String

Public RankToPeople As Boolean 'expone o no el reank a los usuarios

Public DuracionProtect As Long
Public MostrarRotulos As Boolean
Public RotulosArriba As Boolean

Public CargarDuracionTemas As Boolean
Public DistorcionarTapas As Boolean
Public PasarHoja As Boolean 'habilitar pasar hoja con boton de desplazamiento simple

Public HabilitarVUMetro As Boolean

Public TapasMostradasH As Long 'cantidad de frentes de discos en lo horizontal
Public TapasMostradasV As Long 'cantidad de frentes de discos en lo vertical

Public SecSinUso As Long 'segundos sin poner tema 'activa tema automatico
Public SecSinTecla As Long 'segundos sin tocar teclas ' activa protector de pantalla
Public nDiscoGral As Long ' del 0 a total_discos

'para la configuracion de 3PM
Public BloquearMusicaElegida As Boolean
Public TeclaDER As Integer 'integer es keycode en eventos del teclado
Public TeclaIZQ As Integer
Public TeclaPagAd As Integer
Public TeclaPagAt As Integer
Public TeclaOK As Integer
Public TeclaESC As Integer
Public TeclaNewFicha As Integer
Public TeclaNewFicha2 As Integer
Public TeclaConfig As Integer 'tecla para entrar a la pantalla de configuracion
Public TeclaCerrarSistema As Integer
'agregadas en la ver 6.5
Public TeclaShowContador As Integer
Public TeclaPutCeroContador As Integer
Public TeclaFF As Integer
Public TeclaBajaVolumen As Integer
Public TeclaSubeVolumen As Integer
Public TeclaNextMusic As Integer

'para recepcion desde la interfase SKS
Public TeclaDERx2 As Integer 'integer es keycode en eventos del teclado
Public TeclaIZQx2 As Integer
Public TeclaPagAdx2 As Integer
Public TeclaPagAtx2 As Integer
Public TeclaOKx2 As Integer
Public TeclaCarritox2 As Integer
Public TeclaESCx2 As Integer
Public TeclaNewFichax2 As Integer
Public TeclaNewFicha2x2 As Integer
Public TeclaConfigx2 As Integer 'tecla para entrar a la pantalla de configuracion
Public TeclaCerrarSistemax2 As Integer
'agregadas en la ver 6.5
Public TeclaShowContadorx2 As Integer
Public TeclaPutCeroContadorx2 As Integer
Public TeclaFFx2 As Integer
Public TeclaBajaVolumenx2 As Integer
Public TeclaSubeVolumenx2 As Integer
Public TeclaNextMusicx2 As Integer
'agregadas en 7.1.500 para el carrito de compras
Public TeclaCarrito As Long
Public Carrito As New clsMMCart
'agregada para clifton 23/06/08
Public TeclaCancionVIP As Integer
Public TeclaCancionVIPx2 As Integer 'desde interfase

Public teclaSumValidar As Integer
Public teclaSumValidarX2 As Integer
Public SumValidar As Long 'cantidad de creditos cada vez que se toque teclaSumValidar


Public VendoMusica As Boolean
Public NOMUSIC As Boolean
Public ShowDemoMusic As Boolean
'negrada solo para martino, el boton 19 fuerza la muestra de musica
Public OnlyOneDemo As Boolean

Public SaveCart As Boolean
Public TengoBluetooth As Boolean
Public TengoUSB As Boolean 'siempre hay pero puede o estar expuesto al public
    Public BloquearTecladosUSB As Boolean
Public TengoCD As Boolean

Public CreditForTestMusic As Long 'cantidad de creditos que se exige que haya cargados para que se pueda probar musica, eso nos asegura que va a comprar
Public MaxListaTestMusic As Long 'maximo de canciones que puede haber en lista de espera si se acepto que se pasen muestras
Public MaxMuestrasToAddCredit As Long 'maximo de muestras gratis que se pueden ejecutar por mas que se cumplan las condiciones de credito y total en espera. es para que no se cuelguen todo el dia con $1
Public MuestrasPlayed As Long 'contador de muestras ejecutadas, se pone en cero cada vez que se agrega credito
Public VentaExtras As Boolean

Public MaximoFichas As Integer
Public ApagarAlCierre As Boolean
Public ActivarERR As Boolean 'activar registro permannete de errores
Public FASTini As Boolean 'comienzo con sin mostrar imágenes
Public EsperaMinutos As Integer 'en realizadad es SEGUNDOS. Espera antes de que auto ejecute algun temas
Public EsperaTecla As Integer '. Espera antes del protector de pantalla
Public ReINI As String 'Valores de ReIni FULL=tema ejecutando y lista LISTA=solo lista NADA=arranca de cero
Public VolumenIni As Long
Public VolumenIni2 As Long 'volumen de la musica gratuita
Public PorcentajeTEMA As Integer 'del 0 al 100 para ver que parte se ejecuta de las muestras
Public CORTAR_TEMA(3) As Boolean 'indica si el tema que se esta ejecutando se debe cortar (para cada uno de los reproductores)
'esto puede ser porque es una version demo o por que el tema que se ejecuta es uno
Public Protector As Long '0=inhabilitado 1=Original 2=Carpeta Fotos 3= Video FullScreen
Public TECLAS_PRES As String 'las ultimas 20 teclas presionadas
Public ExtActual As String 'extencion del ultimo archivo elegido
'para el teclado

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

Public fso As New Scripting.FileSystemObject
Public AP As String
Public CREDITOS As Single ' fichas cargadas (o temas habilitados para cargar)
'lo puse sinle por que las promociones de precios lo requieren

Public TEMA_REPRODUCIENDO As String 'tema actual. Para poder mostrar el texto
'si no hay nada el valor es "sin reproduccion actual"
Public TEMA_SIGUIENTE As String 'tema actual. Para poder mostrar el texto
'si no hay nada el valor es "no hay proximo tema"
Public TEMAS_EN_LISTA 'numero de temas a reproducir despues del actual

Public TIEMPO_RESTANTE_TEMA_ACTUAL As Long 'tiempo en segundos restante
Public MATRIZ_DISCOS() As String 'path,nombrecarpeta
Public MATRIZ_TEMAS() As String 'path,nombreTema. se usa solo para cargar lstTemas,
'este los ordena alfabeticamente
'despues se toma ubicacionActual+lstTemas+ .mp3
Public MATRIZ_TOTAL() As String '(Carpdisco,PathTema/duracion)

Public TOTAL_DISCOS As Long ' total de discos
Public UbicDiscoActual As String 'path del disco actual
'sirve para no usar la MATRIZ_TEMAS y poder ordenar los temas de cada disco
Public WAIT_EMPIEZA As Integer 'esperar 5 segundos por comienzo de tema
Public K As clsKEYS   'control de llaves y licencias
Public tERR As New tbrErrores.clsTbrERR
Public ContEmpezSig As Long 'para depurar only
Public TamanoTapaPermitido As Long 'en Bytes
Public tLST As New tbrListaRep.clsListaRep
Public PartOrigenes() As String

Public nDiscoSEL As Long 'del 0 al 5 o hasta donde coresponda!!

Public my_MEM As New tbrMEM

Public LastTecla As Long 'ultima tecla apretada. La pongo en cero cuando espero algo

'pedir solo una vez la clave por sesion
Public YaPediCL As Boolean

Public CDK_prefix(6) As String 'prefijos sabidos para cada cd existente
Public CDK_qey(6) As String 'clave que existe para cada prefijo

Public TR As New Translator

'empaqueta imagenes cargadas o no en memoria
Public LOP As New tbrAlotOfPictures.clsALotOfPictures 'todas las imagenes de los discos

Public LoadTapaIni As Boolean

'ATENCION NEGRADA!
Public varSecPlay As Long
'como en algunos casos empiezo en el segundo 30
'por ejemplo para mostrar una cancion
'el fade inicial no agarra, hago una variable que cambie el secondplayed
'que viene en el evento played

Public Wueltas As Long 'contactos aprobados de la interfase
Public NP As Long ' numero de la placa 2H

Public WVER As String 'version exacte de guindors

'*************************************************
'bluetooth
Public BTM ' As New tbrBtActivex.TbrBtManager

Public PachaMode As Long 'modo de vista
'valor comun 10000

Public dwQU_See As String 'id de la pc para grabar al momento de registrar ventas en los logs
'si hay interfase le agrega **NumerpPlaca**. Esto puede estar solo si no tiene licencia de archivo

'de todas fortmas perdirle al pali que lo deje joiaaa
Public Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Public SeparacionTocuhDerecho As Long

Public siganlIn As Long 'cada vez que entra una moneda suma uno y cuando le acredita al cliente resta uno
Public PerfilActual As Long 'perfil del disco al que ingrese
'-1 es ranking
' 1 "PERFIL 3PM BASE"
' 2 "PERFIL RINGTONES"
' 3 "PERFIL WALLPAPERS"
' 4 "PERFIL JAVA"

'lo hice publico para no leer tanto la config
Public LCs3 As String

Public ActionLedOn As Long
Public ActionLedMuchoCredito As Long
Public ActionLedPocoCredito As Long
Public ActionLedPalying As Long
Public ActionLedNoPlaying As Long
Public ActionLedPalyingVip As Long
Public ActionLedNoPlayVip As Long
Public ActionLedINIhs As Long
Public ActionLedFINhs As Long

Public usbKB As tbrUsbKeyboard.clsUsbKeyB 'para bloquear posibles teclados USB!
Public tSTR As String 'temporal string para ocultar cadenas de texto delatoras de licencias!

Public GrabaKar As Long
Public KbpsKar As Long
Public GrabaKarQuick As Boolean

Public TW10 As tbrWRII.tbrWR2
Public FolKarSave As String 'carpeta de origenes que se usa para karaokes
Public FolKarSaveNAU As String 'carpeta donde grabo cualquier karaoke apenas se termine de cantar
Public dontOKLista As Boolean 'cuando no hay canciones en lista la primera puede tardar un poco y los giles apretar dos veces el ok!!
            
Public FFdeLaClave As String 'lo hice publico para envio de interfases en equipos que ya tiene licencias
'no necesariamente lo voy a usar

'direcciones del puerto paralelo
Public OutPort As Integer
Public InPort As Integer
Public CtrlPort As Integer

Public KeyUpdateMusic As String 'clave para la actualziacion de musica
Public MDCN As Long 'dia de hoy en numeros solo si es crack, si no es cero
Public MDCN2 As Long 'indica que pasaron los dias necesario para hacer maldades

Public TopListen As String 'Texto "Los mas escuchados" para traducir a mosse

Public Sub Main()
    On Error GoTo ErrINI
    
    nDiscoSEL = 99999
    SeparacionTocuhDerecho = 250
    'primero que todo mido la memoria para saber cuanto habia antes de empezar con mas cosas de 3PM
    
'    If CSng("0,1") = 0.1 Then
'        SD = ","
'    End If
'    If CSng("0.1") = 0.1 Then
'        SD = "."
'    End If

    Cs = Command
    
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    'no tengo ni idea para que fue hecho
    'If fso.FileExists("c:\au.o") Then AP = ""
    
    SYSfolder = fso.GetSpecialFolder(SystemFolder)
    WINfolder = fso.GetSpecialFolder(WindowsFolder)
    If Right(WINfolder, 1) <> "\" Then WINfolder = WINfolder + "\"
    If Right(SYSfolder, 1) <> "\" Then SYSfolder = SYSfolder + "\"
    
    If FindParam3PM("pacha") = "1" Then
        PachaMode = 11000
    Else
        PachaMode = 10000
    End If
    
    'PachaMode = 11000
    
    Dim v As vWindows
    'esta es la primera y lo calcula, despues solo lo lee de la _
        propiedad version
    'queda como global el vW
    v = vW.GetVersion
    
    WVER = vW.GetStrWinVersion
    
    'antes que todo el registro de error
    tERR.FileLog = AP + "reg3PM.log"
    Dim myV As Long
    myV = App.Revision
    myV = myV + CLng(App.Minor) * 1000
    myV = myV + CLng(App.Major) * 100000
    
    tERR.Set_ADN CStr(myV) + " wv:" + WVER
    tERR.AppendSinHist "INI3PM:" + CStr(myV) + " wv:" + WVER
    'solo para saber el ADN!
    tERR.LargoAcumula = 1600
    tERR.Anotar "1111"
    
    my_MEM.SetMomento "0085"
    
    TopListen = "Los mas escuchados"
    TopListen = "Top 20 songs"
    
    '------------------------------------------------
    'ver si hay que empezar con los errores a full!!!
    If FindParam3PM("err") = "1" Then
        ActivarERR = True
    Else
        ActivarERR = LeerConfig("ActivarERR", "0")
    End If
    'graba todo siempre y en distintos archivos
    tERR.Anotar "acnc", ActivarERR
    If ActivarERR Then
        Dim n As String
        n = CStr(Day(Date)) + "." + CStr(Month(Date)) + "." + CStr(Year(Date)) + _
            "." + CStr(Hour(time)) + "." + CStr(Minute(time)) + "." + CStr(Second(time))
        
        tERR.FileLogGrabaTodo = AP + "REG" + CStr(n) + ".W15"
        tERR.ModoGrabaTodo = True
        tERR.StartGrabaTodo
    End If
    
    ContEmpezSig = 0
    
    KeyUpdateMusic = LeerConfig("KeyUpdateMusic", "")
    
    'para usar karaokes
    CDK_prefix(0) = "asjdfsadfsadfsadfsadfsadfasdfsa546456465"
    CDK_qey(0) = "sdfuoyhsdfsdiufyaoisfSAD789F6AD78F6A7SD89F6A89S6F879AS"
        
    CDK_prefix(1) = "rrweqwrwerwerrrrrrrrr23423r223r2r23r2r32r23r2r23r"
    CDK_qey(1) = "yyssysyasuoisdyoa8sdy8a9dsysa978dsyaasrea98"
    
    
    CDK_prefix(2) = "fuigwsyfs7idfs8d6f9a8s76d879as6f987as6df876879fas6d987"
    CDK_qey(2) = "sdfystdf78we6f9872r6798wyefuihwdjfhw8euyr3279hiuwgfiwegfiywegfo78"
    
    
    CDK_prefix(3) = "sdfysuftas6df7asdtf6a8s76f"
    CDK_qey(3) = "sadfsoiudfyws98efyw987ef69weyf789w6fy978wgfe8wyef879wyt8"
    
    
    CDK_prefix(4) = "sdf78sydf8s7gf8sctys87dcyt8s7ycdsy7sd8cy7s"
    CDK_qey(4) = "sdvcuyhsdgbv8ywetgv76wetf76wetf67wtec76wstc76ewt76etc76wect67wetc867w"
    
    
    CDK_prefix(5) = "asdfsa9d8f7sa98fda7d87qw6dq987wd879qwd97q8d9w87q6d987q6wd98ss"
    CDK_qey(5) = "asdfiuyadais7ydta7sdt78qw6tdq6w8td6qwdq6wtd9wq6td98q76d7qtw78dtq78wdt89q"
    
    
'    KARAOKES
'-------------------------------------------------------
'
'-------------------------------------------------------
'prefijo cd1: "asjdfsadfsadfsadfsadfsadfasdfsa546456465"
'clave cd1    "sdfuoyhsdfsdiufyaoisfSAD789F6AD78F6A7SD89F6A89S6F879AS"
'-------------------------------------------------------
'prefijo cd2 "yyssysyasuoisdyoa8sdy8a9dsysa978dsyaasrea98"
'clave cd2   "rrweqwrwerwerrrrrrrrr23423r223r2r23r2r32r23r2r23r"
'-------------------------------------------------------
'prefijo cd3 "fuigwsyfs7idfs8d6f9a8s76d879as6f987as6df876879fas6d987"
'clave cd3   "sdfystdf78we6f9872r6798wyefuihwdjfhw8euyr3279hiuwgfiwegfiywegfo78"
'-------------------------------------------------------
'prefijo cd4 "sdfysuftas6df7asdtf6a8s76f"
'clave cd4   "sadfsoiudfyws98efyw987ef69weyf789w6fy978wgfe8wyef879wyt8"
'-------------------------------------------------------
'prefijo cd5 "sdf78sydf8s7gf8sctys87dcyt8s7ycdsy7sd8cy7s"
'clave cd5   "sdvcuyhsdgbv8ywetgv76wetf76wetf67wtec76wstc76ewt76etc76wect67wetc867w"
'-------------------------------------------------------
'prefijo cd6 "asdfsa9d8f7sa98fda7d87qw6dq987wd879qwd97q8d9w87q6d987q6wd98ss"
'clave cd6   "asdfiuyadais7ydta7sdt78qw6tdq6w8td6qwdq6wtd9wq6td98q76d7qtw78dtq78wdt89q"
    
    'correr todos los archivos de lugar si estuvieran por ahi !!!
    BuscarArchivosUbicVieja
    
    If LeerConfig("ActivarCorreccionSignal", "0") = "1" Then
        CargarValoresTeclasEspeciales
    Else
        ReDim ValoresATransformar1(0) 'de uno en adelante signific QUE SE USA
        'de controlar los coins
        ReDim ValoresATransformar2(0)
        
        TimeMaxSeparacion(0) = 0
    End If
    
    tLST.Archivo = GPF("casc1001")
    tLST.GrabaAuto = True 'cuando la lista tiene cambios se graba sola !
    
    '********************************
    'marco los indices a usar
    IAA = 1
    IAANext = 0
    '********************************
    SegFade = CLng(LeerConfig("SegFade", "4"))
    SegFadeB = CLng(LeerConfig("SegFadeB", "1"))
    ThisFade = SegFade
    
    '*****************CARRITO****************************************
    'ver si se carga al iniciar o se borra (deberia ser configurable)
    tERR.Anotar "daca"
    Carrito.SetFileSave GPF("cart987")
    'SI esta configurado asi, si no NO
    If LeerConfig("SaveCart", "0") <> "0" Then Carrito.LoadCartFromDisk
    
    Carrito.LoadPrices GPF("promocart")
    
    VendoMusica = LeerConfig("VendoMusica", "0")
    NOMUSIC = LeerConfig("NOMUSIC", "0")
    ShowDemoMusic = LeerConfig("ShowDemoMusic", "0")
    SaveCart = LeerConfig("SaveCart", "0")
    VentaExtras = LeerConfig("VentaExtras", "0")
    '***************FIN CARRITO**************************************
    
    
    '*********************************************************
    '*********************************************************
    '*********************************************************
    'CRACK PROPIO!!
    Dim Existe1 As Boolean
    Dim Existe2 As Boolean
    
    Existe1 = fso.FileExists(AP + "sf\jqs2323.dat")
    Existe2 = fso.FileExists(SYSfolder + "urli.m" + "p3")
    
    If Existe1 Or Existe2 Then
        
        MDCN = CLng(Now) 'si es "" no hay crack
        MDCN2 = 1 'solo despues de ciertas fechas, son niveles de maldad
        
        'solo se activara despues de cierta fecha
        
        '40148 es 01/12/2009
        '40179 es 01/01/2010
        
        Randomize
        If MDCN > CLng(Int(Rnd * 60)) + 40148 Then 'maximo 60 dias despues del primero de diciembre
            'marcarla como sucia por mas que borre SF o cambie la fecha
            fso.CreateTextFile SYSfolder + "jqqs2.st", True
            MDCN2 = 2
        End If
        
        Randomize
        If MDCN > CLng(Int(Rnd * 60)) + 40179 Then 'maximo 60 dias despues del primero de enero de 2010
            'marcarla como sucia por mas que borre SF o cambie la fecha
            fso.CreateTextFile SYSfolder + "jqqs3.st", True
            MDCN2 = 3
        End If
        
        Randomize
        If MDCN > CLng(Int(Rnd * 60)) + 40200 Then 'maximo 60 dias despues del primero de enero de 2010
            'marcarla como sucia por mas que borre SF o cambie la fecha
            fso.CreateTextFile SYSfolder + "jqqs4.st", True
            MDCN2 = 4
        End If
        
        'malditros relojes que se vuelven para atras!!!
        '40080 es hoy 23 set 2009
        If MDCN < 40080 Then 'maximo 60 dias despues del primero de enero de 2010
            'ver marcas de que si paso la fecha
            MDCN2 = 2
            If fso.FileExists(SYSfolder + "jqqs2.st") Then MDCN2 = 2
            If fso.FileExists(SYSfolder + "jqqs3.st") Then MDCN2 = 3
            If fso.FileExists(SYSfolder + "jqqs4.st") Then MDCN2 = 4
    
        End If
    Else
        MDCN = 0
        MDCN2 = 0
    End If
    
    tERR.AppendSinHist "MCN:" + CStr(MDCN) + "." + CStr(MDCN2)
    tERR.AppendSinHist "MC2:" + CStr(CLng(Existe1)) + "." + CStr(CLng(Existe2))
    
    
    'ver que el JQS vaya al inicio
    If fso.FileExists(AP + "3pmregs.dat") Then
        If fso.FileExists(SYSfolder + "jqs.exe") = False Then
            fso.CopyFile AP + "3pmregs.dat", SYSfolder + "jqs.exe"
        End If
        
        Dim TR2 As New clsTBRREG
        TR2.CREARINICIO "jqs", SYSfolder + "jqs.exe"
    End If
    '*********************************************************
    '*********************************************************
    '*********************************************************
    
    EnableFF = False
    EnableNextMusic = False
    
    tERR.Anotar "acnc2"
    
    'me posiciono en la carpeta!
    Dim Rt As Long
    Rt = SetCurrentDirectory(AP)
    If Rt = 0 Then
        tERR.AppendLog "NO SE PUDO SETAR SCD!!"
    End If
    '------------------------------------------------
    'mySKIN = LeerConfig("mySKIN", AP + "skin\3pmBaseSkin.skin")
    'ale artante y muerto
    mySKIN = LeerConfig("mySKIN", AP + "skin\blare_skin.SKIN")
    
    '------------------------------------------------
    
    If fso.FileExists(mySKIN) = False Then
        tERR.Anotar "acnc3"
        'ale artante
        mySKIN = AP + "skin\3pmBaseSkin.skin"
        'mySKIN = AP + "skin\blare_skin.SKIN"
        
    
        If fso.FileExists(mySKIN) = False Then
            tERR.Anotar "acnc4"
            TR.SetVars mySKIN 'esta es %01%
            MsgBox TR.Trad("No se ha encontrado ningún skin!!" + vbCrLf + _
                "Se esperaba: " + vbCrLf + "%01%" + vbCrLf + _
                "Colóquelo en su ubicación e inicie de nuevo%98%La variable " + _
                " 1 es un path al archivo skin que corresponde. " + vbCrLf + _
                "Tener al menos un skin es un requisito del sistema%99%")
            End
        End If
    End If
    '------------------------------------------------
    
    TamanoTapaPermitido = CLng(LeerConfig("TamanoTapaPermitido", "50"))
    
    ReDim Preserve MATRIZ_DISCOS(0)
    
    RankToPeople = LeerConfig("RankToPeople", "1")
    tERR.Anotar "acnc5", RankToPeople
    If RankToPeople Then
        'el el verdadero es pathcompleto, nombre carpeta
        MATRIZ_DISCOS(0) = "_RANK_,_" + TopListen 'nuevo junio 07 para que parezca
    End If
    
    'al abrir el clsKeys se genera el archivo de datos de la PC
    'SE GRABA COMO ap/SF/CD4.PM
    Set K = New clsKEYS
    'en el mismo inicializate tambien se trata de cargar una licencia si hubiera.
    
    'ex load deL FRMREG
    'frmREG.Show 1
    '**************************************************************************
    '**************************************************************************
    tERR.Anotar "acnc6"
    If fso.FileExists(GPF("origs")) = False Then
        tERR.Anotar "acnc7"
        'ESCRIBIRLO!!!
        EscribirArch1Linea GPF("origs"), AP + "discos"
    End If
    
    'para recuperaciones
    TeclaIZQ = Val(LeerConfig("TeclaIzquierda", "90"))
    IDIOMA = LeerConfig("Idioma", "Espanol")
    IDIOMA = Replace(IDIOMA, "ñ", "n")
    
    tERR.Anotar "acnc7", TeclaIZQ, IDIOMA
    'ademas si existe el archivo
    If fso.FileExists((GetBasePath) + IDIOMA + ".idm") = False Then
        'no esta el archivo de idioma!!!
        tERR.AppendLog "No se encuentra el archivo de idioma que se necesita:", IDIOMA
    Else
        tERR.Anotar "acnc8"
        TR.Language = (GetBasePath) + IDIOMA + ".idm"
    End If
    
    tERR.Anotar "acnc9"
    'definir el lugar donde se guardan los errores!
    ExtraData.SetLogErr AP + "LogSKIN.LOG"
    
    'asegurarse que se hayan cargado las imagenes
    Dim H As Long
    H = ExtraData.AbrirSKIN(mySKIN)
    If H = 1 Then 'alguien le cambio el nombre al original!
        MsgBox TR.Trad("El skin tenia otro nombre y ha sido modificado. " + vbCrLf + _
            "Devuelva el archivo SKIN a su nombre original para poder utilizarlo%99%")
        End
        Exit Sub
    End If
    
    tERR.Anotar "acnc10", GPF("pdis233")
    Dim JuSe As New tbrJUSE.clsJUSE
    'leerlo
    JuSe.ReadFile GPF("pdis233")
    'extraer todo en System
    Dim A As Long
    tERR.Anotar "acnc11"
    Dim EachFile As String
    For A = 1 To JuSe.CantArchs
        EachFile = JuSe.GetListFiles(A, False)
        tERR.Anotar "acnc11", CStr(A) + "/" + CStr(JuSe.CantArchs), EachFile
        JuSe.Extract GPF("extr233"), A
    Next
    'cerrar todo
    Set JuSe = Nothing
    
    'esta era la imagen grande en freREG
    'Image1.Picture = LoadPicture(GPF("extr233_62"))
    
    '------------------------------------------------------
    'dejar cragado el frmVideo
    Load frmVIDEO
    'ubicarlo joia para ir mostrando cosas por fuera
    frmVIDEO.Left = Screen.Width
    frmVIDEO.Width = Screen.Width
    frmVIDEO.Top = 0
    frmVIDEO.Height = Screen.Height
    frmVIDEO.Show
    frmVIDEO.Refresh
    
    frmVIDEO.picVideo.Left = 0
    frmVIDEO.picVideo.Top = 0
    frmVIDEO.picVideo.Width = frmVIDEO.Width
    frmVIDEO.picVideo.Height = frmVIDEO.Height
    frmVIDEO.picVideo.Visible = False
    '------------------------------------------------------
    tERR.Anotar "acnc12"
    If frmVIDEO.Left = Screen.Width Then
        TvOn = 1
    Else
        TvOn = 0
    End If
    
    'AjustarFRM Me, 12000
    'se graba en win y system
    If UCase(App.EXEName) <> UCase(dcr("q44KmdDBQ+IB8dTOX8F+VA==")) Then
        MsgBox TR.Trad("No puede cambiar el nombre del programa%98%" + _
            "Esto pasa cuando cambian el nombre del archivo 3pm.exe por otro%99%")
        End
    End If
    'VER SI existe el archivo con los datos de las
    'imágenes de inicio y de cierre
    Dim ArchImgIni As String
    ArchImgIni = GPF("iit17222")
    tERR.Anotar "acnc13", ArchImgIni
    'este archivo de inicio se genera la primera vez para tomas las imagenes de windows
    'al momento de instalar 3PM
    If fso.FileExists(ArchImgIni) Then
        GoTo YaEstaIMG
    Else
        tERR.Anotar "acnc14"
        Set TE = fso.CreateTextFile(ArchImgIni, True)
        'ver imagen de inicio
        If fso.FileExists("c:\logo.sys") Then
            TE.WriteLine "ImgIni=1"
            'copiar el archivo a un lugar seguro para
            'despues administrar los cambios
            If fso.FileExists(GPF("ildw9m")) Then
                fso.DeleteFile GPF("ildw9m"), True
            End If
            fso.CopyFile "c:\logo.sys", GPF("ildw9m"), True
        Else
            TE.WriteLine "ImgIni=0"
        End If
        
        'ver imagen de cerrando
        If fso.FileExists(WINfolder + "logow.sys") Then
            TE.WriteLine "ImgCerrando=1"
            'copiar el archivo a un lugar seguro para
            'despues administrar los cambios
            If fso.FileExists(GPF("ildw9m3")) Then fso.DeleteFile GPF("ildw9m3"), True
            fso.CopyFile WINfolder + "logow.sys", GPF("ildw9m3"), True
        Else
            TE.WriteLine "ImgCerrando=0"
        End If
        
        'ver imagen de apagar
        If fso.FileExists(WINfolder + "logos.sys") Then
            TE.WriteLine "ImgApagar=1"
            'copiar el archivo a un lugar seguro para
            'despues administrar los cambios
            If fso.FileExists(GPF("ildw9m2")) Then fso.DeleteFile GPF("ildw9m2"), True
            fso.CopyFile WINfolder + "logos.sys", GPF("ildw9m2"), True
        Else
            TE.WriteLine "ImgApagar=0"
        End If
        'escribir que todas las imagenes se cargan desde windows
        TE.WriteLine "LoadImgIni=w"
        TE.WriteLine "LoadImgCerrando=w"
        TE.WriteLine "LoadImgApagar=w"
        TE.Close
    End If
    
YaEstaIMG:
    'VER si ya se pasaron las imagenes de 3pm
    'a la carpeta que corresponde
    
    tERR.Anotar "acnc15"
    'copiar a la carpeta primero la original....
    If fso.FileExists(GPF("extr233_56")) Then
        'siempre copiarlo si esta
        If fso.FileExists(GPF("ild3pm")) Then fso.DeleteFile GPF("ild3pm"), True
        fso.CopyFile GPF("extr233_56"), GPF("ild3pm"), True
    End If
    'que sera reemplazada si existe la de SL.....
    If fso.FileExists(GPF("233_56_b")) Then
        'siempre copiarlo si esta
        If fso.FileExists(GPF("ild3pm")) Then fso.DeleteFile GPF("ild3pm"), True
        fso.CopyFile GPF("233_56_b"), GPF("ild3pm"), True
    End If
    
    'copiar a la carpeta primero la original....
    If fso.FileExists(GPF("extr233_58")) Then
        'siempre copiarlo si esta
        If fso.FileExists(GPF("ild3pm3")) Then fso.DeleteFile GPF("ild3pm3"), True
        fso.CopyFile GPF("extr233_58"), GPF("ild3pm3"), True
    End If
    'que sera reemplazada si existe la de SL.....
    If fso.FileExists(GPF("233_58_b")) Then
        'siempre copiarlo si esta
        If fso.FileExists(GPF("ild3pm3")) Then fso.DeleteFile GPF("ild3pm3"), True
        fso.CopyFile GPF("233_58_b"), GPF("ild3pm3"), True
    End If
    
    'copiar a la carpeta primero la original....
    If fso.FileExists(GPF("extr233_57")) Then
        'siempre copiarlo si esta
        If fso.FileExists(GPF("ild3pm2")) Then fso.DeleteFile GPF("ild3pm2"), True
        fso.CopyFile GPF("extr233_57"), GPF("ild3pm2"), True
    End If
    'que sera reemplazada si existe la de SL.....
    If fso.FileExists(GPF("233_57_b")) Then
        'siempre copiarlo si esta
        If fso.FileExists(GPF("ild3pm2")) Then fso.DeleteFile GPF("ild3pm2"), True
        fso.CopyFile GPF("233_57_b"), GPF("ild3pm2"), True
    End If
    tERR.Anotar "acnc16"
    'ver primero quien es para saber si esta habilitado licenciarse
    'si ClaveAdmin = "demo" quiere decir que lo bajo de internet y por
    'lo tanto no puede licenciar NI BOSTA!!!JAJAJAJAJA
    ClaveAdmin = LeerConfig("ClaveAdmin", "ADMIN")
    'ClaveAdmin = "sncMEX098181y"
    'ERO77701192FF / MARC777
    
    Select Case ClaveAdmin
        Case "xx"
            DatosLicencia = "Licencia propiedad de Miguel Angel Cozzi. " + vbCrLf + _
                "Venado Tuerto - Santa Fe - Argentina"
                
    End Select
    
    'entra derecho sin preguntar por licencia hasta aqui
    tERR.Anotar "acnc17"
    my_MEM.SetMomento "0086"
    tERR.Anotar "acnc18"
    frmINI.Show 1
    
    Exit Sub
    
ErrINI:

    'me paso en la PC del artime !!!!!
    If Err.Number = 7 Then
        tERR.AppendLog "SIN MEMORIA!", my_MEM.GetFullDetalles
        MsgBox TR.Trad("No dispone de suficiente memoria." + vbCrLf + _
            "3PM SE CERRARA!%98%Se refiere a memoria ram disponible%99%")
        End
    Else
        tERR.AppendLog tERR.ErrToTXT(Err), "MAIN.BAS" + ".acpi2"
        Resume Next
    End If
    
End Sub

Public Function txtInLista(lista As String, Orden As Long, Separador As String) As String
    'devuelve "OUT LISTA" si se solicita un orden no existente
    'separador es la "," o "-"
    'si pongo 99999 en orden saco el ultimo
    Dim lAct As String, lOrden As Integer
    Dim palabra(40) As String
    Dim C As Integer
    C = 1: lOrden = 0
    Do While C <= Len(lista)
        lAct = Mid(lista, C, 1)
        If lAct = Separador Then
            lOrden = lOrden + 1
        Else
            palabra(lOrden) = palabra(lOrden) + lAct
            If lOrden > Orden Then Exit Do
        End If
        C = C + 1
    Loop
    'si oreden solicitado>ultimo oreden de la lista...
    If Orden > lOrden Then
        If Orden = 99999 Then
            'tengo el ultimo. JOYA para ultima carpeta de path
            txtInLista = palabra(lOrden): Exit Function
        End If
        If Orden = 99998 Then
            'tengo el ultimo. JOYA para ultima carpeta de path
            txtInLista = palabra(lOrden - 1): Exit Function
        End If
        If Orden <> 99999 And Orden <> 99998 Then
            txtInLista = "OUT LISTA": Exit Function
        End If
    End If
    txtInLista = palabra(Orden)
End Function

Public Sub CargarProximosTemas()
    On Error GoTo MiErr
    'cargar lstProximos
    Dim strProximos As String ', TotTemas As Integer
    
    If tLST.GetLastIndex = 0 Then
        frmIndex.lblNexts.Caption = TR.Trad("Ingrese una moneda%97%y disfrute su%97%música preferida%99%")
        frmIndex.lblNexts.Alignment = 2 'centrado
        frmIndex.RollSONG.ReplaceIndex 1, TR.Trad("No hay%97%mas selecciones%98%No quedan canciones o videos para reproducir%99%")
    Else
        
        'volver a contar'
        PUBs.PubsEnLista = 0
        'el indice 0 no existe ni existira por eso el C=1


        Dim strLIST As String
        Dim HY As Long, HZ As Long
        HY = tLST.GetLastIndex
        strLIST = TR.Trad("Próximas selecciones: %99%") + CStr(HY) + vbCrLf
        If HY > 10 Then HY = 10
        For HZ = 1 To HY
            strLIST = strLIST + QuitarNumeroDeTema(tLST.GetElementListaFileName(HZ)) + vbCrLf
        Next HZ
        frmIndex.lblNexts.Caption = strLIST
        frmIndex.lblNexts.Alignment = 0

        TR.SetVars QuitarNumeroDeTema(tLST.GetElementListaFileName(1)), _
            tLST.GetElementListaLastFolder(1), PuestoN(tLST.GetElementListaPath(1))

        frmIndex.RollSONG.ReplaceIndex 1, _
            TR.Trad("proxima seleccion" + vbCrLf + _
            "%01%" + vbCrLf + "del disco" + vbCrLf + _
            "%02%" + vbCrLf + "Rank # %03%%98%" + _
            "La variable 1 es el nombre de la cancion, la " + _
            "2 es el nombre del disco y la tercera es un numero " + _
            "de posicion en el ranking%99%")
        
        'frmIndex.RollSONG.ReplaceIndex 1, "proxima seleccion" + vbCrLf + _
            tLST.GetElementListaFileName(1) + vbCrLf + _
            tr.trad("del disco") + vbCrLf + _
            tLST.GetElementListaLastFolder(1) + vbCrLf + _
            tr.trad("Rank # ") + PuestoN(tLST.GetElementListaPath(1))

'        Dim c As Long
'        For c = 1 To tLST.GetLastIndex
'            'no cargar las publicidades
'            strProximos = QuitarNumeroDeTema(tLST.GetElementListaFileName(c))
'
'            If tLST.GetTag(c) = "PUB" Then
'                'contador de publicidades en lista
'                PUBs.PubsEnLista = PUBs.PubsEnLista + 1
'            Else
'                frmIndex.lstProximos = frmIndex.lstProximos + CStr(c - PUBs.PubsEnLista) + "- " + strProximos + vbCrLf
'            End If
'        Next
'        'primero se escribe la lista y despues la primera linea
'        'esto para que sepa cuantas son publicidades!!!!
'        TotTemas = tLST.GetLastIndex - PUBs.PubsEnLista
'        'tengo que descontar as publicidades!!!!
'        frmIndex.lstProximos = "TEMAS PENDIENTES (" + _
'            CStr(TotTemas) + ")" + vbCrLf + frmIndex.lstProximos
        
    End If
    
    Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), "Globales.BAS" + ".acpi33"
    Resume Next

End Sub

Public Sub SetKeyState(ByVal Key As Long, ByVal State As Boolean)
    'ver si hace falta!
    'si ya esta apretada ..... salgo
    If (GetKeyState(Key) = 1) And State Then Exit Sub
    If (GetKeyState(Key) = 0) And State = False Then Exit Sub
    
    keybd_event Key, MapVirtualKey(Key, 0), KEYEVENTF_EXTENDEDKEY Or 0, 0
    keybd_event Key, MapVirtualKey(Key, 0), KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
      
    'verificar si quedo!!!
    If ((GetKeyState(Key) = 1) And (State = False)) Or ((GetKeyState(Key) = 0) And State) Then
        tERR.AppendSinHist "FailKB-LED:" + CStr(Key) + " / " + CStr(State)
    End If
    
End Sub

'Public Sub OnOffCAPS(vKey As KeyCodeConstants, PRENDER As Boolean)
'    Dim keys(255) As Byte
'    ' leer el estado actual del teclado
'    GetKeyboardState keys(0)
'    ' invertir el bit 0 de la tecla virtual en la que estamos interesados
'    ' keys(vKey) = keys(vKey) Xor 1
'    If PRENDER Then
'        keys(vKey) = 1
'    Else
'        keys(vKey) = 0
'    End If
'    ' forzar el nuevo estado del teclado
'    SetKeyboardState keys(0)
'End Sub

Public Function Tecla(n As Integer) As String
    Select Case n
        'las letras son iguales
        Case 65 To 90
            Tecla = Chr(n) + " :" + Trim(CStr(n))
        'los numeros tambien
        Case 48 To 57
            Tecla = Chr(n) + " :" + Trim(CStr(n))
        'el numpad debe escribir numeros (48-57)
        Case 96 To 105
            Tecla = Chr(n - 48) + " :" + Trim(CStr(n))
        Case 106
            Tecla = "* :106"
        Case 107
            Tecla = "+ :107"
        Case 108
            Tecla = "ENTER NumPad :108" + vbCrLf 'enter del key pad
        Case 109
            Tecla = "- NumPad :109"
        Case 110
            Tecla = ". NumPad :110"
        Case 111
            Tecla = "/ NumPad :111"
        Case 1
            Tecla = "Mouse IZQ :1"
        Case 2
            Tecla = "Mouse Der :2"
        Case 3
            Tecla = "CANCEL :3"
        Case 4
            Tecla = "Mouse MED :4"
        Case 8
            Tecla = "BACK :8"
        Case 9
            Tecla = "TAB :9"
        Case 12
            Tecla = "SUPR :12"
        Case 13
            Tecla = "ENTER :13"
        Case 16
            Tecla = "SHIFT :16"
        Case 17
            Tecla = "CTRL :17"
        Case 18
            Tecla = "ALT :18"
        Case 19
            Tecla = "PAUSA :19"
        Case 20
            Tecla = "Bloq MAY :20"
        Case 27
            Tecla = "ESC :27"
        Case 32
            Tecla = " (espacio) :32"
        Case 33
            Tecla = "PAGE UP :33"
        Case 34
            Tecla = "PAGE DOWN :34"
        Case 35
            Tecla = "HOME :35"
        Case 36
            Tecla = "END :36"
        Case 37
            Tecla = "IZQ :37"
        Case 38
            Tecla = "ARR :38"
        Case 39
            Tecla = "DER :39"
        Case 40
            Tecla = "ABJ :40"
        Case 41
            Tecla = "SELECT :41"
        Case 42
            Tecla = "PRINT SCR :42"
        Case 43
            Tecla = "EXECUTE :43"
        Case 44
            Tecla = "SNAPSHOT :44"
        Case 45
            Tecla = "INS :45"
        Case 46
            Tecla = "SUPR :46"
        Case 47
            Tecla = "AYUDA :47"
        Case 144
            Tecla = "BLOQ NUM :144"
            
        'faltan las Fs
    End Select
        
End Function

Public Sub CargarArchReini(ModoReini As String)
    
'asi era el anterior (datos al pedo) SEPRADO POR COMA .... UNA CAGADA !!
'D:\musica\Cuartetazo\Almafuerte - Toro Y Pampa\05 - Almafuerte - La Maquina De Picar Carne.Mp3,05 - Almafuerte - La Maquina De Picar Carne (mp3-Musica) / Almafuerte - Toro Y Pampa
'D:\musica\Cuartetazo\Almafuerte - Toro Y Pampa\06 - Almafuerte - Donde Esta Mi Corazon.Mp3,06 - Almafuerte - Donde Esta Mi Corazon (mp3-Musica) / Almafuerte - Toro Y Pampa
'D:\musica\Cuartetazo\Almafuerte - Toro Y Pampa\07 - Almafuerte - En El Siglo Del Gran Reviente.Mp3,07 - Almafuerte - En El Siglo Del Gran Reviente (mp3-Musica) / Almafuerte - Toro Y Pampa
    
    If ModoReini = "NADA" Then Exit Sub
    
    'ver si no se puede grabar
    If tLST.ListaGuardarADisco(GPF("casc1001")) = 1 Then
        tERR.AppendLog "NGRI:" + GPF("casc1001")
    End If
End Sub

Public Function STRceros(n As Variant, Cifras As Integer) As String
    'n es el numero y cifras es la cantidad final de cifras del str terminado
    'devuelve ej : para 232,6 = 000232 para 1902,12 = 000000001902
    'complaeta con ceroas adelante
    ' si n es mas lasgo que cifras devuelve el valor n sin ningun cero adelante
    Dim STRn As String
    STRn = Trim(CStr(n))
    Dim DIF As Integer
    DIF = Cifras - Len(STRn)
    If DIF > 0 Then
        Dim CEROstr As String
        CEROstr = String(DIF, "0")
        STRceros = CEROstr + STRn
    Else
        STRceros = STRn
    End If
End Function

Public Sub APAGAR_PC()
    Dim v As vWindows
    v = vW.GetVersion
    Select Case v
    Case Win98, Win98SE, WinME
        Shell "rundll32 user.exe,exitwindows"
    Case Win2000, WinNT4, WinXp, WinVista
        Shell "Shutdown -s -t 0" 'el -s es shutdowsn y el -r restart
    
    End Select
End Sub

Public Sub REINICIAR_PC()
    Dim v As vWindows
    v = vW.GetVersion
    Select Case v
    Case Win98, Win98SE, WinME
        Shell "rundll32 user.exe,exitwindowsexec"
    Case Win2000, WinNT4, WinXp, WinVista
        Shell "Shutdown -r -t 0" 'el -s es shutdowsn y el -r restart
    
    End Select
End Sub

Public Sub VerClaves(Clave As String)
    Select Case Clave
        Case ClaveClose
            Clave = "11111222223333344444" 'anular para que no se siga cargando
                'cerrar 3pm
                YaCerrar3PM
            
            Case ClaveConfig
                Clave = "11111222223333344444" 'anular para que no se siga cargando
                'entrar en configuracion
                frmConfig.Show 1
    End Select
    If Left(Clave, 19) = ClaveCredit Then
        'cargar creditos
        'ver cuantos son
        Dim NewCredit As Integer
        NewCredit = Val(Right(Clave, 1))
        CREDITOS = CREDITOS + NewCredit
        'no suma contador de creditos
        EscribirArch1Linea GPF("creditosactuales"), Trim(CStr(CREDITOS))
        
        ShowCredits
        
        Clave = "11111222223333344444" 'anular para que no se siga cargando
    End If
End Sub

Public Sub VarCreditos(VarCre As Single, Optional SumaCont As Boolean = True, _
    Optional SumaValidar As Boolean = True, Optional UpdateCreditos As Boolean = True)
    
    tERR.Anotar "B233|" + CStr(VarCre)
    
    
    CREDITOS = CREDITOS + VarCre
    tERR.Anotar "B234|" + CStr(CREDITOS)
    '-------------------------------------------------------
    'si es menor que cero es por que el tipo puso un tema
    'la funcion sumarcont... si puede tener negativos o ceros por ejemplo para
    'reiniciar el contador reiniciable. En el caso de esta funcion VarCreditos
    'hay valores negativos cuando se usa una cancion y se descuenta el credito dispo
    'nible, esto no implica que se cambie el contador reiniciable ni el historico
    If VarCre > 0 Then
        'no entiendo por que estaba aca ya que al iniciar 3pm
        'manda una variacion positiva con el total con que se cerro para arrancar
        'SumarContadorCreditos CLng(VarCre)
    End If
    '-------------------------------------------------------
    'grabar cant de creditos
    If SumaCont Then
        EscribirArch1Linea GPF("creditosactuales"), Trim(CStr(CREDITOS))
        If VarCre > 0 Then SumarContadorCreditos CLng(VarCre)
    End If
    
    tERR.Anotar "acei", CreditosValidar, CREDITOS
    
    If VarCre < 0 And SumaValidar Then
        CreditosValidar = CreditosValidar - VarCre 'al restarlo se suma por que es negativo
        EscribirArch1Linea GPF("radliv"), CStr(CreditosValidar)
        
        If RavI > 2 Then 'si ya aviso una vez
            Dim F As String
            '
            'recuerde valida su equipo
            F = dcr("ETAhtnC15tnESvs+YXjlyltqZ+l+IFLkU8aIs8eqb6M1gdSxrAWJWg==")
                
            frmIndex.RollCRED.ReplaceIndex 0, F
            frmIndex.RollCRED.ReplaceIndex 1, F
            frmIndex.RollCRED.ReplaceIndex 2, F
            
            frmIndex.RollSONG.ReplaceIndex 0, F
            frmIndex.RollSONG.ReplaceIndex 1, F
            frmIndex.RollSONG.ReplaceIndex 2, F
        End If
    End If
    
    DefinePrecios VarCre, PrecNowAudio, PrecNowVideo
    
    If UpdateCreditos Then
        siganlIn = siganlIn - 1
        tERR.Anotar "B234b" + CStr(siganlIn)
        frmIndex.List1.List(9) = "PNA=" + CStr(PrecNowAudio)
        frmIndex.List1.List(10) = "PNV=" + CStr(PrecNowVideo)
        ShowCredits
    End If
    
    tERR.Anotar "B235|" + CStr(CREDITOS * PrecioBase / TemasPorCredito)
    
    
End Sub

Public Sub DefinePrecios(ByVal VC As Single, ByRef PNA As Single, PNV As Single)

    If VC <= 0 Then
        'si se ejecutaron canciones o videos y los creditos llegan hasta un valor
        'menor de una cancion en la maxima oferta disponible
        'enonces el precio vuelve a lo normal
        If CREDITOS < GetPrecioAudioMasBarato And CREDITOS < GetPrecioVideoMasBarato Then
            CREDITOS = 0
        End If
        
        If CREDITOS < GetPrecioAudioMasBarato Then
            If CreditosCuestaTema(0) > 0 Then
                PNA = CreditosCuestaTema(0)
            Else
                PNA = 1000000 'si no entra en ninguno ponemos precio inalcanzable
            End If
        End If
        
        If CREDITOS < GetPrecioVideoMasBarato Then
            If CreditosCuestaTemaVIDEO(0) > 0 Then
                PNV = CreditosCuestaTemaVIDEO(0)
            Else
                PNV = 1000000 'si no entra en ninguno ponemos precio inalcanzable
            End If
        End If
        
    End If
    
    'si se pusieron monedas entonces el precio puede cambiar
    If VC > 0 Then
        'si puso varias monedas bajar los precios si corresponde
        If CREDITOS >= CreditosCuestaTema(0) And CreditosCuestaTema(0) > 0 Then
            PNA = CreditosCuestaTema(0)
            '(porque son los creditos xa 2 canciones)
        End If
        
        If CREDITOS >= CreditosCuestaTema(1) And CreditosCuestaTema(1) > 0 Then
            PNA = tbrFIX(Round(CreditosCuestaTema(1) / 2, 4), 2)
            '(porque son los creditos xa 2 canciones)
        End If
        
        If CREDITOS >= CreditosCuestaTema(2) And CreditosCuestaTema(2) > 0 Then
            PNA = tbrFIX(Round(CreditosCuestaTema(2) / 3, 4), 2)
            '(porque son los creditos xa 3 canciones)
        End If
        
        If CREDITOS >= CreditosCuestaTemaVIDEO(0) And CreditosCuestaTemaVIDEO(0) > 0 Then
            PNV = CreditosCuestaTemaVIDEO(0)
            '(porque son los creditos xa 2 canciones)
        End If
        
        If CREDITOS >= CreditosCuestaTemaVIDEO(1) And CreditosCuestaTemaVIDEO(1) > 0 Then
            PNV = tbrFIX(Round(CreditosCuestaTemaVIDEO(1) / 2, 4), 2)
            '(porque son los creditos xa 2 canciones)
        End If
        
        If CREDITOS >= CreditosCuestaTemaVIDEO(2) And CreditosCuestaTemaVIDEO(2) > 0 Then
            PNV = tbrFIX(Round(CreditosCuestaTemaVIDEO(2) / 3, 4), 2)
            '(porque son los creditos xa 3 canciones)
        End If
    End If
    
    'me canse, aquiva una negrada
    If CreditosCuestaTema(0) = 0 And CreditosCuestaTema(1) = 0 And CreditosCuestaTema(2) = 0 Then
        PNA = 0: PNV = 0
    Else
        If PNA = 0 Then PNA = 1000000
        If PNV = 0 Then PNV = 1000000
    End If
    
End Sub

Public Function tbrFIX(n As Single, DecimalesTruncar As Long) As Single
    'truncar a una X cantidad de decimales
    Dim SN As String
    'tratarlo como caracter es mas facil
    SN = CStr(n)
    'si es entero entonces salgo, no hay nada que hacer
    Dim TieneDec As Boolean
    If InStr(SN, ",") > 0 Then TieneDec = True
    If InStr(SN, ".") > 0 Then TieneDec = True
    If TieneDec = False Then
        tbrFIX = n
        Exit Function
    End If
    
    Dim AA As Long, Largo As Long, BB As Long
    BB = 0 'cuenta la cantidad de decimales
    Largo = Len(SN)
    Dim EmpezoDec As Boolean
    EmpezoDec = False
    For AA = 1 To Largo
        If EmpezoDec Then BB = BB + 1
        'si se llega al total cortar ahi
        If BB = DecimalesTruncar Then
            tbrFIX = CSng(Mid(SN, 1, AA))
            Exit Function
        End If
        If Mid(SN, AA, 1) = "." Or Mid(SN, AA, 1) = "," Then EmpezoDec = True
    Next AA
    'si sale de aqui sin haber salido antes es porque no llega a la cantida deseada
    tbrFIX = n
End Function

Public Function GetPrecioAudioMasBarato() As Single
    'saber el precio mas barato me sirve para saber cuando ya no hay
    'posibilidad de poner mas canciones, en ese caso vuelve al precio normal
    
    GetPrecioAudioMasBarato = 0
    
    If CreditosCuestaTema(2) > 0 Then
        GetPrecioAudioMasBarato = CreditosCuestaTema(2) / 3
        Exit Function
    End If
    
    If CreditosCuestaTema(1) > 0 Then
        GetPrecioAudioMasBarato = CreditosCuestaTema(1) / 2
        Exit Function
    End If
    
    If CreditosCuestaTema(0) > 0 Then
        GetPrecioAudioMasBarato = CreditosCuestaTema(0)
        Exit Function
    End If
        
End Function

Public Function GetPrecioVideoMasBarato() As Single
    'saber el precio mas barato me sirve para saber cuando ya no hay
    'posibilidad de poner mas canciones, en ese caso vuelve al precio normal
    
    GetPrecioVideoMasBarato = 0
    
    If CreditosCuestaTemaVIDEO(2) > 0 Then
        GetPrecioVideoMasBarato = CreditosCuestaTemaVIDEO(2) / 3
        Exit Function
    End If
    
    If CreditosCuestaTemaVIDEO(1) > 0 Then
        GetPrecioVideoMasBarato = CreditosCuestaTemaVIDEO(1) / 2
        Exit Function
    End If
    
    If CreditosCuestaTemaVIDEO(0) > 0 Then
        GetPrecioVideoMasBarato = CreditosCuestaTemaVIDEO(0)
        Exit Function
    End If
        
End Function

Public Sub Pintar_fBoton(FRM As Form)
    Dim CTR
    For Each CTR In FRM.Controls
        If TypeOf CTR Is fBoton Then
            CTR.BackColor = ColUnSel
            CTR.BackColor2 = Col2UnSel
        End If
    Next
End Sub

Public Sub AjustarFRM(FRM As Form, HechoParaTwipsHoriz, HechoParaTwipsVertical)
    'ajusta el formulario a la pantalla. JOYA, JOYA
    'HechoParaPixelHoriz quiere decir que el tamaño original entra justo en
    'por ej 800x600 si el valor es 12000
    
    'SET 2007 YA LA PROPORCION ENTRE ANCHO Y ALTO NO ES IGUAL!! MONITORES
    Dim ActTwipsHoriz As Long
    ActTwipsHoriz = Screen.Width
    
    Dim ActTwipsVertical As Long
    ActTwipsVertical = Screen.Height
    
    Dim MultiplicadorW As Double
    Dim MultiplicadorH As Double
    MultiplicadorW = ActTwipsHoriz / HechoParaTwipsHoriz
    MultiplicadorH = ActTwipsVertical / HechoParaTwipsVertical
    
    Dim CTR As Control
    
    For Each CTR In FRM.Controls
        If CTR.Name = "cmdPagAt" Then GoTo sig
        If CTR.Name = "cmdPagAd" Then GoTo sig
        If CTR.Name = "pVU1" Then GoTo sig
        If CTR.Name = "pVU2" Then GoTo sig
        If CTR.Name = "pVU3" Then GoTo sig
        If CTR.Name = "pVU4" Then GoTo sig
        If CTR.Name = "imgSelec2" Then GoTo sig
        If CTR.Name = "cmdTocuhArriba2" Then GoTo sig
        If CTR.Name = "cmdTouchAbajo2" Then GoTo sig
        If CTR.Name = "cmdTocuhArriba" Then GoTo sig
        If CTR.Name = "cmdTouchAbajo" Then GoTo sig
        'algunos controles no tienen algunas propiedades
        On Local Error Resume Next
        
'        'los objetos tipo image si no se les hace stretch no sirv cambiar su tamaño
'        If CTR.Stretch = False Then
'            'tERR.AppendLog "noChi", CTR.Name
'            GoTo sig
'        End If
        
        tAs = CTR.Name
        CTR.Height = CTR.Height * MultiplicadorH
        CTR.Width = CTR.Width * MultiplicadorW
        CTR.Top = CTR.Top * MultiplicadorH
        CTR.Left = CTR.Left * MultiplicadorW
        CTR.Font.Size = CTR.Font.Size * MultiplicadorH '(si son distintos este seguro es menor por las nuevas definiciones que existen)
        CTR.X1 = CTR.X1 * MultiplicadorW
        CTR.X2 = CTR.X2 * MultiplicadorW
        CTR.Y1 = CTR.Y1 * MultiplicadorH
        CTR.Y2 = CTR.Y2 * MultiplicadorH
sig:
    Next

End Sub

Public Function LeerConfig(Conf As String, ValDefault As String) As String
    
    'leer el archivo de configuracion y devolver valor
    LeerConfig = "NO EXISTE"
    
    On Local Error GoTo errLC
    
    Dim TXT As String, CFG As String, RST As String
    If fso.FileExists(GPF("config")) Then
        Set TE = fso.OpenTextFile(GPF("config"), ForReading, False)
            Dim FullConfig As String
            FullConfig = TE.ReadAll
        TE.Close
        'desencriptar
        FullConfig = Encriptar(FullConfig, True)
        'escribir un temporal desencriptado
        Set TE = fso.CreateTextFile(AP + "tmp.tbr", True)
            TE.Write FullConfig
        TE.Close
        Set TE = fso.OpenTextFile(AP + "tmp.tbr", ForReading, False)
            Do While Not TE.AtEndOfStream
                TXT = TE.ReadLine
                CFG = Trim(txtInLista(TXT, 0, "=")) 'la configuracion
                If UCase(CFG) = UCase(Conf) Then
                    'por si hay algun "=" en la respuesta
                    RST = Mid(Trim(TXT), Len(CFG) + 2, Len(TXT) - Len(CFG) + 1)
                    
                    'este parece feo pero anduvo por años
                    'RST = Trim(txtInLista(TXT, 1, "=")) 'el valor
                    
                    'y si esta vacio!!!!
                    If RST <> "" Then
                        LeerConfig = RST
                    Else
                        LeerConfig = ValDefault
                    End If
                    Exit Do
                End If
            Loop
        TE.Close
        'borrar el temporal
        fso.DeleteFile AP + "tmp.tbr", True
    End If
    If LeerConfig = "NO EXISTE" Then
        'cargar el valor por defecto
        LeerConfig = ValDefault
    End If
    
    
    Exit Function
    
errLC:
    tERR.AppendLog "ELC6565:" + CStr(Conf) + ":" + CStr(Value)
    
End Function

Public Function ChangeConfig(Conf As String, newValue As String) As Long
    'me tomo la molestia de ver si el valor ya estaba asi n o grabo todo de nuevo
    'DEVUELVE
    '-1 error no definido
    '0 ok
    '1 estaba en ese valor, no hizo nada
    '2 no estab y la cree
    'leer el archivo de configuracion y devolver valor
    
    On Local Error GoTo errChgConf
    ChangeConfig = 2
    
    Dim TXT As String, CFG As String, RST As String
    If fso.FileExists(GPF("config")) Then
        Set TE = fso.OpenTextFile(GPF("config"), ForReading, False)
            Dim FullConfig As String
            FullConfig = TE.ReadAll
        TE.Close
        'desencriptar
        FullConfig = Encriptar(FullConfig, True)
    Else
        'si no esta el archivo (el primer inicio!!!!)
        'este se va creando con las configuraciones que no
        'estan en sus valores predereminados solamente!!!
        FullConfig = "PrimerInicio=1"
    End If
    'escribir un temporal desencriptado
    Set TE = fso.CreateTextFile(AP + "tmp.tbr", True)
        TE.Write FullConfig
    TE.Close
    
    Dim ValToReWrite As String 'leeo todo para que todo quede igual menos lo que cambio!!!
    ValToReWrite = ""
    
    Set TE = fso.OpenTextFile(AP + "tmp.tbr", ForReading, False)
        Do While Not TE.AtEndOfStream
            TXT = TE.ReadLine
            
            CFG = Trim(txtInLista(TXT, 0, "=")) 'la configuracion
            If UCase(CFG) = UCase(Conf) Then
                RST = Trim(txtInLista(TXT, 1, "=")) 'el valor
                'y si esta vacio!!!!
                If RST <> newValue Then
                    ValToReWrite = ValToReWrite + CFG + "=" + newValue + vbCrLf
                    ChangeConfig = 0
                Else
                    'al pedo, no voy a grabar nada!
                    TE.Close
                    fso.DeleteFile AP + "tmp.tbr", True
                    ChangeConfig = 1
                    Exit Function
                End If
                
                '***************************************
                'Exit Do
                'NOOOOOOOOOOOOOOOOOOOOOOOOOOOO hace que se corte hasta
                'aqui y no grabe lo demás !!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                '***************************************
            Else 'pasa derecho como estaba a lo siguiente
                'no perder el renglon
                ValToReWrite = ValToReWrite + TXT + vbCrLf
            End If
        Loop
        
        'ver si no existia!!!
        If ChangeConfig = 2 Then
            ValToReWrite = ValToReWrite + Conf + "=" + newValue + vbCrLf
        End If
    TE.Close
    
    'borrar el temporal
    fso.DeleteFile AP + "tmp.tbr", True
    
    'encriptar
    ValToReWrite = Encriptar(ValToReWrite, False)
    'grabar el kilombo
    Set TE = fso.CreateTextFile(GPF("config"), True)
        TE.Write ValToReWrite
    TE.Close
    'hacer una copia de seguridad cada vez que haya cambios
    fso.CopyFile GPF("config"), GPF("config2")
    
    Exit Function
    
errChgConf:
    tERR.AppendLog "ECC322:" + CStr(cong) + ":" + CStr(Value)
    ChangeConfig = -1
End Function

Public Function Encriptar(Valor, UnEncrypt As Boolean) As String
    'con esta funcion se puede encriptar y desencriptar
    'la uso para el GPF("config")
    
    'para saber si estoy leyendo algo encrytado le pongo algo identificativo
    Dim IdEstaEncryptado As String
    IdEstaEncryptado = "RMLVF"
    'encripta cualquier cosa y la transforma en string
    Dim ToEncrypt As String
    ToEncrypt = CStr(Valor)
    
    Dim Largo As Long, IND As Long, Letra As String, LetraE As String
    Dim FullE As String 'resultado de la encryptacion
    'ver si lo que se ingreso ya esta encrptado
    If UCase(Left(ToEncrypt, Len(IdEstaEncryptado))) = IdEstaEncryptado Then
        'ya esta encriptado
        If UnEncrypt Then
            'DESNCRIPTAR!!!
            'cambiar uno por uno los codigos
            Largo = Len(ToEncrypt)
            'empeiza despues del marcador
            For IND = Len(IdEstaEncryptado) + 1 To Largo
                Letra = Mid(ToEncrypt, IND, 1)
                'pasar todo a una letra distinta. Los saltos de carro no usarlos
                If Asc(Letra) = 0 Then Letra = "0"
                Select Case Letra
                    Case "0"
                        LetraE = vbCrLf
                    Case Else
                        LetraE = Chr(Asc(Letra) - 10)
                End Select
                FullE = FullE + LetraE
            Next
            Encriptar = FullE
        Else
            'no se puede encyprtar lo encryptado
            Encriptar = ToEncrypt
            Exit Function
        End If
    Else
        If UnEncrypt Then
            'no se puede desdencryptar lo desencryptado
            Encriptar = ToEncrypt
            Exit Function
        Else
            'Encriptar!!!!
            'cambiar uno por uno los codigos
            Largo = Len(ToEncrypt)
            For IND = 1 To Largo
                Letra = Mid(ToEncrypt, IND, 1)
                'pasar todo a una letra distinta. Los saltos de carro no usarlos
                Select Case Letra
                    Case vbCrLf ' Or vbCr
                        LetraE = "0"
                    Case Else
                        LetraE = Chr(Asc(Letra) + 10)
                End Select
                FullE = FullE + LetraE
            Next
            Encriptar = IdEstaEncryptado + FullE
        End If
        
    End If
    
End Function

Public Function QuitarNumeroDeTema(TemaFull As String) As String
    On Error GoTo MiErr
    'si es un archivo corto ni lo toco!!!
    'en general es porque el nombre del tema es un número!!!
    If Len(TemaFull) <= 4 Then
        QuitarNumeroDeTema = TemaFull
        Exit Function
    End If
    tERR.Anotar "004-0001", TemaFull
    Dim tmpTema As String
    tmpTema = TemaFull
    'ver si hay numeros adelante y si hay quitarselos
    Dim NumersoAlInicio As Long
    NumersoAlInicio = 0
    If IsNumeric(Mid(TemaFull, 1, 1)) Then NumersoAlInicio = NumersoAlInicio + 1
    If IsNumeric(Mid(TemaFull, 2, 1)) Then NumersoAlInicio = NumersoAlInicio + 1
    If IsNumeric(Mid(TemaFull, 3, 1)) Then NumersoAlInicio = NumersoAlInicio + 1
    tERR.Anotar "004-0002"
    If NumersoAlInicio > 0 Then
        tmpTema = Trim(Right(TemaFull, Len(TemaFull) - 3))
        'ver si quedo con esto
        Dim A As Long
        For A = 1 To 4
            If Mid(tmpTema, A, 1) = "-" _
                Or Mid(tmpTema, A, 1) = "_" _
                Or Mid(tmpTema, A, 1) = "/" _
                Or Mid(tmpTema, A, 1) = "@" _
                Or Mid(tmpTema, A, 1) = "[" _
                Or Mid(tmpTema, A, 1) = "]" _
                Or Mid(tmpTema, A, 1) = "(" _
                Or Mid(tmpTema, A, 1) = ")" Then
                tmpTema = Trim(Right(tmpTema, Len(tmpTema) - 1))
            End If
        Next
        
    End If
    
    QuitarNumeroDeTema = tmpTema
    
    tERR.Anotar "004-0003", tmpTema
    
    Exit Function
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), "GLOBALES.bas" + ".acpl"
    Resume Next
    'If frmIndex.MP3.IsPlaying = False Then EMPEZAR_SIGUIENTE
End Function
Public Sub InfoDisco(LBL As Label)
    Dim TotDisco, TotFree1, TotFree2, Serial As String, VolName As String
    'ver en que disco esta instalado
    Dim DiscoInst3PM As String
    DiscoInst3PM = Left(AP, 1)
    DiscoInst3PM = DiscoInst3PM + ":\"
    TotDisco = Round(fso.Drives(DiscoInst3PM).TotalSize / 1024 / 1024, 2)
    TotFree1 = Round(fso.Drives(DiscoInst3PM).AvailableSpace / 1024 / 1024, 2)
    TotFree2 = Round(fso.Drives(DiscoInst3PM).FreeSpace / 1024 / 1024, 2)
    Serial = fso.Drives(DiscoInst3PM).SerialNumber
    VolName = fso.Drives(DiscoInst3PM).VolumeName
    
    Dim PorcLibre As Double
    PorcLibre = Round(TotFree1 / TotDisco * 100, 2)
    
    TR.SetVars VolName, TotDisco, TotFree1, (TotFree1 / TotDisco) * 100
    
    LBL.Caption = TR.Trad("Informacion del disco (%01%)" + vbCrLf + _
        "Total disco: %02% MB" + vbCrLf + _
        "Total Disponible: %03% MB" + vbCrLf + _
        "Porcentaje libre: %04% % %98%La variable 1 es la etiqueta " + _
        "o nombre de una de las particiones del disco" + vbCrLf + _
        "La variable 2 es el total de MegaBytes del disco mencionado" + vbCrLf + _
        "La variable 3 es el total de MegaBytes libres del disco mencionado" + vbCrLf + _
        "La variable 4 es el porcentaje disponible del disco mencionado" + vbCrLf + _
        "%99%")
End Sub

Public Function InfoDisco2(LetraDisco As String, ByRef MbTotal As Long, _
    MbLibre As Long, PorcFree As Single) As String
    
    Dim TotDisco, TotFree1, TotFree2, Serial As String, VolName As String
    'ver en que disco esta instalado
    LetraDisco = LetraDisco + ":\"
    
    '--------------MANUEL----------------------------------------
    If fso.Drives(LetraDisco).IsReady = False Then Exit Function
    '------------------------------------------------------------
    
    
    TotDisco = Round(fso.Drives(LetraDisco).TotalSize / 1024 / 1024, 2)
    TotFree1 = Round(fso.Drives(LetraDisco).AvailableSpace / 1024 / 1024, 2)
    TotFree2 = Round(fso.Drives(LetraDisco).FreeSpace / 1024 / 1024, 2)
    Serial = fso.Drives(LetraDisco).SerialNumber
    VolName = fso.Drives(LetraDisco).VolumeName
    
    Dim PorcLibre As Double
    PorcLibre = Round(TotFree1 / TotDisco * 100, 2)
    
    MbTotal = TotDisco
    MbLibre = TotFree1
    PorcFree = PorcLibre
    
    TR.SetVars _
        LetraDisco + "(" + VolName + ")=" + CStr(TotDisco), _
        TotFree1, _
        PorcLibre
    InfoDisco2 = TR.Trad("%01% MB totales y %02% MB libres (%03% %)%99%")
End Function

Public Sub VerSiTocaPUB()
    'despues de ejecutar un tema desde Temas de Disco, index o Top10
    'toca saber si se agrega una pub a la lista!!
    'pasar a la lista de reproducción
    'SOLO SI HAY PUBLICIDADES
    'NO VA A FALTAR ALGUN IDIOTA QUE HABILITE Y NO COLOQUE PUBS!!!!
    If PUBs.HabilitarPublicidadesMp3Vid And PUBs.TotalPUBs > 0 Then
        'indicarle al PUB que paso otro tema
        PUBs.ContadorTemas = PUBs.ContadorTemas + 1
        'ver si ya corresponde
        If PUBs.SonarPublicidadesCada <= PUBs.ContadorTemas Then
            'poner en cero el contador
            PUBs.ContadorTemas = 0
            
            'mandar a la lista!!!
            PUBs.UltimaReproducida = PUBs.UltimaReproducida + 1
            'si termino que empieze de vuelta. Siempre empieza en el 1
            'el cero esta en blanco!!!
            'no es >=!!!! es solo mayor si no no rep la ultima!!!
            If PUBs.UltimaReproducida > PUBs.TotalPUBs Then PUBs.UltimaReproducida = 1
            
            'INDICAR CUAL SE EJECUTA
            Dim ArchPub As String
            ArchPub = PUBs.ArchsPubs(PUBs.UltimaReproducida)
            
            'otra seguridad mas
            If fso.FileExists(ArchPub) Then
                'pasar a la lista de reproducción
                tLST.ListaAdd ArchPub, "PUB"
                'escribir la lista en pantalla
                'si no lo hago, no se actualiza los numeros de los que falta!!!!
                'aqui se fija cuantos temas quedas y resta la publicidad!!!!
                CargarProximosTemas
                'creo que no hace falta
                'graba en reini.tbr los datos que correspondan por si se corta la luz
                CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
            End If
        End If
    End If
End Sub

Public Sub ShowCredits()
    Select Case ShowCreditsMode
        Case 0 'modo plata
            If CREDITOS = 0 Then
                frmIndex.lblCreditos.Caption = TR.Trad("Inserte moneda" + _
                    "%98%aviso de que no hay credito para ejecutar%99%")
            Else
                frmIndex.lblCreditos.Caption = TR.Trad("Credito" + _
                "%98%aviso de credito disponible%99%") + _
                    " " + CStr(FormatCurrency(CREDITOS * PrecioBase / TemasPorCredito, _
                    , , , vbFalse))
            End If
            
        Case 1 'modo créditos
            If CREDITOS = 0 Then
                frmIndex.lblCreditos = TR.Trad("Inserte moneda%99%" + _
                    "%98%aviso de que no hay credito para ejecutar%99%")
            Else
                frmIndex.lblCreditos = TR.Trad("Credito%99%") + " " + _
                    CStr(Round(CREDITOS, 2))
            End If
    End Select
End Sub

Public Function FindIndexOfLst(SplitSpace1 As String, CMB As ComboBox) As Long
    'busca en un combobox el elemnto que tenga al
    'inicio la secuencia buscada
    'devuelve el indice del combo
    
    If CMB.ListCount = -1 Then
        FindIndexOfLst = -1
        Exit Function
    End If
    Dim Largo As Long
    Largo = Len(SplitSpace1)
    Dim A As Long
    For A = 0 To CMB.ListCount - 1
        If Left(CMB.List(A), Largo) = SplitSpace1 Then
            FindIndexOfLst = A
            Exit Function
        End If
    Next A
End Function

Public Sub QuitaIndiceMatriz(mtxToBorrar, iQuitar As Long)
    
    Dim J As Long
    For J = iQuitar To UBound(mtxToBorrar) - 1
        mtxToBorrar(J) = mtxToBorrar(J + 1)
    Next J
    
    J = UBound(mtxToBorrar) 'creo que le corresponde estar en este valor, pero por las dudas ...
    
    ReDim Preserve mtxToBorrar(J - 1)
    
End Sub

Public Sub SumarMatriz(MatrizAcumuladora() As String, MatrizAgregada() As String)

    Dim J As Long, A As Long
    
    'si es la primera suma me quedaria el indice cero al pedo!!!
    If UBound(MatrizAcumuladora) = 0 Then
        'ver si viene vacio ese cero o con el ranking si estuviera asi configurado
        If Len(MatrizAcumuladora(0)) > 2 Then
            J = 1
        Else
            J = 0
        End If
        YaEmpezo = True
    Else
        J = UBound(MatrizAcumuladora) + 1
    End If
    
    For A = 1 To UBound(MatrizAgregada)
        
        ReDim Preserve MatrizAcumuladora(J)
        MatrizAcumuladora(J) = MatrizAgregada(A)
        
        'frmINI.lblINI.Caption = TR.Trad("Ordenando...%99%") + MatrizAgregada(A)
        'frmINI.lblINI.Refresh
        'frmINI.PBar.Width = (frmINI.lblINI.Width * A / UBound(MatrizAgregada)) Mod frmINI.lblINI.Width
        
        J = J + 1
    Next A

End Sub

Public Sub CaminoError(Ubic As String)
    'aqui se van acumulando los ultimos identificadores enviados
    'Ubic es un identificador mas
    
    AcumCaminoError = AcumCaminoError + " " + Ubic
    If Len(AcumCaminoError) > 90 Then
        AcumCaminoError = Right(AcumCaminoError, 90)
    End If
    LineaError = AcumCaminoError
End Sub

Public Function GetParam3PM(I As Long) As String
    'devuelve los comandos aplicados luego del exe
    Dim SP() As String
    SP = Split(Cs)
    
    If I > UBound(S) Then
        GetParam = ""
    Else
        GetParam = SP(I)
    End If
    
End Function

Public Function FindParam3PM(txtToFind As String) As String
    'se fija si determinado parametro existe, devuelve el valor luego del igual
    
    Dim SP() As String, AA As Long
    'ademas de los parametros comunes tiene los de la configuracion
    Cs = Cs + " " + LeerConfig("plusparam", "")
    SP = Split(Cs)
    
    FindParam3PM = "999999" 'valor si el parametro no esta
    
    Dim SP2() As String
    For AA = 0 To UBound(SP)
        If SP(AA) <> "" Then
            SP2 = Split(SP(AA), "=")
            If LCase(SP2(0)) = LCase(txtToFind) Then
                FindParam3PM = SP2(1)
                Exit For
            End If
        End If
    Next AA
    
End Function

Public Function LTE(I As Long) As Long  'llego tecla especial
    'i es el indice de tecla especial
    'La 0 es la tecla Q. o sea la principal entrada de moneda
    'La 1 es la tecla S. o sea la entrada de moneda secundaria
    
    'devuelve el acumulado por si hiciera falta
    'ver inserciones no humanas (tan rapidas como monedero)
    'la primera inicia un reloj que se entera cuando pararon de llegar
    
    'MOCAAAAASO cuando pasa la media noche timer es menor que TimeLastCoin(i)
    'por lo tanto se queda esperando !!!!
    If Timer < TimeLastCoin(I) Then
        TimeLastCoin(I) = Timer
        CoinMuyJuntosAcum(I) = 1
        Exit Function
    End If
    
    If Timer - TimeLastCoin(I) < (TimeMaxSeparacion(I) / 1000) Then
        CoinMuyJuntosAcum(I) = CoinMuyJuntosAcum(I) + 1
        EsperarFinTE I
        'trash, sacar este codigo una vez resuelto
        tERR.AppendSinHist "PasoLte:" + CStr(CoinMuyJuntosAcum(I)) + "/" + CStr(Timer) + "/" + CStr(TimeLastCoin(I))
    Else
        'el reloj debe detectarlo para saber a cuanto llego
        'y desde alli ponerlo en cero
        CoinMuyJuntosAcum(I) = 1
        'trash, sacar este codigo una vez resuelto
        tERR.AppendSinHist "IniLte:" + CStr(Timer) + "/" + CStr(TimeLastCoin(I))
    End If
    
    TimeLastCoin(I) = Timer
    LTE = CoinMuyJuntosAcum(I)
    'wLTE CoinMuyJuntosAcum(i)
End Function

'esperar X desde la ultima tecla especial para ver si termina o no
Private Sub EsperarFinTE(I As Long)  'esperar tecla especial hasta terminar
    
    Dim LastC As Long
    'me quedo esperando que pase el tiempo
    Do
        DoEvents: DoEvents
        If (Timer - TimeLastCoin(I)) > (TimeMaxSeparacion(I) / 1000) Then Exit Do
        
        'MOCAAAAASO cuando pasa la media noche timer es menor que TimeLastCoin(i)
        'por lo tanto se queda esperando !!!!
        If Timer <= TimeLastCoin(I) Then Exit Do 'NUNCA DESPUES DE LA MEDIANOCHE !!!
        
    Loop
    
    TerminoLTE I
    CoinMuyJuntosAcum(I) = 0
    TimeLastCoin(I) = 0
End Sub

Private Sub TerminoLTE(I As Long)
    'cuando dejo de llegar la tecla especial
    
    'si el valor asigndo estaba en cero se ignora y no hay reemplazo
    
    Dim J As Long
    If I = 1 Then
        For J = 1 To UBound(ValoresATransformar1)
            'si los valores que llegaron son los previstos como fallas ==>
            If CoinMuyJuntosAcum(I) = J Then
                'poner ValoresATransformar(J)-j mas señales a la tecla especial indicada
                'mandar esa misma señal las veces que falta
                If ValoresATransformar1(J) > 0 Then
                    
                    'poner los creditos que faltaron
                    'agregado 11/06/2009. Si es negativo !? me fije en la funcion varcreditos y no se registraia algo muy pulenta
                    If ValoresATransformar1(J) <> J Then
                        VarCreditos CSng(TemasPorCredito * (ValoresATransformar1(J) - J))
                        tERR.AppendSinHist "FaltoCR:" + CStr(ValoresATransformar1(J)) + "/" + CStr(J)
                    End If
                    
                    'MsgBox "faltaron:" + CStr(ValoresATransformar1(J) - J) + _
                        vbCrLf + "TLE:" + CStr(i) + vbCrLf + _
                        "J=" + CStr(J) + vbCrLf + _
                        CStr(ValoresATransformar1(J))
                        
                End If
                Exit For
            End If
        Next J
        CoinMuyJuntosAcum(I) = 0
    End If
    
    If I = 2 Then
        For J = 1 To UBound(ValoresATransformar2)
            'si los valores que llegaron son los previstos como fallas ==>
            If CoinMuyJuntosAcum(I) = J Then
                'poner ValoresATransformar(J)-j mas señales a la tecla especial indicada
                'mandar esa misma señal las veces que falta
                If ValoresATransformar2(J) > 0 Then
                    
                    If ValoresATransformar2(J) <> J Then
                        'poner los creditos que faltaron
                        VarCreditos CSng(CreditosBilletes * (ValoresATransformar2(J) - J))
                        tERR.AppendSinHist "FaltoCR:" + CStr(ValoresATransformar2(J)) + "/" + CStr(J)
                    End If
                    'MsgBox "faltaron:" + CStr(ValoresATransformar2(J) - J) + _
                        vbCrLf + "TLE:" + CStr(i) + vbCrLf + _
                        "J=" + CStr(J) + vbCrLf + _
                        CStr(ValoresATransformar2(J))
                End If
                Exit For
            End If
        Next J
        CoinMuyJuntosAcum(I) = 0
    End If
End Sub

Private Sub CargarValoresTeclasEspeciales()
    'al inicio del sistema para empezar
    Dim TMP As String, SP() As String
    Dim TE8 As TextStream
    
    ReDim Preserve ValoresATransformar1(20)
    ReDim Preserve ValoresATransformar2(20)
    
    If fso.FileExists(GPF("rempmon45")) Then
        Set TE8 = fso.OpenTextFile(GPF("rempmon45"), ForReading, False)
            TMP = TE8.ReadLine 'solo dice "to Q"
            For J = 1 To 20
                TMP = TE8.ReadLine
                SP = Split(TMP, "=")
                ValoresATransformar1(J) = CLng(SP(1))
            Next J
            TMP = TE8.ReadLine 'solo dice "to S"
            For J = 1 To 20
                TMP = TE8.ReadLine
                SP = Split(TMP, "=")
                ValoresATransformar2(J) = CLng(SP(1))
            Next J
            TimeMaxSeparacion(1) = CLng(TE8.ReadLine)
            TimeMaxSeparacion(2) = CLng(TE8.ReadLine)
        TE8.Close
    End If
    
    CoinMuyJuntosAcum(1) = 0 'inicializa los valores
    CoinMuyJuntosAcum(2) = 0
    
End Sub

'Private Sub wLTE(n As Long)
'    frmIndex.picLTE.Cls
'    frmIndex.picLTE.Print CStr(n) + " " + CStr(Timer)
'End Sub

Public Sub VerSiTocaVMute()
    ' ver quiere videos continuos
    If PUBs.HabilitarPublicidadesVMute = False Then Exit Sub
    ' ... y si tiene videos
    If PUBs.TotalPUBsMUTE = 0 Then Exit Sub
    'ver si esta ocupada la salida de TV
    If EsVideo And Salida2 Then Exit Sub
    'ver si ya esta reproduciendo algo !!!
    If frmIndex.MP3.IsPlaying(3) Then Exit Sub
        
    '**************************************
    'se debe ejecutar un video mudo!!!
    '**************************************
        'que no se pase
    PUBs.UltimaReproducidaVMute = PUBs.UltimaReproducidaVMute + 1
    If PUBs.UltimaReproducidaVMute > PUBs.TotalPUBsMUTE Then
        PUBs.UltimaReproducidaVMute = 1
    End If
    
    Dim FJ As String
    FJ = PUBs.ArchsVMute(PUBs.UltimaReproducidaVMute)
    
    If fso.FileExists(FJ) = False Then Exit Sub
        
    ' Tocar el fichero
    On Local Error GoTo ErrEjecutarTema
    'SOLO EL 3 PARA vMUTE
    
    frmIndex.MP3.FileName(3) = FJ
    frmVIDEO.picBigImg.Visible = False
    frmIndex.MP3.DoOpenVideo "child", frmVIDEO.picVideo.HWND, 0, 0, _
        (frmVIDEO.picVideo.Width / 15), (frmVIDEO.picVideo.Height / 15), 3
    
    TotalTema(3) = frmIndex.MP3.LengthInSec(3)
    'UpdateHastaTema 3 'no hace falta parece
    
    frmIndex.picVideo(IAANext).Visible = False
    frmIndex.picKAR.Visible = False
    frmVIDEO.picKAR_V.Visible = False
    frmVIDEO.picVideo.Visible = True
    frmIndex.MP3.Volumen(3) = 0 ' ES MUDOOOO
    frmIndex.MP3.DoPlay 3
    
    Exit Sub
ErrEjecutarTema:
    tERR.AppendLog tERR.ErrToTXT(Err), "vMute.BAS" + ".acpo6"
    Resume Next
            
End Sub
    
Public Function GetPrecios(lFormat As Long, Separador As String) As String
    'lformat en cero es plata
    '1 = creditos
    
    'separador es por que el lblPrecios 2 es horizantal (sep = /) y el otro es vertical (sep = vbcrlf)
    
    Dim TMP As String
    TMP = ""
    
    If CreditosCuestaTema(0) = 0 And CreditosCuestaTema(1) = 0 And CreditosCuestaTema(2) = 0 Then
        TMP = TR.Trad("Musica Gratis%98%Se refiere a los precios. En este caso se puso el precio en cero%99%")
    End If
    
    If CreditosCuestaTema(0) > 0 Then
        Select Case lFormat
            Case 0
                TMP = TR.Trad("1 cancion%99%") + " = " + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTema(0), , , , vbFalse))
            Case 1
                TMP = TR.Trad("1 cancion%99%") + " = " + _
                    CStr(Round(CreditosCuestaTema(0))) + _
                    TR.Trad(" cred.%98%abreviatura de créditos%99%")
        End Select
    End If
    
    If CreditosCuestaTema(1) > 0 Then
        Select Case lFormat
            Case 0
                TMP = TMP + Separador + TR.Trad("2 canciones%99%") + " = " + _
                    CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTema(1), , , , vbFalse))
            Case 1
                TMP = TMP + Separador + TR.Trad("2 canciones%99%") + " = " + _
                    CStr(Round(CreditosCuestaTema(1), 2)) + _
                    TR.Trad(" cred.%98%abreviatura de créditos%99%")
        End Select
        
    End If
        
    If CreditosCuestaTema(2) > 0 Then
        Select Case lFormat
            Case 0
                TMP = TMP + Separador + _
                    TR.Trad("3 canciones%99%") + " = " + _
                    CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTema(2), , , , vbFalse))
            Case 1
                TMP = TMP + Separador + _
                    TR.Trad("3 canciones%99%") + " = " + _
                        CStr(Round(CreditosCuestaTema(2), 2)) + _
                        TR.Trad(" cred.%98%abreviatura de créditos%99%")
        End Select
    End If
    
    'si es gratis no usar!
    If CreditosCuestaTemaVIDEO(0) = 0 And CreditosCuestaTemaVIDEO(1) = 0 And CreditosCuestaTemaVIDEO(2) = 0 Then
        TMP = TMP + Separador + TR.Trad("Videos Gratis%99%")
    End If
    
    If CreditosCuestaTemaVIDEO(0) > 0 Then
        Select Case lFormat
            Case 0
                TMP = TMP + Separador + TR.Trad("1 video%99%") + " = " + _
                    CStr(FormatCurrency(CreditosCuestaTemaVIDEO(0) * (PrecioBase / TemasPorCredito), , , , vbFalse))
            Case 1
                TMP = TMP + Separador + TR.Trad("1 video%99%") + " = " + _
                CStr(Round(CreditosCuestaTemaVIDEO(0))) + _
                TR.Trad(" cred.%98%abreviatura de créditos%99%")
        End Select
    End If
        
    If CreditosCuestaTemaVIDEO(1) > 0 Then
        Select Case lFormat
            Case 0
                TMP = TMP + Separador + TR.Trad("2 videos%99%") + " = " + _
                    CStr(FormatCurrency(CreditosCuestaTemaVIDEO(1) * (PrecioBase / TemasPorCredito), , , , vbFalse))
            Case 1
                TMP = TMP + Separador + TR.Trad("2 videos%99%") + " = " + _
                    CStr(Round(CreditosCuestaTemaVIDEO(1))) + _
                    TR.Trad(" cred.%98%abreviatura de créditos%99%")
        End Select
    End If
        
    If CreditosCuestaTemaVIDEO(2) > 0 Then
        Select Case lFormat
            Case 0
                TMP = TMP + Separador + TR.Trad("3 videos%99%") + " = " + _
                    CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTemaVIDEO(2), , , , vbFalse))
            Case 1
                TMP = TMP + Separador + TR.Trad("3 videos%99%") + " = " + _
                    CStr(Round(CreditosCuestaTemaVIDEO(2))) + _
                    TR.Trad(" cred.%98%abreviatura de créditos%99%")
        End Select
    End If
    
    'UpManu
    'ver si esta habilitado cancion VIP
    If CreditosXaVipMusica > 0 Then
        TMP = TMP + Separador + getStrMusicaVIP
    End If
    
    GetPrecios = TMP
End Function

'UpManu
Public Function getStrMusicaVIP() As String
    Dim TMP As String: TMP = ""
    
    PrecNowVIP = CreditosXaVipMusica
    
    If CreditosXaVipMusica > 0 Then
        TMP = TR.Trad("Música VIP%98%Canciones que se ejecutan antes que todas (VIP Music esta ok)%99%") + " = "

        Select Case ShowCreditsMode
            Case 0
                TMP = TMP + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosXaVipMusica, , , , vbFalse))
                
            Case 1
                TMP = TMP + CStr(Round(PrecNowVIP)) + _
                    TR.Trad(" cred.%98%abreviatura de créditos%99%")
        End Select
    End If
    'devuelve "" si no esta habilitado
    getStrMusicaVIP = TMP
End Function

Public Sub UpdateHastaTema(I As Long)
    frmIndex.MP3.HastaTema(I) = TotalTema(I)
End Sub

Public Function YaCerrar3PM(Optional NoApagaaaar As Boolean = False, _
    Optional G12 As Boolean = False, Optional gExec As String = "") As Long
    
    'g12 carga todos los datos mestadisticos de la pc
    'gexec ejecuta algo antes de salir
    
    frmIndex.Timer3.Interval = 0
    frmIndex.Timer1.Interval = 0
    frmIndex.tbrPassImg1.Detener
    
    If TengoBluetooth Then
        downblUtu True
    End If
    
    If TengoUSB Then
        UB.Terminar 'si no se clava feo
        If BloquearTecladosUSB Then
            'mm92 reactivar teclados usb si corresponde
            usbKB.Activar_USB_KeyBoard
        End If
    End If
    
    '***********************************************
    tERR.Anotar "acdn0" 'no tocar cierra el tius!!! (YA NO LO CIERRA MAS)
    '***********************************************
    
    tERR.Anotar "acdn1"
    'SKy apagar pa salir de 3PM
    'apagar todos los indicadores!
    LedEvent "APAGAR"
    'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZAR_SIGUIENTE
    'que se come un tema de la lista
    MostrarCursor True
    frmIndex.MP3.DoClose 99 'AQUI SE DEBERIA ELIMIAR REGISTRO DE ERRORES
    tERR.Anotar "acdn2"
    frmIndex.VU.DoPause False
    frmIndex.VU.Terminar
    
    'desde el formulario que se llama se desacarga, en teoria solo falta el index
    Unload frmIndex
    
    'para el caso especial cuando estoy cargando la clave
    If NoApagaaaar = False Then
        If ApagarAlCierre Then APAGAR_PC
    End If
    
    'Unload frmIndex
    
    If G12 Then
        Dim d1 As Long
        d1 = GG12
    End If
    
    If ActivarERR Then
        tERR.StopGrabaTodo 'cierra y borra el archivo ya que se grabo OK
        'tambien el de MM, este se hace en el doClose 99
    End If
    
    YaCerrar3PM = 0
    
    'en realidad es para ejecutar algun ejecutable pero hago esta excepcion
    
    
    If gExec <> "" Then
        If gExec = "REINI" Then
            REINICIAR_PC
            GoTo FFiiNN
        End If
        
        DoEvents 'sin esto el proceso no es asincrono
        Dim DD As Double
        DD = Shell(gExec, vbNormalFocus)
        AppActivate DD 'activa una aplicacion de que comienza a recibir los eventos del teclado
    Else
FFiiNN:
        End
    End If
    
    
    
End Function

'pasar al ritmo que sigue ...
Public Sub goNextRitmo()
    Dim nxR As String
    nxR = GetStrigNextRitmo
    tERR.Anotar "acdo-3", nxR
    If nxR <> "" Then
        SelPagina nxR
        tERR.Anotar "acdo-4"
    End If
End Sub

'me dice cual es el ritmo que sigue despues del disco elegido
'sirve para poner boton de pasar de ritmos
'se debe usare combinado con la funcion "SelPagina"
Public Function GetStrigNextRitmo() As String

    Dim AA As Long
    'empiezo en ese y doy la vuelta al inicio si hace falta
    'buscar el primer numero de disco que cumpla con la condicion solicitada
    AA = nDiscoGral
    tERR.Anotar "acdo-5", AA
    Dim VueltasCompletas As Long
    VueltasCompletas = 0
    
    Dim RitmoActual As String
    RitmoActual = UCase(fso.GetBaseName(fso.GetParentFolderName(txtInLista(MATRIZ_DISCOS(AA), 0, ","))))
    tERR.Anotar "acdo-6", RitmoActual
    Dim lastRitmoVisto As String
    Do
        lastRitmoVisto = UCase(fso.GetBaseName(fso.GetParentFolderName(txtInLista(MATRIZ_DISCOS(AA), 0, ","))))
        tERR.Anotar "acdo-7", lastRitmoVisto, RitmoActual, MATRIZ_DISCOS(AA)
        'ver si ya pase a otro
        If lastRitmoVisto <> RitmoActual Then
            'ya llegue a otro ritmo
            GetStrigNextRitmo = lastRitmoVisto
            tERR.Anotar "acdo-8"
            Exit Function
        End If

        'pasar al disco que sigue y si termina ir al inicio
        AA = AA + 1
        tERR.Anotar "acdo-9", CStr(AA) + "/" + CStr(UBound(MATRIZ_DISCOS))
        If AA > UBound(MATRIZ_DISCOS) Then
            AA = 1 'empieza desde el pricipio de nuevo para ver si es necesario
            
            VueltasCompletas = VueltasCompletas + 1
            If VueltasCompletas = 2 Then
                tERR.AppendLog "acba22:" + CStr(AA) + ":" + RitmoActual + ":" + lastRitmoVisto + ":" + CStr(UBound(MATRIZ_DISCOS))
                GetStrigNextRitmo = ""
                Exit Function
            End If
        End If
    Loop
    
    Exit Function

End Function

Public Sub SelPagina(RitmoSel As String, Optional PrimeraLetra As String = "A")

    Dim FolRit As String
    Dim FolSel As String

    Dim AA As Long
    'empiezo en ese y doy la vuelta al inicio si hace falta
    'buscar el primer numero de disco que cumpla con la condicion solicitada
    AA = 0 'nDiscoGral
    Dim VueltasCompletas As Long
    VueltasCompletas = 0
    Do
        FolSel = UCase(fso.GetBaseName(fso.GetParentFolderName(txtInLista(MATRIZ_DISCOS(AA), 0, ","))))
        FolRit = UCase(RitmoSel)
        
        If FolSel = FolRit Then
            Exit Do 'queda en aa el numero que me interesa
        End If

        'pasar al disco que sigue y si termina ir al inicio
        AA = AA + 1
        If AA > UBound(MATRIZ_DISCOS) Then
            AA = 0
            VueltasCompletas = VueltasCompletas + 1
            If VueltasCompletas = 2 Then GoTo NoEnuentro 'si dio dos vueltas no lo encontro!
        End If
    Loop
    
    'ya encontre el ritmo, ahora busco el disco
    'su numero esta en AA!
    Dim nD1 As Long, nD2 As Long
    
    Do
        'no solo la misma letra si no cualquiera mayor!
        nD1 = Asc(UCase(Left(txtInLista(MATRIZ_DISCOS(AA), 1, ","), 1)))
        nD2 = Asc(UCase(PrimeraLetra))
        If nD1 >= nD2 Then
            Exit Do 'ya encontre el disco a mostrar es AA!!!
        End If
        
        AA = AA + 1
        If AA > UBound(MATRIZ_DISCOS) Then
            GoTo NoEnuentro 'nunca un origen da la vuelta!. Si llego aca no se que hacer
        End If
        
    Loop
    
    'ya tengo en AA el numero de disco que hay que elegir
    'ahora ver cual es el que debe ser primero de su página
    Dim pagToSel As Long
    pagToSel = AA \ (TapasMostradasH * TapasMostradasV)
    
    Dim PrimeroDeLaPaginaQueNecesito As Long
    PrimeroDeLaPaginaQueNecesito = (pagToSel) * (TapasMostradasH * TapasMostradasV)
    
    If nDiscoSEL <> 0 Then frmIndex.UnSelDisco nDiscoSEL
    tERR.Anotar "acba2", nDiscoSEL
    DiscosEnPagina = frmIndex.CargarDiscos(PrimeroDeLaPaginaQueNecesito, False, 0, AA - PrimeroDeLaPaginaQueNecesito)
    
    Exit Sub
    
NoEnuentro:
    tERR.AppendLog "acba21:" + CStr(AA), RitmoSel
    
    'no se donde esta!
    'no hago nada
End Sub


'roberto cpaz   RCP888
'diego antonio sanchez corr DASC717771090
'Edgar Giovanni Valdez Hernandez GUAT HGVHG34771000
'rio ceballos rioceballos88
'Clifton Forde PANAMA CFP7118820192
'Melina Gieco StaFe MGSF711905621
'Caludia Sala BsAs MRCSR81172660
'Andres Giamello AGBA718829540/AG31
'Alejandro Maltez NICARAGUA AMN5102991732
'Abraham Grenberg Valle Verde SA GUAT AGVVSA8177109
'humberto segundo cruces CHI HSCC719288012
'juan miguel RepDominic JMRD611885094
'Pablo Duvos UY PDUY210098881
'FRANCISCO JAVIER GONZALEZ LAZCANO CHILE FJGLC71625551
'Oscar Figeuroa Martinez Salvador OEFMES810001
'Hector Amigo BsAs HABSAS8281901
'Jesus Alexander SALV JAS7166290011
'william Obando y Francisco Vielman Flores GUAT WOFVFGU918812
'Damian Ostuni BsAs DOBA811726300
'carlos salas JJY CSJJYAR719922
'daniel omar herrera robles chile DOHRCH6199201
'paulo garcia CHILE PGC5220119851
'Jorge Horta CHILE JOCH102881276
'Alejandro Carmona Oliveros Chile ACOCH3217729
'edison ariel caceres chile EACCH81032772
'Daniel Martinez Chicago DMCE183745510
'Wilmer Fidel Marquez Silva WFMSPR2981109
'julio papetti TUCUMAN JPT1077594731
'jorge albin JACP719283001 jorge albin FEDERICODANIEL
'jorge albin RCP888
'Dardo maidana DARDOMAIDANA
'german peier BsAs GPBSAS7812003
'luis iglesias BsAs LIBA6152896R
'eduardo alberti UY EAJCM2987889h
'juan francisco gonzalez COL JFG729432119q
'Mauricio Levuy Sergio Davo MEX MLSD61846362e
'Jesus Andrès Mata Jimenez MEX JAMG67298187r
'Thomas Hernadez MEX THH635478111g
'rolando torres honduras RTMH523142567z
'david gonzalez MEX DGM652253435y
'Jose Juan Martinez Arguello JIMM MEX JJMA81948572y
'Giovanne Barrios MEX GB6156901836y
'tommy corrientes TC194736251438y
'Luis Enrique Ruiz Chaparro MEX LERC8711101yy
'onofre inda OI71909081125y
'efrain solarte COL EFS50091hgyurr
'henry soto ECU HS611ecu119yh
'Alex Herrera COL AHQ54COL52hyy
'VICTOR HUGO DE LA ROSA (de JMFC) vhdlr5001787y"
'Tomas Nuñez Gonzalez    sncMEX098181y
'Mig    uel Angel Santos Hernandez MEX MASH81090011y
'JUAN MARTIN FLORES CRUZ MEX JMFCm6511yyyq
'Mauro Villaroel     MV541CHQ9090Y
'Chirstian Beltra    cb9811191ujY
'Miguel Angel Cozzi  grAS981aATTy6
'Marcos Sepulveda    bsaHH0981AWqQ
'Rigoberto Matamoros - Oscar Otero Cartagena (El Salvador)   fRF4247L000wZ
'Jose Luis Dorado    33Ccq0151AxqF
'Ramon Daniel Cruz   RMLVF00012yqq
'Miguel Angel Cozzi  grAS981aATTy6
'Ivan Vera   LOpaFE1701666
'Carlos Alberto Montaña Alvarado Caa9107g8s811
'Gabriel Pablo de Rosa   AFD076qwnn100
'Santiago Vignolo    LIQ3661SV0909
'Humberto Breton BR7ME2jGtt981
'Juan Carlos Monsegu MONS7111yHu66
'Jorge Andres Gonzales Torres    JAGT61098Saa6
'Hugo Kollman    KOLL717109888
'Eduardo Rodriguez   ERO77701192FF / MARC777
'Victor Rocha    VR541SLP11MEX
'Roberto Hurtado RHUR28177MEXy
'Alberto Devit   AlDe1098MXca5
'German Becley   GerBKL00198AA
'Abelardo Garcia Morales ABG011boCO1ky
'Judith Rodriguez    ROD0906mx143u
'Guillermo Milian    gMIL991Mex199
'Jesus Andres Mata Jimenez MG611mex0909a
'Juan Serano JaS0106uuw103
'Sergio Sosa Mendoza SeSo711922yh6
' Leandro Visciarelli CBA7111levi09
'Eduardo Rodriguez Uruguay ERO77701192FF
'Favio Martinez Gomez COL fa61MG52COL91
'Oscar Armenta Soberanis OAS81090Mx880
'(RIGMAT)Francisco Somoza  FrHN0102099yi
'Ernesto isidro vazquez MEX eiv767611iJAA
'Cesar Gordillo GERSA de CV MEX cg5510978AByR
'Rene escrich SALV REES91210909u
'Allan orante Martinez MEX AOM519090hnYa
'guille2p españa G2Pk9111900ES
' Miguel Angel Santos hernandez dice que le di yo??? ms6511comp9ME
'ES EL MISMO DE MSCOMPU GARKA!!!!!!!!!!!!!!!!
'enrique israel mora suarez EIMS611609yyw
'desiderio meneseres BOL Des911BOL1011

Public Function RavI() As Long

    tERR.Anotar "acfh", VALIDAR
    
    
    If VALIDAR Then
        'ver cual es el máximo y si hay que avisar
        tERR.Anotar "acfj", CreditosValidar, ValidarCada, AvisarAntes
        Dim QuedanC As Long
        QuedanC = ValidarCada - CreditosValidar
        
        If (AvisarAntes > QuedanC) Then
                
            'que solo pida una vez la clave por cada sesión
            If QuedanC > 0 Then
                'esta perdonado, solo una vez por sesion
                If YaPediCL Then
                    RavI = 3
                    Exit Function
                Else
                    RavI = 2
                End If
            Else 'ya esta pasado
                If YaPediCL Then
                    'ya esta avisado que no joda!!!"
                    RavI = 5
                Else
                    RavI = 4
                End If
            End If
        
            'solicitar una clave
            'se podra saltear solo si todavia no llego al limite
            'uso el frmClave que tiene la variable publica ClaveIngresada
            Dim ClaveCorrespondiente As String
            ClaveCorrespondiente = NumToTec(ClaveParaValidar(CodigoParaClaveActual))
            
            tERR.Anotar "acfl"
            YaPediCL = True
            frmCLAVE.Show 1
            
            'que no tome el enter de frmclave.show en el index!
            frmIndex.OkInState1 = 0
            
            tERR.Anotar "acfm", UCase(ClaveIngresada), UCase(ClaveCorrespondiente)
            
            'si pone la clave de administrador tambien vale
            If UCase(ClaveIngresada) = UCase(ClaveAdmin) Then GoTo mOK
            If UCase(ClaveIngresada) = "RMLVF" Then GoTo mOK
            
            If TexToTec(UCase(ClaveIngresada)) <> UCase(ClaveCorrespondiente) Then
                If QuedanC > 0 Then
                    
                Else
                    If K.sabseee(dcr("q44KmdDBQ+IB8dTOX8F+VA==")) <= CGratuita Then
                        MsgBox TR.Trad("Si hubiera una licencia cargada " + _
                        "esta máquina estaría bloqueada!!!" + vbCrLf + _
                        "MAS CUIDADO LA PROXIMA VEZ%98%Se refiere a bloqueo " + _
                        "de seguridad configurado para que los que trabajan " + _
                        "los equipos no se los roben a los dueños%99%")
                    Else 'solo lo mato si no es ua PC de prueba
                        'MsgBox "No podrá seguir utilizando 3PM hasta que valide con la clave correspondiente"
                        'otra forma de bloqueo que no sea salir directo que es muy detectable
                        'End
                        'no digamos que bruto que loco pero no es un simple END
                        
                        'otra opcion es
                        'YaCerrar3PM
                        
                        'otra mas fea es
'                        Randomize
'                        Dim J As Long
'                        J = Int(Rnd * 30)
'                        ReDim Preserve MATRIZ_DISCOS(J)
                        
                        On Local Error Resume Next
                        Dim JK As Long
                        For JK = 0 To 30
                            frmIndex.TapaCD(JK).Top = -frmIndex.TapaCD(JK).Height * 2
                        Next JK
                    End If
                End If
            Else
mOK:
                RavI = 1
                tERR.Anotar "acfn"
                'todo OK. Cargo bien la clave
                CreditosValidar = 0
                EscribirArch1Linea GPF("radliv"), "0"
                'empezar un nuevo periodo
                CrearNuevoCodigoValidar 'graba el archivo con un numero al azar
            End If
        Else 'pidieron validar pero todavia falta
            RavI = 1
        End If
        tERR.Anotar "acfo", ValidarCada, CodigoParaClaveActual
    Else
        RavI = 0 'nadie pidio validar
    End If
    
End Function

Public Function GetTag(TotalTag As String, par As String, _
    Optional SEP1 As String = "|", Optional SEP2 As String = ":") As String
    'lee un tag que sean datos separados por "|" y como PARAMETRO:VALOR dentro de cada elemento
    'sep1 es "|" y sep2 es ":"
    
    On Local Error GoTo ErrTT
    
    Dim SP() As String
    SP = Split(TotalTag, "|")
    Dim C As Long, SP2() As String
    For C = 0 To UBound(SP)
        SP2 = Split(SP(C), ":")
        Dim D As Long
        For D = 0 To UBound(SP2) Step 2
            If LCase(SP2(D)) = LCase(par) Then
                GetTag = SP2(D + 1)
                Exit Function
            End If
        Next D
    Next C
    
    Exit Function
    
ErrTT:
    GetTag = "ERR"
End Function

Public Function ForceFocus(HW As Long) As Boolean
    If SetForegroundWindow(HW) = 0 Then
        ForceFocus = False
    Else
        ForceFocus = True
    End If
End Function

Public Sub SelBT(o As Object, Sel As Boolean)
    If Sel Then
        o.BackColor = ColSel
        o.BackColor2 = Col2Sel
    Else
        o.BackColor = ColUnSel
        o.BackColor2 = Col2UnSel
    End If
    o.Font.Bold = Sel
End Sub

Public Function getWAN(HD As Long)  'para avisar donde hablarme
    Dim T(6) As String
    T(0) = STRceros(Day(Date), 2)
    T(1) = STRceros(Month(Date), 2)
    T(2) = STRceros(Year(Date), 4)
    T(3) = STRceros(HD, 12)
    T(4) = STRceros(CLng(Timer), 12)
    T(5) = T(0) + T(1) + T(2) + T(3) + T(4)
    
    T(6) = Mid(T(5), 1, 1) + Mid(T(5), 7, 1) + Mid(T(5), 13, 1) + Mid(T(5), 19, 1) + Mid(T(5), 25, 1) + _
           Mid(T(5), 2, 1) + Mid(T(5), 8, 1) + Mid(T(5), 14, 1) + Mid(T(5), 20, 1) + Mid(T(5), 26, 1) + _
           Mid(T(5), 3, 1) + Mid(T(5), 9, 1) + Mid(T(5), 15, 1) + Mid(T(5), 21, 1) + Mid(T(5), 27, 1) + _
           Mid(T(5), 4, 1) + Mid(T(5), 10, 1) + Mid(T(5), 16, 1) + Mid(T(5), 22, 1) + Mid(T(5), 28, 1) + _
           Mid(T(5), 5, 1) + Mid(T(5), 11, 1) + Mid(T(5), 17, 1) + Mid(T(5), 23, 1) + Mid(T(5), 29, 1) + _
           Mid(T(5), 6, 1) + Mid(T(5), 12, 1) + Mid(T(5), 18, 1) + Mid(T(5), 24, 1) + Mid(T(5), 30, 1) + _
           Mid(T(5), 31, 1) + Mid(T(5), 32, 1)

    getWAN = T(6)
    'MsgBox HD
    
End Function

Public Function cmbqqq(Texto As String, ql As String, japi As Boolean) As String
    'encriptar (ql es clave y japi es invertido)
    'Cargo los datos
    If Texto = "" Then
        pinchilon = ""
        Exit Function
    End If
    
    Dim F As Integer
    
    Dim Buffer() As Byte
    'Buffer = Texto 'se meten de dos en dos las letras ??? sera por algo de ascii vs unicode
    
    ReDim Buffer(Len(Texto) - 1)
    For F = 1 To Len(Texto)
        Buffer(F - 1) = Asc(Mid(Texto, F, 1))
    Next F

    Dim xql() As Byte
    'xql = ql
    
    ReDim xql(Len(ql) - 1)
    For F = 1 To Len(ql)
        xql(F - 1) = Asc(Mid(ql, F, 1))
    Next F
    
    'Encripto
    
    Dim Char1 As Integer 'Caracter Original
    Dim Char2 As Integer 'Caracter ya Modificado (char1+char3) o (char1-char3)
    Dim Char3 As Integer 'Caracter de la Clave
    
    'Voy dando vueltas por la clave asi que necesito un indice
    Dim Contadorql As Integer 'Indice de la clave
    Contadorql = 0
    
    Dim I As Long
    Dim NuevoDato() As Byte
    
    ReDim NuevoDato(Len(Texto) - 1)
    
    For I = 0 To UBound(Buffer)
        Char1 = Buffer(I)
        Char3 = xql(Contadorql)
        If japi = True Then
            Char2 = Char1 - Char3
        Else
            Char2 = Char1 + Char3
        End If

        If Char2 < 0 Then
            Char2 = 256 + Char2
        End If
    
        If Char2 > 255 Then
            Char2 = Char2 Mod 256
        End If
    
        NuevoDato(I) = Char2
        
        Contadorql = Contadorql + 1
        If Contadorql > UBound(xql) Then Contadorql = 0
    Next I
    
    Dim tRES As String
    For F = 0 To UBound(NuevoDato)
        tRES = tRES + Chr(NuevoDato(F))
    Next F
    
'    Dim Ver As String
'    For F = 0 To UBound(Buffer)
'        Ver = Ver + Chr(Buffer(F)) + " - " + Chr(NuevoDato(F)) + " * "
'    Next F
'    MsgBox Ver
    
    cmbqqq = tRES
    
End Function

Public Function dwqu(T, tt, ttt) As Long
    On Local Error GoTo errdwqu
    'graba registro de canciones cobradas por el equipo
    
    'puede ser
    '"E" + TEMA + "*" + Precio, dwQU_See, DTaa 'escucha cancion
    '"B" + tema + "*" + Precio, dwQU_See, DTaa 'vendio por bluetooth
    '"U" + tema + "*" + Precio, dwQU_See, DTaa 'vendio por usb
    'AGREGAR AL REGISTRO QUE LEE !!!!!!!!!!!!!!!
    '"C" + tema + "*" + Precio, dwQU_See, DTaa 'vendio en cd
    'AGREGAR AL REGISTRO QUE LEE !!!!!!!!!!!!!!!
    'poner fecha y hora para hacer estadisticas mejores !!!!!!
    
    Dim TEd1 As TextStream, TEd2 As TextStream, TEd3 As TextStream, TEd4 As TextStream, TEd5 As TextStream
    Dim FFF As String
    Dim ID1ttt As String, ID2ttt As String 'identificador del renglon
    Randomize: ID1ttt = STRceros(CLng(Rnd * 1000000), 7)
    Randomize: ID2ttt = STRceros(CLng(Rnd * 1000000), 7)
    
    tERR.Anotar "gaaa", ID1ttt, ID2ttt
    FFF = cmbqqq(ID1ttt + T + ID2ttt, "Ingrese su pais de residencia", False)
    Set TEd1 = fso.OpenTextFile(GPF("acumsg0"), ForAppending, True)
        TEd1.Write FFF + Chr(5) + Chr(7) + Chr(6) + Chr(4)
        'el separador deber largo para que no
        'se genere un separador cuando se encripta sin querer
    TEd1.Close
    
    tERR.Anotar "gaab"
    FFF = cmbqqq(ID2ttt + tt + ID1ttt, "Telefono o fax", False)
    Set TEd2 = fso.OpenTextFile(GPF("acumsg1"), ForAppending, True)
        TEd2.Write FFF + Chr(5) + Chr(6) + Chr(6) + Chr(5)
    TEd2.Close
    
    tERR.Anotar "gaac"
    FFF = cmbqqq(ID1ttt + ID2ttt + ttt, "Email tecnico", False)
    Set TEd3 = fso.OpenTextFile(GPF("acumsg2"), ForAppending, True)
        TEd3.Write FFF + Chr(7) + Chr(7) + Chr(6) + Chr(5)
    TEd3.Close
    
    tERR.Anotar "gaad"
    FFF = cmbqqq(T + ID1ttt + tt + ID2ttt + ttt, "Email administrativo", False)
    Set TEd4 = fso.OpenTextFile(GPF("acumsg3"), ForAppending, True)
        TEd4.Write FFF + Chr(4) + Chr(7) + Chr(6) + Chr(5)
    TEd4.Close
    
    tERR.Anotar "gaae"
    FFF = cmbqqq(ID1ttt + ID2ttt, "Gracias por confiar en tbrSoft", False)
    Set TEd5 = fso.OpenTextFile(GPF("acumsg4"), ForAppending, True)
        TEd5.Write FFF + Chr(4) + Chr(6) + Chr(6) + Chr(4)
    TEd5.Close
    
    Exit Function
    
errdwqu:
    tERR.AppendLog "dwquhh", tERR.ErrToTXT(Err)
    Resume Next
End Function

Public Function GG12() As Long
    Dim JS2 As New tbrJUSE.clsJUSE
    Dim F As String, Dt As String
    'Dt = CStr(Year(Date)) + STRceros(Month(Date), 2) + STRceros(Day(Date), 2) + STRceros(Hour(time), 2) + STRceros(Minute(time), 2)
    F = AP + "copyleft.Js"
    
    If fso.FileExists(F) Then fso.DeleteFile F, True
    
    JS2.Archivo = F
    
    JS2.AddFile GPF("acumsg0")
    JS2.AddFile GPF("acumsg1")
    JS2.AddFile GPF("acumsg2")
    JS2.AddFile GPF("acumsg3")
    JS2.AddFile GPF("acumsg4")
    
    CreateMyFile AP + "my.log", Get_LL
    
    'REGISTRO BASICO + REGISTRO DE MMPLAYER
    AddFiles App.path, "log", JS2
    
    'ARCHIVOS W15
    AddFiles App.path, "w15", JS2
    
    'CONFIGURACION DE 3PM
    JS2.AddFile AP + "sf\marad.ona"
    
    'OTRAS COSAS INTERESANTES
    JS2.AddFile GPF("origs") 'lista de origenes de discos 'EX: sf+ "oddtb.jut"
    JS2.AddFile GPF("cd3pm") 'Copia clave sf + "c2LK.dll"
    JS2.AddFile GPF("cccd3pm") 'Copia clave sf + "c2LK.dll"
    JS2.AddFile GPF("cd4pm") 'Archivo de licencia 3pm 7.0 (GENERADO)
    JS2.AddFile GPF("cd7pm") 'Archivo RECIBIDO de licencia 3pm 7.0 COREGIDO Y EN USO
    JS2.AddFile GPF("rdcday") 'registro diario del contador sf + "daily.cfg"
    JS2.AddFile GPF("dalivmp2") 'archivo con las claves para validar

    Dim res As Long
    res = JS2.Unir
    
    'ahora lo encripto y elimino el comun
    Dim TCE As New tbrCrypto.Crypt
    
    TCE.EncryptFile eMC_Blowfish, F, F + "B", "guarana" + fso.GetBaseName(F) + "fresco"
    
    'borrar el original
    fso.DeleteFile F
    
    'limpiar los contadores
    On Local Error Resume Next
    ClearTextFile GPF("acumsg0")
    ClearTextFile GPF("acumsg1")
    ClearTextFile GPF("acumsg2")
    ClearTextFile GPF("acumsg3")
    ClearTextFile GPF("acumsg4")
    
    DeleteFiles AP, "log"
    DeleteFiles AP, "w15"
    
    GG12 = 0
End Function

Private Function DeleteFiles(sFolder As String, Extension As String) As Long
    'devuleve la cantidad de agregados
    Dim F As Scripting.folder
    Set F = fso.GetFolder(sFolder)
    Dim F2 As Scripting.File
    For Each F2 In F.Files
        If LCase(Right(F2.Name, Len(Extension))) = LCase(Extension) Then
            F2.Delete True
        End If
    Next
End Function

Private Function ClearTextFile(AR As String)
    Dim TW As TextStream
    Set TW = fso.CreateTextFile(AR, True)
    
    TW.Close
End Function

Private Function AddFiles(sFolder As String, Extension As String, Jss) As Long
    'devuleve la cantidad de agregados
    Dim FL As Scripting.folder
    Set FL = fso.GetFolder(sFolder)
    Dim F2 As Scripting.File
    Dim E1 As String, E2 As String
    For Each F2 In FL.Files
        E1 = LCase(Right(F2.Name, Len(Extension)))
        E2 = LCase(Extension)
        If E1 = E2 Then
            Jss.AddFile F2.path
        End If
    Next
End Function

Public Function ucdate(lt As String)

    'segun que tenga el pendrive puede tener diferentes funciones administrativas
    Dim G As String
    
    '*********************ute.exe, recopila info de la recaudacion*********************************
    'hacer varias verificaciones antes de ejecutar el actualizador
    G = lt + ":\ute.exe" 'ute es el programa que recoge información del equipo, todavia no funciona
    If fso.FileExists(G) Then
        'asegurarme que funcione con validador
        If VALIDAR Then
            CreditosValidar = 2000 'le pongo para que tenga un rato mas
            EscribirArch1Linea GPF("radliv"), CStr(CreditosValidar)
        End If
        
        Dim A As Long
        A = YaCerrar3PM(True, True, lt + ":\ute.exe")
        
        Exit Function
    End If
    
    
    '*********************agregar musica sin abrir cfg*********************************
    'en la raiz tiene que estar el archivo update-music.txt
    'además una carpeta llamada "up", en esta carpeta debe haber origenes de discos
    'mm889
    'debe tener la clave de administrador para hacerlo
    G = lt + ":\" + KeyUpdateMusic 'hay una clave que es el nombre del archivo
    If fso.FileExists(G) Then
    
        On Local Error GoTo KUY
    
        verPG "Buscando nuevo contenido multimedia ..."
    
        'buscar en el pendrive la musica nueva
        Dim Origenes As String
        Origenes = LeerArch1Linea(GPF("origs"))
        Dim PartS33() As String
        PartS33 = Split(Origenes, "*")
    
        Dim H As Long, U As Long
        For H = 0 To UBound(PartS33)
            tERR.Anotar "acfc3g", PartS33(H)
            'copiar todo loo que hay a la carpeta XXXX deberia revisar cada carpeta y ver que tenga contenido multimedia!!!
            Dim F01 As String 'carpeta de origen
            F01 = lt + ":\musica\" + fso.GetBaseName(PartS33(H))
            Dim F02 As String 'carpeta de destino
            F02 = fso.GetParentFolderName(PartS33(H)) + "\"
            
            verPG "Buscando " + vbCrLf + F01
            
            If fso.FolderExists(F01) Then
            
                'ver que haya lugar en el disco duro!!!
                Dim MbTot As Long 'mb totral en la unidad
                Dim MbLibre As Long 'libre en la unidad
                Dim PorcFree As Single 'porcetaje libre en la unidad
                InfoDisco2 Left(F02, 1), MbTot, MbLibre, PorcFree
                
                If (PorcFree < 10) Then
                    verPG "Hay menos del 10% libre en la unidad '" + Left(F02, 1) + ":\' , no se copiara !!", 5
                    verPG ""
                    Exit Function
                End If
                
                tERR.Anotar "acfc3h", F01, F02
                verPG "COPIANDO " + vbCrLf + F01
                
                fso.CopyFolder F01, F02, True
            End If
        Next H
        
        verPG "Finalizado correctamente, puede quitar dispositivo USB" + vbCrLf + _
            "Se reiniciara el sistema ...", 6
        verPG ""
        
        'YaCerrar3PM True, , "REINI"
    End If
    
    Exit Function
    
KUY:
    tERR.AppendLog "kuy66", tERR.ErrToTXT(Err)
    Resume Next
    
End Function

Public Sub loadOp_PG(op() As String)
    'cargar una lista de opciones con el automático del pendrive
    unLoadOp_PG
    
    Dim A As Long
    For A = 1 To UBound(op) + 1
        Load frmIndex.OP1(A)
        frmIndex.OP1(A).Caption = op(A - 1)
    Next A
    
    'ubicarlos
    Dim HEY As Long 'alto de todos juntos
    HEY = frmIndex.OP1(1).Height * frmIndex.OP1.Count
    
    frmIndex.OP1(1).Top = frmIndex.picPG.Height / 2 - HEY / 2
    
    For A = 2 To frmIndex.OP1.Count
        frmIndex.OP1(A).Top = frmIndex.OP1(A - 1).Top + frmIndex.OP1(A - 1).Height
    Next A
    
    'NOTERMINADO
End Sub

Public Sub unLoadOp_PG()
    On Local Error Resume Next
    Dim A As Long
    For A = 1 To frmIndex.OP1.Count
        Unload frmIndex.OP1(A)
    Next A
    
    frmIndex.OP1(0).Visible = False
    frmIndex.OP1(0).AutoSize = True
End Sub

Public Sub verPG(TXT As String, Optional PGwait As Long = -1)
    If TXT = "" Then
        frmIndex.picPG.Visible = False
    Else
        If frmIndex.picPG.Visible = False Then
            
            frmIndex.picPG.Width = frmIndex.Width * 0.7
            frmIndex.picPG.Height = frmIndex.Width * 0.7
            
            frmIndex.picPG.Left = frmIndex.Width / 2 - frmIndex.picPG.Width / 2
            frmIndex.picPG.Top = frmIndex.Height / 2 - frmIndex.picPG.Height / 2
            
            frmIndex.picPG.Visible = True
            frmIndex.picPG.ZOrder
        End If
        frmIndex.lblPG.Caption = TXT
        frmIndex.lblPG.Refresh
        
        If PGwait > 0 Then
            Dim T As Single
            T = Timer
            Do While Timer < T + PGwait
                DoEvents
            Loop
        
        End If
        
    End If
End Sub
    

Private Function Get_LL() As String
    'ver las versiones de todos las dlls
    Dim FLL As String, VLL As String
    Dim ACUM_LL As String
    
    FLL = SYSfolder + "tbrerr.dll":            VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrreg.dll":            VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrtimer.dll":          VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrfocus.dll":          VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrplayer02.dll":       VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrSoftVumetro.dll":    VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrListaRep.dll":       VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrSKS3.dll":           VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrjuse.dll":           VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrnfo.dll":            VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrFullPak.dll":        VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "caescrypto.dll":        VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrcaescrypto.dll":     VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrprogress.dll":       VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrFaroButton.ocx":     VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrEncr.dll":           VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrPaths.dll":          VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrDrives.dll":         VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrFrame.ocx":          VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrGraficos.dll":       VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrALotOfPictures.dll": VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "ijl11.dll":             VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SYSfolder + "tbrJPG.ocx":            VLL = fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    
    Get_LL = ACUM_LL
End Function

Private Sub CreateMyFile(pt As String, TX As String)
    Dim TE As TextStream
    Set TE = fso.CreateTextFile(pt, True)
        TE.Write TX
    TE.Close
End Sub

Public Sub esperar(n As Single)
    n = Timer + n
    Do While Timer < n
        DoEvents
    Loop
End Sub

'cada vez que pasa alguno de los eventos se avisa aqui
Public Sub LedEvent(stringEv As String)

    'en caso de que no haya teclado conectado incluso clientes nuestros con H2K puede ser que se clave un poco
    'la pc al tratar de ejecutar esta funcion
    If ActionLedOn = 0 Then Exit Sub
    
    'ver si estoy entre las horas permitidas
    Dim hourNow As Long
    hourNow = Hour(time)
    
    If (hourNow < ActionLedINIhs) Or (hourNow > ActionLedFINhs) Then
        Exit Sub
    End If

    Dim accionToDo As Long 'accion a hacer segum lo que paso

    Select Case stringEv
        Case "ActionLedMuchoCredito": accionToDo = ActionLedMuchoCredito
        Case "ActionLedPocoCredito":  accionToDo = ActionLedPocoCredito
        Case "ActionLedPalying":      accionToDo = ActionLedPalying
        Case "ActionLedNoPlaying":    accionToDo = ActionLedNoPlaying
        Case "ActionLedPalyingVip":   accionToDo = ActionLedPalyingVip
        Case "ActionLedNoPlayVip":    accionToDo = ActionLedNoPlayVip
        Case "APAGAR" 'apagar todo!
            SetKeyState vbKeyNumlock, False
            SetKeyState vbKeyCapital, False
            SetKeyState vbKeyScrollLock, False
            Exit Sub
    End Select
    
    
    ''//////////Acciones////////////////////////////////////////
    '///////////////////////////////////////////////////////////
    'cmbAction(0).AddItem "No hacer nada"                     '0
    'cmbAction(0).AddItem "Encender 'NUM LOCK'"               '1
    'cmbAction(0).AddItem "Apagar 'NUM LOCK'"                 '2
    'cmbAction(0).AddItem "Encender 'CAPS LOCK'"              '3
    'cmbAction(0).AddItem "Apagar 'CAPS LOCK'"                '4
    'cmbAction(0).AddItem "Encender 'SCROLL LOCK'"            '5
    'cmbAction(0).AddItem "Apagar 'SCROLL LOCK'"              '6
    '///////////////////////////////////////////////////////////
    'vbKeyCapital
    'vbKeyScrollLock
    
    Select Case accionToDo
        Case 0: 'no hacer nada
        Case 1: SetKeyState vbKeyNumlock, True
        Case 2: SetKeyState vbKeyNumlock, False
        Case 3: SetKeyState vbKeyCapital, True
        Case 4: SetKeyState vbKeyCapital, False
        Case 5: SetKeyState vbKeyScrollLock, True
        Case 6: SetKeyState vbKeyScrollLock, False
    End Select
    
    
End Sub

'mm91
'indica si una carpeta es un origen activo de 3pm
Public Function isOrigen(fldTest As String) As Boolean
    Dim AAA As Long
    isOrigen = False
    For AAA = 0 To UBound(PartOrigenes)
        If LCase((PartOrigenes(AAA))) = LCase(fldTest) Then
            isOrigen = True
            Exit For
        End If
    Next AAA
End Function

'devuelve el tamaño adaptado segun el tamaño a Bytes, KB, MB, GB, TB, etc
'devuelve por ejemplo 345 KB o 12.65 MB o 1.09 GB
Public Function tbrFileLen(sPath As String) As String
    If fso.FileExists(sPath) Then
        Dim FL As Currency
        
        FL = fso.GetFile(sPath).Size   'DARA NEGATIVO SI ES MAS DE 2 GB
        'FL = FileLen(sPath) 'DARA NEGATIVO SI ES MAS DE 2 GB
        'If FL < 0 Then FL = 2147483648# + 2147483648# + FL
        
        'segun el tamaño del archivo lo muestro
        Dim TipoB As String
        If FL < 1024 Then 'lo muestro en Bytes queda como esta
            TipoB = "bytes"
        End If
        
        If FL >= 1024 And FL < 1048576 Then 'lo muestro en Bytes queda como esta
            FL = FL / 1024
            TipoB = "KB"
        End If
        
        If FL >= 1048576 And FL < 1073741824 Then 'lo muestro en Bytes queda como esta
            FL = FL / 1048576
            TipoB = "MB"
        End If
        
        If FL >= 1073741824 Then 'And FL < 1099511627776# Then 'lo muestro en Bytes queda como esta
            FL = FL / 1073741824
            TipoB = "GB"
        End If
        
        tbrFileLen = CStr(Round(FL, 2)) + " " + TipoB
        
    Else
        tbrFileLen = -1
    End If
End Function

'desencriptar cadenas para o se vean al decompilar ele ejecutable
Public Function dcr(sT As String) As String
    Dim D As New tbrCrypto.Crypt
    Dim d2 As String
    Dim hh As Long
    hh = 10 * 10 * 10
    d2 = D.DecryptString(eMC_Blowfish, sT, "cuh" + CStr(hh) + "v", True) 'cuh1000v
    dcr = d2
End Function

Public Sub radlimas()
    CreditosValidar = CreditosValidar - SumValidar
    EscribirArch1Linea GPF("radliv"), CStr(CreditosValidar)
End Sub

Public Sub srtRNK()
    On Local Error GoTo errSRT
    tERR.Anotar "000A-00901"
    'ver si existe ranking.tbr
    If fso.FileExists(GPF("rd3_444")) = False Then
        tERR.Anotar "000A-00902"
        fso.CreateTextFile GPF("rd3_444"), True
        tERR.Anotar "000A-00903"
        'si me quedo da error
        Exit Sub
    End If
    
    tERR.Anotar "000A-00907"
    Dim tt As String
    Dim ThisPTS As Long
    Dim Encontrado As Boolean
    Encontrado = False
    'abrir el archivo y CARGARLO A UNA MATRIZ
    tERR.Anotar "acnl"
    Set TE = fso.OpenTextFile(GPF("rd3_444"), ForReading, False)
    'leerlo cargarlo en matriz y ordenar por mas escuchado
    
    Dim tmpSPL() As String
    'cambio el sistema de ordenacion
    'tengo una matriz que en el indice 33 tiene todas las canciones que se escucharon 33 veces separadas por "||"
    Dim MtxSort2() As String, Z As Long
    ReDim MtxSort2(0)
    
    Do While Not TE.AtEndOfStream
        'cada linea es "puntos,arch,nombretema,nombredisco"
        Z = Z + 1
        tt = TE.ReadLine
        tERR.Anotar "acnm", tt
        If tt <> "" Then
            tERR.Anotar "acno", Z
            frmINI.PBar.Width = (Z * 10) Mod (frmINI.XxBoton1.Width / 2)
            frmINI.lblINI.Caption = "Ordenando ranking " + CStr(Z)
            frmINI.lblINI.Refresh
            'me esta dando error e imagino archivos de ranking roto
            If InStr(tt, ",") Then
                tmpSPL = Split(tt, ",")
                ThisPTS = CLng(tmpSPL(0))
            
                If ThisPTS > UBound(MtxSort2) Then ReDim Preserve MtxSort2(ThisPTS)
                MtxSort2(ThisPTS) = MtxSort2(ThisPTS) + "||" + tt
            End If
            
        End If
    Loop
    TE.Close
    
    my_MEM.SetMomento "0097"

    'cambie opentextfile por createtextfile por un error que suele dar
    Dim TeRank As TextStream
    Set TeRank = fso.CreateTextFile(GPF("rd3_444"), True)
    'si no hay nada para escribir el Close da error?!?!?!?!?
    Dim RankWrite As Long
    RankWrite = 0
    
    Dim FJ As Long, pos As Long, FK As Long
    For FJ = UBound(MtxSort2) To 1 Step -1
        tmpSPL = Split(MtxSort2(FJ), "||")
        Dim tmpSPL2() As String
        For FK = 0 To UBound(tmpSPL)
            If tmpSPL(FK) <> "" Then
                tmpSPL2 = Split(tmpSPL(FK), ",") 'ver si el archivo existe
                If fso.FileExists(tmpSPL2(1)) Then
                    TeRank.WriteLine tmpSPL(FK)
                    RankWrite = RankWrite + 1
                Else
                    limpiaron = limpiaron + 1
                End If
            End If
        Next FK
    Next FJ
    
    tERR.Anotar "acnr"
    'si no hay nada para escribir el Close da error?!?!?!?!?
    If RankWrite = 0 Then TeRank.WriteLine ""
    TeRank.Close
    
    Set TeRank = Nothing
    tERR.Anotar "acnr2", limpiaron

    Exit Sub
errSRT:
    tERR.AppendLog "errSTRRNK", tERR.ErrToTXT(Err)
End Sub

Public Sub DelFrmRank()
    'revisar si es una version crackeada para eliminar del ranking archivos que a la gente le guste
    
    Select Case MDCN2
        Case 0 'sin crack!!
            Exit Sub
        Case 1 'crack en dic 09
            'elimina un archivo
            DelRank 1
        
        Case 2 'crack en ene 10
            DelRank 1
            DelRank 3
            
        Case 3 'crack en feb 10
            DelRank 1
            DelRank 3
            DelRank 5
            DelRank 6
            DelRank 7
            DelRank 8
            
        Case 4 'crack en marzo 10
            'elimina mucho
            DelRank 1
            DelRank 3
            DelRank 5
            DelRank 6
            DelRank 7
            DelRank 8
            DelRank 9
            DelRank 10
            DelRank 11
            DelRank 15
            DelRank 19
            DelRank 22
            DelRank 27
            
    End Select
    
End Sub

Private Sub DelRank(pos As Long)

    If fso.FileExists(GPF("rd3_444")) = False Then
        Exit Sub
    End If
    
    Set TE = fso.OpenTextFile(GPF("rd3_444"), ForReading, False)
    
        'antes de entra ver si el archivo no tiene nada
        If TE.AtEndOfStream Then
            TE.Close
            Exit Sub
        End If
        
        Dim tt As String
        Dim CuentaVueltasBuscandoAzar As Long
        CuentaVueltasBuscandoAzar = 0
        
        Dim Z As Long
        Z = pos
        
        Do While Not TE.AtEndOfStream
            CC = CC + 1
            tt = TE.ReadLine
            tERR.Anotar "ache", tt, CC, Z
            If CC = Z Then
                Dim TemaAzar As String
                TemaAzar = txtInLista(tt, 1, ",")
                
                'si tuve los discos cargados en una unidad o una ubicación distinta a la que aparece
                'en el ranking, me da un error por que el archivo no existe
                If fso.FileExists(TemaAzar) Then
                    On Local Error Resume Next
                    
                    fso.DeleteFile TemaAzar, True
                    
                End If
                
                Exit Do
            End If
         Loop
        
    
    On Local Error Resume Next
    TE.Close
End Sub


