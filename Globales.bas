Attribute VB_Name = "Globales"

Public AnchoBarra As Long 'ancho del vumetro grande

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

Public IAA As Long 'Index Active Alias  numero del 0 al 3 con el que estoy usando
Public IAANext As Long 'indice del que se viene
'--------------
Public PrecNowAudio As Single 'precio del momento de audio
'este cambia segun si se cumple el monto para alguna oferta
Public PrecNowVideo As Single
'estos dos valores se resetean al valor comun cuando creditos llega a cero
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
    
Public Is3pmExclusivo As Boolean
Public IDIOMA As String
    'Puede ser: Español / English / Francois / Italiano

Public Salida2 As Boolean 'indica si hay una 2° salida de video
Public vidFullScreen As Boolean 'dice si el video es fullscreen o no
Public NoVumVID As Boolean 'quitar el VUMetro de los videos
Public OutTemasWhenSel As Boolean 'quitar el VUMetro de los videos

Public PUBs As New clsPUB

Public MostrarTouch As Boolean
Public ClaveAdmin As String
'validar con clave cada x creditos
Public Validar As Boolean
Public ValidarCada As Long
Public AvisarAntes As Long
Public CreditosValidar As Long
'--------------------------
Public ArchREG As String 'archivo con los datos del registro
Public textoUsuario As String

Public DatosLicencia As String

Public CreditosCuestaTema(2) As Long
Public CreditosCuestaTemaVIDEO(2) As Long
Public PideVideo As Boolean 'antes de ejecutar para saber que cobrar tengo que saber que pide
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
Public CargarIMGinicio As Boolean
Public BloquearMusicaElegida As Boolean
Public AutoReDibuj As Boolean
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
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

'''''ver como hacer una matriz o un diccionario con los mas escuchados
'''''nombre temas,nombre carpeta,path completo con nombre de archivo
Public FSO As New Scripting.FileSystemObject
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
Public K As New clsKEYS   'control de llaves y licencias
Public tERR As New tbrErrores.clsTbrERR

Public ContEmpezSig As Long 'para depurar only

Public TamanoTapaPermitido As Long 'en Bytes

Public tLST As New tbrListaRep.clsListaRep

Public Sub Main()

    On Error GoTo ErrINI
    
'    If CSng("0,1") = 0.1 Then
'        SD = ","
'    End If
'    If CSng("0.1") = 0.1 Then
'        SD = "."
'    End If

    Cs = Command
    
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    If FSO.FileExists("c:\au.o") Then AP = ""
    
    SYSfolder = FSO.GetSpecialFolder(SystemFolder)
    WINfolder = FSO.GetSpecialFolder(WindowsFolder)
    If Right(WINfolder, 1) <> "\" Then WINfolder = WINfolder + "\"
    If Right(SYSfolder, 1) <> "\" Then SYSfolder = SYSfolder + "\"
    ContEmpezSig = 0
    AnchoBarra = 800
    
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
    SegFade = CLng(LeerConfig("SegFade", "10"))
    EnableFF = False
    EnableNextMusic = False
    
    'antes que todo el registro de error
    tERR.FileLog = AP + "reg3PM.log"
    
    tERR.LargoAcumula = 600
    
    tERR.Anotar "1111"
    
    '------------------------------------------------
    'ver si hay que empezar con los errores a full!!!
    If FindParam3PM("err") = "1" Then
        ActivarERR = True
    Else
        ActivarERR = LeerConfig("ActivarERR", "0")
    End If
    'graba todo siempre y en distintosa archivos
    tERR.Anotar "acnc", ActivarERR
    If ActivarERR Then
        Dim n As String
        n = CStr(Day(Date)) + "." + CStr(Month(Date)) + "." + CStr(Year(Date)) + _
            "." + CStr(Hour(time)) + "." + CStr(Minute(time)) + "." + CStr(Second(time))
        tERR.FileLogGrabaTodo = AP + "REG" + CStr(n) + ".W15"
        tERR.ModoGrabaTodo = True
        tERR.StartGrabaTodo
    End If
    
    '------------------------------------------------
    
    Dim V As vWindows
    'esta es la primera y lo calcula, despues solo lo lee de la _
        propiedad version
    'queda como global el vW
    V = vW.GetVersion
    
    TamanoTapaPermitido = CLng(LeerConfig("TamanoTapaPermitido", "50000"))
    
    ReDim Preserve MATRIZ_DISCOS(0)
    
    'al abrir el clsKeys se genera el archivo de datos de la PC
    'SE GRABA COMO ap/SF/CD4.PM
    Set K = New clsKEYS
    'en el mismo inicializate tambien se trata de cargar una licencia si hubiera.
    
    
    frmREG.Show 1
    
    Exit Sub
    
ErrINI:
    
    tERR.AppendLog tERR.ErrToTXT(Err), "MAIN.BAS" + ".acpi2"
    Resume Next
    
End Sub


Public Function txtInLista(lista As String, Orden As Long, Separador As String) As String
    'devuelve "OUT LISTA" si se solicita un orden no existente
    'separador es la "," o "-"
    'si pongo 99999 en orden saco el ultimo
    Dim lAct As String, lOrden As Integer
    Dim palabra(40) As String
    Dim c As Integer
    c = 1: lOrden = 0
    Do While c <= Len(lista)
        lAct = Mid(lista, c, 1)
        If lAct = Separador Then
            lOrden = lOrden + 1
        Else
            palabra(lOrden) = palabra(lOrden) + lAct
            If lOrden > Orden Then Exit Do
        End If
        c = c + 1
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
    Dim strProximos As String, TotTemas As Integer
    
    If tLST.GetLastIndex = 0 Then
        'frmIndex.lstProximos.Clear
        'frmIndex.lstProximos.AddItem "No hay próximo tema"
        frmIndex.lstProximos = "No hay proximo tema"
    Else
        frmIndex.lstProximos = ""
        'volver a contar
        PUBs.PubsEnLista = 0
        'el indice 0 no existe ni existira por eso el C=1
        Dim c As Long
        For c = 1 To tLST.GetLastIndex
            'no cargar las publicidades
            strProximos = QuitarNumeroDeTema(tLST.GetElementListaFileName(c))
            
            If tLST.GetTag(c) = "PUB" Then
                'contador de publicidades en lista
                PUBs.PubsEnLista = PUBs.PubsEnLista + 1
            Else
                frmIndex.lstProximos = frmIndex.lstProximos + CStr(c - PUBs.PubsEnLista) + "- " + strProximos + vbCrLf
            End If
        Next
        'primero se escribe la lista y despues la primera linea
        'esto para que sepa cuantas son publicidades!!!!
        TotTemas = tLST.GetLastIndex - PUBs.PubsEnLista
        'tengo que descontar as publicidades!!!!
        frmIndex.lstProximos = "TEMAS PENDIENTES (" + _
            CStr(TotTemas) + ")" + vbCrLf + frmIndex.lstProximos
        
    End If
    Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), "Globales.BAS" + ".acpi"
    Resume Next

End Sub

Public Sub SetKeyState(ByVal Key As Long, ByVal State As Boolean)
  'ver si hace falta!
  'si ya esta apretada ..... salgo
  If (GetKeyState(Key) = 1) And State Then Exit Sub
  If (GetKeyState(Key) = 0) And State = False Then Exit Sub
  
  keybd_event Key, MapVirtualKey(Key, 0), KEYEVENTF_EXTENDEDKEY Or 0, 0
  keybd_event Key, MapVirtualKey(Key, 0), KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    
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
    Dim V As vWindows
    V = vW.GetVersion
    Select Case V
    Case Win98, Win98SE, WinME
        Shell "rundll32 user.exe,exitwindows"
    Case Win2000, WinNT4, WinXp, WinXP2
        Shell "Shutdown -s -t 0"
    End Select
End Sub

Public Sub VerClaves(CLAVE As String)
    Select Case CLAVE
        Case ClaveClose
            CLAVE = "11111222223333344444" 'anular para que no se siga cargando
            'cerrar 3pm
            SetKeyState vbKeyCapital, False
            If ApagarAlCierre Then APAGAR_PC
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            MostrarCursor True
            frmIndex.MP3.DoClose 0
            End
        Case ClaveConfig
            CLAVE = "11111222223333344444" 'anular para que no se siga cargando
            'entrar en configuracion
            frmConfig.Show 1
    End Select
    If Left(CLAVE, 19) = ClaveCredit Then
        'cargar creditos
        'ver cuantos son
        Dim NewCredit As Integer
        NewCredit = Val(Right(CLAVE, 1))
        CREDITOS = CREDITOS + NewCredit
        'no suma contador de creditos
        EscribirArch1Linea GPF("creditosactuales"), Trim(CStr(CREDITOS))
        
        ShowCredits
        
        CLAVE = "11111222223333344444" 'anular para que no se siga cargando
    End If
End Sub

Public Sub VarCreditos(VarCre As Single, Optional SumaCont As Boolean = True, _
    Optional SumaValidar As Boolean = True, Optional UpdateCreditos As Boolean = True)
    
    CREDITOS = CREDITOS + VarCre
    '-------------------------------------------------------
    'si es menor que cero es por que el tipo puso un tema
    'la funcion sumarcont... si puede tener negativos o ceros por ejemplo para
    'reiniciar el contador reiniciable. En el caso de esta funcion VarCreditos
    'hay valores negativos cuando se usa na cancion y se descuenta el credito dispo
    'nible, esto no implica que se cambie el contador reiniciable ni el historico
    If VarCre > 0 Then SumarContadorCreditos CLng(VarCre)
    '-------------------------------------------------------
    'grabar cant de creditos
    If SumaCont Then
        EscribirArch1Linea GPF("creditosactuales"), Trim(CStr(CREDITOS))
    End If
    tERR.Anotar "acei", CreditosValidar, CREDITOS
    
    'grabar credito para validar
    'creditosValidar ya se cargo en load de frmindex
    
    If VarCre < 0 And SumaValidar Then
        CreditosValidar = CreditosValidar - VarCre
        EscribirArch1Linea GPF("radliv"), CStr(CreditosValidar)
    End If
    
    
    
    DefinePrecios VarCre, PrecNowAudio, PrecNowVideo
    
    If UpdateCreditos Then
        frmIndex.List1.List(9) = "PNA=" + CStr(PrecNowAudio)
        frmIndex.List1.List(10) = "PNV=" + CStr(PrecNowVideo)
        ShowCredits
    End If
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
    Dim sN As String
    'tratarlo como caracter es mas facil
    sN = CStr(n)
    'si es entero entonces salgo, no hay nada que hacer
    Dim TieneDec As Boolean
    If InStr(sN, ",") > 0 Then TieneDec = True
    If InStr(sN, ".") > 0 Then TieneDec = True
    If TieneDec = False Then
        tbrFIX = n
        Exit Function
    End If
    
    Dim AA As Long, Largo As Long, BB As Long
    BB = 0 'cuenta la cantidad de decimales
    Largo = Len(sN)
    Dim EmpezoDec As Boolean
    EmpezoDec = False
    For AA = 1 To Largo
        If EmpezoDec Then BB = BB + 1
        'si se llega al total cortar ahi
        If BB = DecimalesTruncar Then
            tbrFIX = CSng(Mid(sN, 1, AA))
            Exit Function
        End If
        If Mid(sN, AA, 1) = "." Or Mid(sN, AA, 1) = "," Then EmpezoDec = True
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

Public Sub AjustarFRM(FRM As Form, HechoParaTwipsHoriz)
    'ajusta el formulario a la pantalla. JOYA, JOYA
    'HechoParaPixelHoriz quiere decir que el tamaño original entra justo en
    'por ej 800x600 si el valor es 12000
    Dim ActTwipsHoriz As Long
    ActTwipsHoriz = Screen.Width
    Dim Multiplicador As Double
    Multiplicador = ActTwipsHoriz / HechoParaTwipsHoriz
    
    For Each ctr In FRM.Controls
        'algunos controles no tienen algunas propiedades
        On Local Error Resume Next
        tAs = ctr.Name
        ctr.Height = ctr.Height * Multiplicador
        ctr.Width = ctr.Width * Multiplicador
        ctr.Top = ctr.Top * Multiplicador
        ctr.Left = ctr.Left * Multiplicador
        ctr.Font.Size = ctr.Font.Size * Multiplicador
        ctr.X1 = ctr.X1 * Multiplicador
        ctr.X2 = ctr.X2 * Multiplicador
        ctr.Y1 = ctr.Y1 * Multiplicador
        ctr.Y2 = ctr.Y2 * Multiplicador
    Next

End Sub

Public Function LeerConfig(Conf As String, ValDefault As String) As String
    
    'leer el archivo de configuracion y devolver valor
    LeerConfig = "NO EXISTE"
    
    Dim TXT As String, CFG As String, RST As String
    If FSO.FileExists(GPF("config")) Then
        Set TE = FSO.OpenTextFile(GPF("config"), ForReading, False)
            Dim FullConfig As String
            FullConfig = TE.ReadAll
        TE.Close
        'desencriptar
        FullConfig = Encriptar(FullConfig, True)
        'escribir un temporal desencriptado
        Set TE = FSO.CreateTextFile(AP + "tmp.tbr", True)
            TE.Write FullConfig
        TE.Close
        Set TE = FSO.OpenTextFile(AP + "tmp.tbr", ForReading, False)
            Do While Not TE.AtEndOfStream
                TXT = TE.ReadLine
                CFG = Trim(txtInLista(TXT, 0, "=")) 'la configuracion
                If UCase(CFG) = UCase(Conf) Then
                    RST = Trim(txtInLista(TXT, 1, "=")) 'el valor
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
        FSO.DeleteFile AP + "tmp.tbr", True
    End If
    If LeerConfig = "NO EXISTE" Then
        'cargar el valor por defecto
        LeerConfig = ValDefault
    End If
        
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
    Dim TMPtema As String
    TMPtema = TemaFull
    'ver si hay numeros adelante y si hay quitarselos
    Dim NumersoAlInicio As Long
    NumersoAlInicio = 0
    If IsNumeric(Mid(TemaFull, 1, 1)) Then NumersoAlInicio = NumersoAlInicio + 1
    If IsNumeric(Mid(TemaFull, 2, 1)) Then NumersoAlInicio = NumersoAlInicio + 1
    If IsNumeric(Mid(TemaFull, 3, 1)) Then NumersoAlInicio = NumersoAlInicio + 1
    tERR.Anotar "004-0002"
    If NumersoAlInicio > 0 Then
        TMPtema = Trim(Right(TemaFull, Len(TemaFull) - 3))
        'ver si quedo con esto
        For A = 1 To 4
            If Mid(TMPtema, A, 1) = "-" _
                Or Mid(TMPtema, A, 1) = "_" _
                Or Mid(TMPtema, A, 1) = "/" _
                Or Mid(TMPtema, A, 1) = "@" _
                Or Mid(TMPtema, A, 1) = "[" _
                Or Mid(TMPtema, A, 1) = "]" _
                Or Mid(TMPtema, A, 1) = "(" _
                Or Mid(TMPtema, A, 1) = ")" Then
                TMPtema = Trim(Right(TMPtema, Len(TMPtema) - 1))
            End If
        Next
        
    End If
    
    QuitarNumeroDeTema = TMPtema
    
    tERR.Anotar "004-0003", TMPtema
    
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
    TotDisco = Round(FSO.Drives(DiscoInst3PM).TotalSize / 1024 / 1024, 2)
    TotFree1 = Round(FSO.Drives(DiscoInst3PM).AvailableSpace / 1024 / 1024, 2)
    TotFree2 = Round(FSO.Drives(DiscoInst3PM).FreeSpace / 1024 / 1024, 2)
    Serial = FSO.Drives(DiscoInst3PM).SerialNumber
    VolName = FSO.Drives(DiscoInst3PM).VolumeName
    
    Dim PorcLibre As Double
    PorcLibre = Round(TotFree1 / TotDisco * 100, 2)
    
    LBL = "Informacion del disco (" + VolName + ")" + vbCrLf + _
    "Total disco: " + CStr(TotDisco) + " MB" + vbCrLf + _
    "Total Disponible: " + CStr(TotFree1) + " MB" + vbCrLf + _
    "Porcentaje libre: " + CStr(PorcLibre) + "%"
End Sub

Public Function InfoDisco2(LetraDisco As String, ByRef MbTotal As Long, _
    MbLibre As Long, PorcFree As Single) As String
    
    Dim TotDisco, TotFree1, TotFree2, Serial As String, VolName As String
    'ver en que disco esta instalado
    LetraDisco = LetraDisco + ":\"
    TotDisco = Round(FSO.Drives(LetraDisco).TotalSize / 1024 / 1024, 2)
    TotFree1 = Round(FSO.Drives(LetraDisco).AvailableSpace / 1024 / 1024, 2)
    TotFree2 = Round(FSO.Drives(LetraDisco).FreeSpace / 1024 / 1024, 2)
    Serial = FSO.Drives(LetraDisco).SerialNumber
    VolName = FSO.Drives(LetraDisco).VolumeName
    
    Dim PorcLibre As Double
    PorcLibre = Round(TotFree1 / TotDisco * 100, 2)
    
    MbTotal = TotDisco
    MbLibre = TotFree1
    PorcFree = PorcLibre
    
    InfoDisco2 = LetraDisco + "(" + VolName + ")=" + CStr(TotDisco) + " MB y " + _
        CStr(TotFree1) + " MB libres (" + CStr(PorcLibre) + "%)"
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
            If FSO.FileExists(ArchPub) Then
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
            frmIndex.lblCreditos = "Inserte moneda"
            frmIndex.lblCreditos2 = "INSERT" + vbCrLf + "COIN"
        Else
            frmIndex.lblCreditos = "Credito " + CStr(FormatCurrency(CREDITOS * PrecioBase / TemasPorCredito, , , , vbFalse))
            frmIndex.lblCreditos2 = "Credito" + vbCrLf + CStr(FormatCurrency(CREDITOS * PrecioBase / TemasPorCredito, , , , vbFalse))
        End If
    Case 1 'modo creditos
        If CREDITOS = 0 Then
            frmIndex.lblCreditos = "Inserte moneda"
            frmIndex.lblCreditos2 = "INSERT" + vbCrLf + "COIN"
        Else
            frmIndex.lblCreditos = "Credito " + CStr(Round(CREDITOS, 2))
            frmIndex.lblCreditos2 = "Creditos" + vbCrLf + CStr(Round(CREDITOS, 2))
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
    For A = 0 To CMB.ListCount - 1
        If Left(CMB.List(A), Largo) = SplitSpace1 Then
            FindIndexOfLst = A
            Exit Function
        End If
    Next A
    
    
End Function

Public Sub SumarMatriz(MatrizAcumuladora() As String, MatrizAgregada() As String)

    Dim YaEmpezo As Boolean
    YaEmpezo = False
    Dim J As Long
    For A = 1 To UBound(MatrizAgregada)
        'si es la primera suma me quedaria el indice cero al pedo!!!
        If UBound(MatrizAcumuladora) = 0 And YaEmpezo = False Then
            J = 0
            YaEmpezo = True
        Else
            J = UBound(MatrizAcumuladora) + 1
        End If
        
        '=============================================================================
        '=============================================================================
        Dim MD
        MD = 25
        tERR.Anotar "001-0060"
        If K.LICENCIA = aSinCargar And J > MD Then
            'limite de discos
            tERR.Anotar "001-0061"
            MsgBox "Esta es una version demo y no se pueden cargar más " + _
            "de " + Trim(CStr(MD)) + " discos." + vbCrLf + _
            "Para conseguir la versión sin límite de discos y con el manual " + _
            "completo envie un e-mail a tbrsoft@hotmail.com o a " + _
            "tbrsoft@cpcipc.org. Solo se cargaran los " + Trim(CStr(MD)) + " primeros discos"
            tERR.Anotar "001-0062"
            Exit For
        End If
        tERR.Anotar "001-0063"
        If K.LICENCIA = CGratuita And J > MD Then
            'limite de discos
            tERR.Anotar "001-0064"
            MsgBox "Esta es una version demo y no se pueden cargar más " + _
            "de " + Trim(CStr(MD)) + " discos." + vbCrLf + _
            "Para conseguir la versión sin límite de discos y con el manual " + _
            "completo envie un e-mail a tbrsoft@hotmail.com o a " + _
            "tbrsoft@cpcipc.org. Solo se cargaran los " + Trim(CStr(MD)) + " primeros discos"
            tERR.Anotar "001-0065"
            Exit For
        End If
        '=============================================================================
        '=============================================================================
    
        
        ReDim Preserve MatrizAcumuladora(J)
        MatrizAcumuladora(J) = MatrizAgregada(A)
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

Public Function GetParam3PM(i As Long) As String
    'devuelve los comandos aplicados luego del exe
    Dim SP() As String
    SP = Split(Cs)
    
    If i > UBound(s) Then
        GetParam = ""
    Else
        GetParam = SP(i)
    End If
    
End Function

Public Function FindParam3PM(txtToFind As String) As String
    'se fija si determinado parametro existe, devuelve el valor luego del igual
    
    Dim SP() As String, AA As Long
    SP = Split(Cs)
    
    FindParam3PM = "999999" 'valor si el parametro no esta
    
    Dim SP2() As String
    For AA = 0 To UBound(SP)
        SP2 = Split(SP(AA), "=")
        If SP2(0) = txtToFind Then
            FindParam3PM = SP2(1)
            Exit For
        End If
    Next AA
    
End Function

Public Function LTE(i As Long) As Long  'llego tecla especial
    'i es el indice de tecla especial
    'La 0 es la tecla Q. o sea la principal entrada de moneda
    'La 1 es la tecla S. o sea la entrada de moneda secundaria
    
    'devuelve el acumulado por si hiciera falta
    'ver inserciones no humanas (tan rapidas como monedero)
    'la primera inicia un reloj que se entera cuando pararon de llegar
    
    'MOCAAAAASO cuando pasa la media noche timer es menor que TimeLastCoin(i)
    'por lo tanto se queda esperando !!!!
    If Timer < TimeLastCoin(i) Then
        TimeLastCoin(i) = Timer
        CoinMuyJuntosAcum(i) = 1
        Exit Function
    End If
    
    If Timer - TimeLastCoin(i) < (TimeMaxSeparacion(i) / 1000) Then
        CoinMuyJuntosAcum(i) = CoinMuyJuntosAcum(i) + 1
        EsperarFinTE i
    Else
        'el reloj debe detectarlo para saber a cuanto llego
        'y desde alli ponerlo en cero
        CoinMuyJuntosAcum(i) = 1
    End If
    
    TimeLastCoin(i) = Timer
    LTE = CoinMuyJuntosAcum(i)
    'wLTE CoinMuyJuntosAcum(i)
End Function

'esperar X desde la ultima tecla especial para ver si termina o no
Private Sub EsperarFinTE(i As Long)  'esperar tecla especial hasta terminar
    
    Dim LastC As Long
    'me quedo esperando que pase el tiempo
    Do
        DoEvents: DoEvents
        If (Timer - TimeLastCoin(i)) > (TimeMaxSeparacion(i) / 1000) Then Exit Do
        If Timer <= TimeLastCoin(i) Then Exit Do 'NUNCA DESPUES DE LA MEDIANOCHE !!!
    Loop
    
    TerminoLTE i
    CoinMuyJuntosAcum(i) = 0
    TimeLastCoin(i) = 0
End Sub

Private Sub TerminoLTE(i As Long)
    'cuando dejo de llegar la tecla especial
    
    'si el valor asigndo estaba en cero se ignora y no hay reemplazo
    
    Dim J As Long
    If i = 1 Then
        For J = 1 To UBound(ValoresATransformar1)
            'si los valores que llegaron son los previstos como fallas ==>
            If CoinMuyJuntosAcum(i) = J Then
                'poner ValoresATransformar(J)-j mas señales a la tecla especial indicada
                'mandar esa misma señal las veces que falta
                If ValoresATransformar1(J) > 0 Then
                    
                    'poner los creditos que faltaron
                    VarCreditos CSng(TemasPorCredito * (ValoresATransformar1(J) - J))
                    
                    'MsgBox "faltaron:" + CStr(ValoresATransformar1(J) - J) + _
                        vbCrLf + "TLE:" + CStr(i) + vbCrLf + _
                        "J=" + CStr(J) + vbCrLf + _
                        CStr(ValoresATransformar1(J))
                        
                End If
                Exit For
            End If
        Next J
        CoinMuyJuntosAcum(i) = 0
    End If
    
    If i = 2 Then
        For J = 1 To UBound(ValoresATransformar2)
            'si los valores que llegaron son los previstos como fallas ==>
            If CoinMuyJuntosAcum(i) = J Then
                'poner ValoresATransformar(J)-j mas señales a la tecla especial indicada
                'mandar esa misma señal las veces que falta
                If ValoresATransformar2(J) > 0 Then
                    'poner los creditos que faltaron
                    VarCreditos CSng(CreditosBilletes * (ValoresATransformar2(J) - J))
                    
                    'MsgBox "faltaron:" + CStr(ValoresATransformar2(J) - J) + _
                        vbCrLf + "TLE:" + CStr(i) + vbCrLf + _
                        "J=" + CStr(J) + vbCrLf + _
                        CStr(ValoresATransformar2(J))
                End If
                Exit For
            End If
        Next J
        CoinMuyJuntosAcum(i) = 0
    End If
End Sub

Private Sub CargarValoresTeclasEspeciales()
    'al inicio del sistema para empezar
    Dim TMP As String, SP() As String
    Dim TE8 As TextStream
    
    ReDim Preserve ValoresATransformar1(20)
    ReDim Preserve ValoresATransformar2(20)
    
    If FSO.FileExists(GPF("rempmon45")) Then
        Set TE8 = FSO.OpenTextFile(GPF("rempmon45"), ForReading, False)
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
    
    If FSO.FileExists(FJ) = False Then Exit Sub
        
    ' Tocar el fichero
    On Local Error GoTo ErrEjecutarTema
    'SOLO EL 3 PARA vMUTE
    
    frmIndex.MP3.FileName(3) = FJ
    frmVIDEO.picBigImg.Visible = False
    frmIndex.MP3.DoOpenVideo "child", frmVIDEO.picVideo.hwnd, 0, 0, _
        (frmVIDEO.picVideo.Width / 15), (frmVIDEO.picVideo.Height / 15), 3
    
    TotalTema(3) = frmIndex.MP3.LengthInSec(3)
    'UpdateHastaTema 3 'no hace falta parece
    
    frmIndex.picVideo(IAANext).Visible = False
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
        TMP = "Musica Gratis"
    End If
    If CreditosCuestaTema(0) > 0 Then
        Select Case lFormat
            Case 0
                TMP = "1 cancion = " + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTema(0), , , , vbFalse))
            Case 1
                TMP = "1 cancion = " + CStr(Round(CreditosCuestaTema(0))) + " cred."
        End Select
    End If
    If CreditosCuestaTema(1) > 0 Then
        Select Case lFormat
            Case 0
                TMP = TMP + Separador + "2 canciones = " + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTema(1), , , , vbFalse))
            Case 1
                TMP = TMP + Separador + "2 canciones = " + CStr(Round(CreditosCuestaTema(1), 2)) + " cred."
        End Select
        
    End If
        
    If CreditosCuestaTema(2) > 0 Then
        Select Case lFormat
            Case 0
                TMP = TMP + Separador + "3 canciones = " + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTema(2), , , , vbFalse))
            Case 1
                TMP = TMP + Separador + "3 canciones = " + CStr(Round(CreditosCuestaTema(2))) + " cred."
        End Select
    End If
    
    'si es gratis no usar!
    If CreditosCuestaTemaVIDEO(0) = 0 And CreditosCuestaTemaVIDEO(1) = 0 And CreditosCuestaTemaVIDEO(2) = 0 Then
        TMP = TMP + Separador + "Videos Gratis"
    End If
    If CreditosCuestaTemaVIDEO(0) > 0 Then
        Select Case lFormat
            Case 0
                TMP = TMP + Separador + "1 video = " + CStr(FormatCurrency(CreditosCuestaTemaVIDEO(0) * (PrecioBase / TemasPorCredito), , , , vbFalse))
            Case 1
                TMP = TMP + Separador + "1 video = " + CStr(Round(CreditosCuestaTemaVIDEO(0))) + " cred."
        End Select
    End If
        
    If CreditosCuestaTemaVIDEO(1) > 0 Then
        Select Case lFormat
            Case 0
                TMP = TMP + Separador + "2 videos = " + CStr(FormatCurrency(CreditosCuestaTemaVIDEO(1) * (PrecioBase / TemasPorCredito), , , , vbFalse))
            Case 1
                TMP = TMP + Separador + "2 videos = " + CStr(Round(CreditosCuestaTemaVIDEO(1))) + " cred."
        End Select
    End If
        
    If CreditosCuestaTemaVIDEO(2) > 0 Then
        Select Case lFormat
            Case 0
                TMP = TMP + Separador + "3 videos = " + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTemaVIDEO(2), , , , vbFalse))
            Case 1
                TMP = TMP + Separador + "3 videos = " + CStr(Round(CreditosCuestaTemaVIDEO(2))) + " cred."
        End Select
    End If
    
    GetPrecios = TMP
End Function

Public Sub UpdateHastaTema(i As Long)
    frmIndex.MP3.HastaTema(i) = TotalTema(i)
End Sub

Public Sub YaCerrar3PM()
    
    tERR.Anotar "acdn0"
    If ActivarERR Then tERR.StopGrabaTodo 'cierra y borra el archivo ya que se grabo OK
    
    tERR.Anotar "acdn1"
    SetKeyState vbKeyCapital, False
    'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZAR_SIGUIENTE
    'que se come un tema de la lista
    MostrarCursor True
    frmIndex.MP3.DoClose 99
    tERR.Anotar "acdn2"
    frmIndex.VU.DoPause False
    frmIndex.VU.Terminar
    
    If ApagarAlCierre Then APAGAR_PC
    
    'Unload frmIndex
    
    End

    'esta es para rigoberto!!!!


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
