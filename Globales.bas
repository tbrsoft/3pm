Attribute VB_Name = "Globales"
'--------------
Public PrecNowAudio As Single 'precio del momento de audio
'este cambia segun si se cumple el monto para alguna oferta
Public PrecNowVideo As Single
'estos dos valores se resetean al valor comun cuando creditos llega a cero
'--------------

Public PrecioBase As Single
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

Public MostrarPUB As Boolean 'se reproducen Publicidades MP3 o video?
Public PubliCada As Long 'cada cuantos temas la publicidad

Public MostrarPUBIMG As Boolean 'se muestran Publicidades (imagen rotativa en index)?
Public PubliIMGCada As Long 'cada cuantos segundos la publicidad

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
Public CORTAR_TEMA As Boolean 'indica si el tema que se esta ejecutando se debe cortar
'esto puede ser porque es una version demo o por que el tema que se ejecuta es uno
Public Protector As Long '0=inhabilitado 1=Original 2=Carpeta Fotos 3= Video FullScreen
Public TECLAS_PRES As String 'las ultimas 20 teclas presionadas
Public ExtActual As String 'extencion del ultimo archivo elegido
'para el teclado
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long

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
Public MATRIZ_LISTA() As String 'lista de temas a reproducir. No incluye el TEMA_REPRODUCIENDO
Public TOTAL_DISCOS As Long ' total de discos
Public UbicDiscoActual As String 'path del disco actual
'sirve para no usar la MATRIZ_TEMAS y poder ordenar los temas de cada disco
Public WAIT_EMPIEZA As Integer 'esperar 5 segundos por comienzo de tema
Public K 'control de llaves y licencias
Public tERR As New tbrErrores.clsTbrERR

Public Sub Main()
    
    On Error GoTo ErrINI
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    SYSfolder = FSO.GetSpecialFolder(SystemFolder)
    WINfolder = FSO.GetSpecialFolder(WindowsFolder)
    If Right(WINfolder, 1) <> "\" Then WINfolder = WINfolder + "\"
    If Right(SYSfolder, 1) <> "\" Then SYSfolder = SYSfolder + "\"
    
    'antes que todo el registro de error
    tERR.FileLog = AP + "reg3PM.log"
    
    tERR.LargoAcumula = 130
    
    tERR.Anotar "1111"
    
    '------------------------------------------------
    'ver si hay que empezar con los errores a full!!!
    ActivarERR = LeerConfig("ActivarERR", "0")
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
    
    ReDim Preserve MATRIZ_DISCOS(0)
    
    Set K = New clsKEYS
    K.ClaveDLL = "ashjdklahsJKLHASL65456456456"
    
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
    
    If UBound(MATRIZ_LISTA) = 0 Then
        'frmIndex.lstProximos.Clear
        'frmIndex.lstProximos.AddItem "No hay próximo tema"
        frmIndex.lstProximos = "No hay proximo tema"
    Else
        frmIndex.lstProximos = ""
        'volver a contar
        PUBs.PubsEnLista = 0
        'el indice 0 no existe ni existira por eso el C=1
        For c = 1 To UBound(MATRIZ_LISTA)
            'no cargar las publicidades
            strProximos = QuitarNumeroDeTema(txtInLista(MATRIZ_LISTA(c), 1, ","))
            'frmIndex.lstProximos.AddItem CStr(c) + "- " + strProximos
            If strProximos = "Publicidad" Then
                'contador de publicidades en lista
                PUBs.PubsEnLista = PUBs.PubsEnLista + 1
            Else
                frmIndex.lstProximos = frmIndex.lstProximos + CStr(c - PUBs.PubsEnLista) + "- " + strProximos + vbCrLf
            End If
        Next
        'primero se escribe la lista y despues la primera linea
        'esto para que sepa cuantas son publicidades!!!!
        TotTemas = UBound(MATRIZ_LISTA)
        'tengo que descontar as publicidades!!!!
        frmIndex.lstProximos = "TEMAS PENDIENTES (" + _
            CStr(TotTemas - PUBs.PubsEnLista) + ")" + vbCrLf + frmIndex.lstProximos
        
    End If
    Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), "Globales.BAS" + ".acpi"
    Resume Next

End Sub

Public Sub OnOffCAPS(vKey As KeyCodeConstants, PRENDER As Boolean)
    Dim keys(255) As Byte
    ' leer el estado actual del teclado
    GetKeyboardState keys(0)
    ' invertir el bit 0 de la tecla virtual en la que estamos interesados
    ' keys(vKey) = keys(vKey) Xor 1
    If PRENDER Then
        keys(vKey) = 1
    Else
        keys(vKey) = 0
    End If
    ' forzar el nuevo estado del teclado
    SetKeyboardState keys(0)
End Sub

Public Function Tecla(n As Integer) As String
    Select Case n
        'las letras son iguales
        Case 65 To 90
            Tecla = Chr(n) + " :" + Trim(Str(n))
        'los numeros tambien
        Case 48 To 57
            Tecla = Chr(n) + " :" + Trim(Str(n))
        'el numpad debe escribir numeros (48-57)
        Case 96 To 105
            Tecla = Chr(n - 48) + " :" + Trim(Str(n))
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
    On Error GoTo MiErr
    Set TE = FSO.CreateTextFile(AP + "reini.tbr", True)
    Select Case ModoReini
        Case "FULL" 'tema actual + lista posterior
            Dim nombreTEMA As String, nombreDISCO As String
            nombreTEMA = FSO.GetBaseName(TEMA_REPRODUCIENDO)
            nombreDISCO = FSO.GetBaseName(FSO.GetParentFolderName(TEMA_REPRODUCIENDO))
            TE.WriteLine TEMA_REPRODUCIENDO + "," + nombreTEMA + "," + nombreDISCO
            '''ver como es la matriz_lista
            '''MATRIZ_LISTA(NewIndLista + 1) = temaElegido + "," + lstTemas + " / " + FSO.GetBaseName(UbicDiscoActual)
            For CC = 1 To UBound(MATRIZ_LISTA)
                TE.WriteLine MATRIZ_LISTA(CC)
            Next
            TE.Close
        Case "LISTA" 'solo la lista despues del tema actual
            For CC = 1 To UBound(MATRIZ_LISTA)
                TE.WriteLine MATRIZ_LISTA(CC)
            Next
            TE.Close
        Case "NADA"
            TE.WriteLine ""
            TE.Close
    End Select
    Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), "GLOBALES.bas" + ".acpj"
    Resume Next
End Sub

Public Function STRceros(n As Variant, Cifras As Integer) As String
    'n es el numero y cifras es la cantidad final de cifras del str terminado
    'devuelve ej : para 232,6 = 000232 para 1902,12 = 000000001902
    'complaeta con ceroas adelante
    ' si n es mas lasgo que cifras devuelve el valor n sin ningun cero adelante
    Dim STRn As String
    STRn = Trim(Str(n))
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
            OnOffCAPS vbKeyCapital, False
            If ApagarAlCierre Then APAGAR_PC
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            MostrarCursor True
            frmIndex.MP3.DoClose
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
        EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
        
        ShowCredits
        
        CLAVE = "11111222223333344444" 'anular para que no se siga cargando
    End If
End Sub

Public Sub VarCreditos(VarCre As Single)
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
    EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
    tERR.Anotar "acei", CreditosValidar, CREDITOS
    ShowCredits
    'grabar credito para validar
    'creditosValidar ya se cargo en load de frmindex
    If VarCre < 0 Then
        CreditosValidar = CreditosValidar - VarCre
        EscribirArch1Linea SYSfolder + "radilav.cfg", CStr(CreditosValidar)
        'si se ejecutaron canciones o videos y los creditos llegan hasta un valor
        'menor de una cancion en la maxima oferta disponible
        'enonces el precio vuelve a lo normal
        If CREDITOS < GetPrecioAudioMasBarato Then
            CREDITOS = 0
            PrecNowAudio = CreditosCuestaTema(0)
            PrecNowVideo = CreditosCuestaTemaVIDEO(0)
            ShowCredits
        End If
    End If
    
    'si se pusieron monedas entonces el precio puede cambiar
    If VarCre > 0 Then
        'si puso varias monedas bajar los precios
        If CREDITOS >= CreditosCuestaTema(1) And CreditosCuestaTema(1) > 0 Then
            PrecNowAudio = tbrFIX(Round(CreditosCuestaTema(1) / 2, 4), 2)
            '(porque son los creditos xa 2 canciones)
        End If
        
        If CREDITOS >= CreditosCuestaTema(2) And CreditosCuestaTema(2) > 0 Then
            PrecNowAudio = tbrFIX(Round(CreditosCuestaTema(2) / 3, 4), 2)
            '(porque son los creditos xa 3 canciones)
        End If
        
        If CREDITOS >= CreditosCuestaTemaVIDEO(1) And CreditosCuestaTemaVIDEO(1) > 0 Then
            PrecNowVideo = tbrFIX(Round(CreditosCuestaTemaVIDEO(1) / 2, 4), 2)
            '(porque son los creditos xa 2 canciones)
        End If
        
        If CREDITOS >= CreditosCuestaTemaVIDEO(2) And CreditosCuestaTemaVIDEO(2) > 0 Then
            PrecNowVideo = tbrFIX(Round(CreditosCuestaTemaVIDEO(2) / 3, 4), 2)
            '(porque son los creditos xa 3 canciones)
        End If
    End If
    
    'frmIndex.p1.Cls
    'frmIndex.p1.Print "Audio:" + CStr(PrecNowAudio) + " / Video:" + CStr(PrecNowVideo)
    
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

Public Function GetPrecioAudioMasBarato() As Long
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
    If FSO.FileExists(SYSfolder + "3pmcfg.tbr") Then
        Set TE = FSO.OpenTextFile(SYSfolder + "3pmcfg.tbr", ForReading, False)
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
    'la uso para el SYSfolder + "3pmcfg.tbr"
    
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

'Public Sub WriteTBRLog(TXT As String, PonerFecha As Boolean)
'
'    TXT = vbCrLf + "Linea: " + LineaError + vbCrLf + TXT
'    If FSO.FileExists(AP + "TBRlog.txt") = False Then
'        Set TE = FSO.CreateTextFile(AP + "TBRlog.txt", False)
'        TE.Close
'    End If
'    'ver si no es demasiado grande
'    If FileLen(AP + "tbrlog.txt") > 100000 Then 'hasta 100 KB aguanto
'        'pasarlo a otro archivo y volver a vrearlo
'        If FSO.FileExists(AP + "OLDtbrlog.txt") Then FSO.DeleteFile AP + "OLDtbrlog.txt", True
'        FSO.MoveFile AP + "tbrlog.txt", AP + "OLDtbrlog.txt"
'        Set TE = FSO.CreateTextFile(AP + "TBRlog.txt", False)
'        TE.Close
'    End If
'    'finalmente escribir
'    Set TE = FSO.OpenTextFile(AP + "TBRlog.txt", ForAppending, False)
'    TE.WriteLine "" 'dejar un renglon en blanco
'    If PonerFecha Then
'        TE.WriteLine Trim(Str(Date)) + " / " + Trim(Str(time)) + vbCrLf + TXT
'    Else
'        TE.WriteLine TXT
'    End If
'    TE.Close
'
'    If LCase(AP) = "h:\ahora\3pmv65 kabalin\" Then MsgBox TXT
'End Sub


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
                Dim NewInd As Long
                NewInd = UBound(MATRIZ_LISTA) + 1
                ReDim Preserve MATRIZ_LISTA(NewInd)
                'se graba en Matriz_Listas como patah, nombre(sin .mp3)
                MATRIZ_LISTA(NewInd) = ArchPub + "," + "Publicidad"
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
    'frmIndex.lblPuesto = CStr(CREDITOS)
    If CREDITOS = 0 Then
        frmIndex.lblCreditos = "Credito $ 0"
        frmIndex.lblCreditos2 = "INSERT" + vbCrLf + "COIN"
    Else
        frmIndex.lblCreditos = "Credito " + CStr(FormatCurrency(CREDITOS * PrecioBase / TemasPorCredito, , , , vbFalse))
        frmIndex.lblCreditos2 = "Credito" + vbCrLf + CStr(FormatCurrency(CREDITOS * PrecioBase / TemasPorCredito, , , , vbFalse))
    End If

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
    Dim j As Long
    For A = 1 To UBound(MatrizAgregada)
        'si es la primera suma me quedaria el indice cero al pedo!!!
        If UBound(MatrizAcumuladora) = 0 And YaEmpezo = False Then
            j = 0
            YaEmpezo = True
        Else
            j = UBound(MatrizAcumuladora) + 1
        End If
        
        '=============================================================================
        '=============================================================================
        Dim MD
        MD = 25
        tERR.Anotar "001-0060"
        If K.LICENCIA = aSinCargar And j > MD Then
            'limite de discos
            tERR.Anotar "001-0061"
            MsgBox "Esta es una version demo y no se pueden cargar más " + _
            "de " + Trim(Str(MD)) + " discos." + vbCrLf + _
            "Para conseguir la versión sin límite de discos y con el manual " + _
            "completo envie un e-mail a tbrsoft@hotmail.com o a " + _
            "tbrsoft@cpcipc.org. Solo se cargaran los " + Trim(Str(MD)) + " primeros discos"
            tERR.Anotar "001-0062"
            Exit For
        End If
        tERR.Anotar "001-0063"
        If K.LICENCIA = CGratuita And j > MD Then
            'limite de discos
            tERR.Anotar "001-0064"
            MsgBox "Esta es una version demo y no se pueden cargar más " + _
            "de " + Trim(Str(MD)) + " discos." + vbCrLf + _
            "Para conseguir la versión sin límite de discos y con el manual " + _
            "completo envie un e-mail a tbrsoft@hotmail.com o a " + _
            "tbrsoft@cpcipc.org. Solo se cargaran los " + Trim(Str(MD)) + " primeros discos"
            tERR.Anotar "001-0065"
            Exit For
        End If
        '=============================================================================
        '=============================================================================
    
        
        ReDim Preserve MatrizAcumuladora(j)
        MatrizAcumuladora(j) = MatrizAgregada(A)
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

