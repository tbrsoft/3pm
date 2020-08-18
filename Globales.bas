Attribute VB_Name = "Globales"
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

Public CreditosCuestaTema As Long
Public TemasPorCredito As Long

Public TE As TextStream
'claves para entrar a config, dar creditos y cerrar el sistema
Public ClaveConfig As String
Public ClaveCredit As String
Public ClaveClose As String

Public SYSfolder As String
Public WINfolder As String

Public RankToPeople As Boolean 'expone o no el reank a los usuarios


Public TypeVersion As String
'puede ser DEMO o FULL o SUPERLICENCIA

Public ClaveIngresada As String

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
Public MaximoFichas As Integer
Public ApagarAlCierre As Boolean
Public FASTini As Boolean 'comienzo con sin mostrar imágenes
Public EsperaMinutos As Integer 'en realizadad es SEGUNDOS. Espera antes de que auto ejecute algun temas
Public EsperaTecla As Integer '. Espera antes del protector de pantalla
Public ReINI As String 'Valores de ReIni FULL=tema ejecutando y lista LISTA=solo lista NADA=arranca de cero
Public VolumenIni As Long
Public PorcentajeTEMA As Integer 'del 0 al 100 para ver que parte se ejecuta de las muestras
Public CORTAR_TEMA As Boolean 'indica si el tema que se esta ejecutando se debe cortar
'esto puede ser porque es una version demo o por que el tema que se ejecuta es uno
'al azar que no se pasa entero
Public ProtectOriginal As Boolean 'true carga el protector de pantalla original. False es alguna carpeta con fotos
Public TECLAS_PRES As String 'las ultimas 20 teclas presionadas
Public ExtActual As String 'extencion del ultimo archivo elegido
'para el teclado
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long

'''''ver como hacer una matriz o un diccionario con los mas escuchados
'''''nombre temas,nombre carpeta,path completo con nombre de archivo
Public FSO As New Scripting.FileSystemObject
Public AP As String
Public CREDITOS As Long ' fichas cargadas (o temas habilitados para cargar)
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
    On Error GoTo Errores
    'cargar lstProximos
    Dim strProximos As String, TotTemas As Integer
    If UBound(MATRIZ_LISTA) = 0 Then
        frmIndex.lstProximos.Clear
        frmIndex.lstProximos.AddItem "No hay próximo tema"
    Else
        TotTemas = UBound(MATRIZ_LISTA)
        frmIndex.lstProximos.Clear
        frmIndex.lstProximos.AddItem "TEMAS PENDIENTES (" + CStr(TotTemas) + ")"
        For c = 1 To UBound(MATRIZ_LISTA)
            'el indice 0 no existe ni existira por eso el C=1
            strProximos = QuitarNumeroDeTema(txtInLista(MATRIZ_LISTA(c), 1, ","))
            frmIndex.lstProximos.AddItem CStr(c) + "- " + strProximos
        Next
    End If
    Exit Sub
Errores:
    WriteTBRLog "Error en Sub_CargarProximosTemas: " + Err.Description + " (" + CStr(Err.Number) + "). Se continua...", True
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
    On Error GoTo Errores
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
Errores:
    WriteTBRLog "Error en CargarArchReIni: " + Err.Description + " (" + CStr(Err.Number) + "). Se continua...", True
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
    Shell "rundll32 user.exe,exitwindows"
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
        If CREDITOS >= 10 Then
            frmIndex.lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
        Else
            frmIndex.lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
        End If
        
        CLAVE = "11111222223333344444" 'anular para que no se siga cargando
    End If
End Sub

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
    If FSO.FileExists(SYSfolder + "\3pmcfg.tbr") Then
        Set TE = FSO.OpenTextFile(SYSfolder + "\3pmcfg.tbr", ForReading, False)
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
                    LeerConfig = RST
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

Public Sub WriteTBRLog(TXT As String, PonerFecha As Boolean)

    If FSO.FileExists(AP + "TBRlog.txt") = False Then
        Set TE = FSO.CreateTextFile(AP + "TBRlog.txt", False)
        TE.Close
    End If
    'ver si no es demasiado grande
    If FileLen(AP + "tbrlog.txt") > 100000 Then 'hasta 100 KB aguanto
        'pasarlo a otro archivo y volver a vrearlo
        If FSO.FileExists(AP + "OLDtbrlog.txt") Then FSO.DeleteFile AP + "OLDtbrlog.txt", True
        FSO.MoveFile AP + "tbrlog.txt", AP + "OLDtbrlog.txt"
        Set TE = FSO.CreateTextFile(AP + "TBRlog.txt", False)
        TE.Close
    End If
    'finalmente escribir
    Set TE = FSO.OpenTextFile(AP + "TBRlog.txt", ForAppending, False)
    TE.WriteLine "" 'dejar un renglon en blanco
    If PonerFecha Then
        TE.WriteLine Trim(Str(Date)) + " / " + Trim(Str(Time)) + vbCrLf + TXT
    Else
        TE.WriteLine TXT
    End If
    TE.Close
End Sub


Public Function QuitarNumeroDeTema(TemaFull As String) As String
    Dim TMPtema As String
    TMPtema = TemaFull
    'ver si hay numeros adelante y si hay quitarselos
    Dim NumersoAlInicio As Long
    NumersoAlInicio = 0
    If IsNumeric(Mid(TemaFull, 1, 1)) Then NumersoAlInicio = NumersoAlInicio + 1
    If IsNumeric(Mid(TemaFull, 2, 1)) Then NumersoAlInicio = NumersoAlInicio + 1
    If IsNumeric(Mid(TemaFull, 3, 1)) Then NumersoAlInicio = NumersoAlInicio + 1
    If NumersoAlInicio > 0 Then
        TMPtema = Trim(Right(TemaFull, Len(TemaFull) - 3))
        'ver si quedo con esto
        If Mid(TMPtema, 1, 1) = "-" Or Mid(TMPtema, 1, 1) = "_" Or Mid(TMPtema, 1, 1) = "/" Or Mid(TMPtema, 1, 1) = "@" Then
            TMPtema = Trim(Right(TMPtema, Len(TMPtema) - 1))
        End If
    End If
    QuitarNumeroDeTema = TMPtema
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

