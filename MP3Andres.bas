Attribute VB_Name = "MP3Andres"
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Public CONTADOR As Long
Public CONTADOR2 As Long
Public Contador_Cart As Long
Public CONTADOR2_Cart As Long

Public EsVideo As Boolean 'saber si el tema en ejecucion es video
Public EsKar As Boolean

Public Function TrataEjecutarTema(TEMA As String, Optional ToVIP As Boolean = False) As Long
    
    'devuelve 0 si todo ok
    ' 1 no alcanza el credito (puede ser tambien para VIP)
    '-1 no llega por error!
    ' 2 ya esta ejecutando
    ' 3 si sigue un video
    
    
    On Local Error GoTo ErrTrata
    
    If LCase(Right(TEMA, 3)) = "mp3" Or LCase(Right(TEMA, 3)) = "wma" Then '''Or LCase(Right(temaElegido, 3)) = "mp4" Then
        PideVideo = False
    Else
        PideVideo = True
    End If
                      
    TrataEjecutarTema = -1 'valor predeterminado
              
    'ver si puede pagar lo que pide!!!
    'que joyita papa!!!. Parece que supieras programar
    
    'oct 2007
    'si es de venta de musica esta pasa a ser gratuita (en caso de que lo defina asi)
    'ya que se puede pasar solo algnos segundos de cada canción
    If NOMUSIC Then
        If ShowDemoMusic Then
            'no debe ni mirar precios ni descontar
            GoTo Parte444
        Else
            GoTo FIN443
        End If
    End If
    '--------------------------------------------------------------
    
    If ToVIP And (CREDITOS < PrecNowVIP) Then
        TrataEjecutarTema = 1 'no alcanza el credito para tema VIP
        VerSiTocaPUB
        Exit Function
    End If
    
    If (PideVideo = False And CREDITOS < PrecNowAudio) Or _
        (PideVideo And CREDITOS < PrecNowVideo) Then
        
        TrataEjecutarTema = 1 'no alcanza el credito
        
        VerSiTocaPUB
        
        Exit Function
    End If
    '--------------------------------------------------------------
    'siempre que se ejecute un credito estaremos por debajo de maximo
    SetKeyState vbKeyScrollLock, True
    
    'registrar gasto de plata del usuario!
    Dim YU As Long, DTaa As String
    DTaa = CStr(Year(Date)) + STRceros(Month(Date), 2) + STRceros(Day(Date), 2) + STRceros(Hour(time), 2) + STRceros(Minute(time), 2)
    
    
    'restar lo que corresponde!!!
    'tener en cuenta lo vip !!!
    
    If PideVideo Then
        If ToVIP Then
            VarCreditos -PrecNowVIP
            dwqu _
                "E" + TEMA + "*" + CStr(Round(PrecNowVIP * (PrecioBase / TemasPorCredito), 2)), _
                dwQU_See, _
                DTaa
            
            
        Else
            VarCreditos -PrecNowVideo
            dwqu _
                "E" + TEMA + "*" + CStr(Round(PrecNowVideo * (PrecioBase / TemasPorCredito), 2)), _
                dwQU_See, _
                DTaa
        End If
    Else
        If ToVIP Then
            VarCreditos -PrecNowVIP
            dwqu _
                "E" + TEMA + "*" + CStr(Round(PrecNowVIP * (PrecioBase / TemasPorCredito), 2)), _
                dwQU_See, _
                DTaa
            
            
        Else
    
            VarCreditos -PrecNowAudio
            dwqu _
                "E" + TEMA + "*" + CStr(Round(PrecNowAudio * (PrecioBase / TemasPorCredito), 2)), _
                dwQU_See, _
                DTaa
        End If
    End If

    tERR.Anotar "accy"
    'si esta ejecutando pasa a la lista de reproducción
Parte444:
    'ver si es una cancion gratuita
    With frmIndex.MP3
        Dim Usado As Long, hh As Long
        Usado = -1
        For hh = 0 To 2
            If .IsPlaying(hh) Then Usado = hh
        Next hh
        
        If Usado <> -1 Then 'quiere decir que algo se esta ejecutando
            TrataEjecutarTema = 2
            'pasar a la lista de reproducción
            'el segundo parametro es un tag por ejemplo "PUB" pero en genral para temas comunes es ""
            'el tercer parametro es -1 predeterminado al ultimo de la lista
            ' o puede ser cero para que se ejecute iaaa (NO FUNCIONA AUN VER DLLListaRep
            ' o 1 para proximo, 2 para segundo, 3 para tercero, etc, etc
            If ToVIP Then
                tLST.ListaAdd TEMA, "", 1
            Else
                tLST.ListaAdd TEMA
            End If
            
            tERR.Anotar "accz", TEMA, tLST.GetLastIndex
            CargarProximosTemas
            'graba en reini.tbr los datos que correspondan por si se corta la luz
            CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
            'si esta en uno gratis tengo que sacarlo y seguir
            If .EsGratis(Usado) Then
                If PideVideo Then TrataEjecutarTema = 3 'especial por si es un video
                EMPEZAR_SIGUIENTE 2
            End If
            
        Else
            TrataEjecutarTema = 0 'se larga ya
            If PideVideo Then TrataEjecutarTema = 3 'especial por si es un video
            'NUNCA ENTRARA AQUI si esta en modo de video para los otros si sirve!
            tERR.Anotar "acdc", TEMA
            frmIndex.MP3.EsGratis(IAANext) = False
            CORTAR_TEMA(IAANext) = False 'este tema va entero ya que lo eligio el usuario
            EjecutarTema TEMA, True
        End If
    
    End With
    
FIN443:
    'dejar programada un publicidad si corresponde!
    VerSiTocaPUB

    'estaria bueno que avise que salio como VIP
    'If ToVIP Then frmIndex.lblNOCREDIT.Caption = "Cancion VIP elegida !"

    Exit Function

ErrTrata:
    tERR.AppendLog "TRATA_TM", tERR.ErrToTXT(Err)
    Resume Next
End Function

Public Sub EjecutarTema(TEMA As String, ByRef SumaRanking As Boolean)
    'XXXXX bloquear si esta especificado
    'NOMUSIC ya que solo sera de venta
    'fijarse tambien si
    'ShowDemoMusic es true para mostrar solo un parte
    'ademas las canciones gratuitas tambien ver que hacer
    'tambien las publicidadaes
    
    On Local Error GoTo ErrEjecutarTema

    EstoyEnModoVideoMiniSelDisco = False
    'volver a PasarHoja a su estado original3
    PasarHoja = LeerConfig("PasarHoja", "1")
    
    '*****************************************
    'VER SI SE SUMA AL RANKING O NO
    'si el tema es una publicidad then descuenta de la lista de temas pendientes
    'ademas no va al ranking
    Dim Carp As String
    Carp = txtInLista(TEMA, 99998, "\")
    If LCase(Carp) = "pub" Then
        PUBs.PubsEnLista = PUBs.PubsEnLista - 1
        'tampoco sumar al ranking!!!!
        SumaRanking = False
    End If
    
    If LCase(Carp) = "pubmute" Then SumaRanking = False
    
    tERR.Anotar "003-0001"
    
    '*****************************************
    'VER SI EXISTE EL ARCHIVO
    If fso.FileExists(TEMA) = False Then
        tERR.Anotar "003-0002"
        frmIndex.RollSONG.ReplaceIndex 0, TR.Trad("No se encontro" + vbCrLf + _
            "la seleccion%98%Seleccion se refiere a musica, video o " + _
            "karaoke que se quiso reproducir%99%")
        tERR.Anotar "003-0003"
        EMPEZAR_SIGUIENTE 4
    End If
    
    '*****************************************
    'VER SI ESTA VALIDADO
    If RavI >= 4 Then 'revisar validacion!
        Exit Sub 'no se ejecuta nada
    End If
    
    'ver si es solo expendedor de musica
    If NOMUSIC And (ShowDemoMusic = False) Then
        'salgo de aca!
        Exit Sub
    End If
    
    '*****************************************
    'ACTUALIZAR VISTA DE TEXTOS INFERIORES
    
    Dim p1 As String
    p1 = GetPuestoN(1)
    tERR.Anotar "003-0003b"
    
    If fso.FileExists(PL) Then
        TR.SetVars _
            fso.GetBaseName(p1), _
            fso.GetFolder(fso.GetParentFolderName(p1)).Name
            
        frmIndex.RollSONG.ReplaceIndex 2, TR.Trad("el mas escuchado" + vbCrLf + _
                                      "%01%" + vbCrLf + _
                                      "del disco" + vbCrLf + _
                                      "%02%%98%La variable 1 es la cancion " + _
                                      "y la 2 el disco%99%")
    Else 'si el rank no se inicio no exitse!!!!!!!!
        frmIndex.RollSONG.ReplaceIndex 2, TR.Trad("disfrute" + vbCrLf + "su música%99%")
    End If
    
    tERR.Anotar "003-0004"
     SetKeyState vbKeyCapital, True
    ' Tocar el fichero
    
    ' El valor de cada paso del HScrollPos
    tERR.Anotar "003-0005", TEMA
    TEMA_REPRODUCIENDO = TEMA
    Dim nombreTEMA As String, nombreDISCO As String
    nombreTEMA = fso.GetBaseName(TEMA)
    nombreDISCO = fso.GetBaseName(fso.GetParentFolderName(TEMA))
    TR.SetVars _
        QuitarNumeroDeTema(nombreTEMA), _
        nombreDISCO, _
        PuestoN(tLST.GetElementListaPath(1))
    frmIndex.RollSONG.ReplaceIndex 0, TR.Trad("Estas escuchando" + vbCrLf + _
                                      "%01%" + vbCrLf + _
                                      "del disco" + vbCrLf + "%02%" + vbCrLf + _
                                      "Rank # %03%%99%")
    
    '*****************************************
    'UBICAR LOS CONTROLES SEGUN CORRESPONDA
    
    tERR.Anotar "003-0009"
    If UCase(fso.GetExtensionName(TEMA)) <> "MP3" And _
        UCase(fso.GetExtensionName(TEMA)) <> "WMA" Then '''And UCase(FSO.GetExtensionName(tema)) <> "MP4" Then
        
        EsVideo = True
        tERR.Anotar "003-0010", vidFullScreen, Salida2, HabilitarVUMetro
        'cerrar el protector si estaba activo
        Unload frmProtect
        'acomodar los controles en modo video
        'modo texto pata elegir los discos
        
        'ver si esta en el modo de listas de texto !!
        EstoyEnModoVideoMiniSelDisco = (vidFullScreen = False And Salida2 = False)
         
        If Left(LCase(fso.GetExtensionName(TEMA)), 2) = "mn" Then
            EsKar = True
            'el pick kar se pone visible en el wait ok
            frmIndex.picVideo(IAANext).Visible = False
            frmIndex.picVideo(IAA).Visible = False
            
            frmIndex.picVideo(IAANext).ZOrder 1
            frmIndex.picVideo(IAA).ZOrder 1
    
        Else
            frmIndex.picVideo(IAANext).Visible = True
            frmIndex.picKAR.Visible = False
            frmVIDEO.picKAR_V.Visible = False
            frmIndex.picVideo(IAANext).ZOrder
            frmIndex.picVideo(IAA).Visible = False
        End If
        
        'habilitar pasar las paginas con teclas simples
        'por que en el modo texto la lista no
        'tiene paginas
        tERR.Anotar "003-0027"
        PasarHoja = True
    Else
        EsVideo = False
        EsKar = False
        'acomodar los controles en modo normal
        frmIndex.UpdateVista 'empieza una cancion comun
        
        'los karaokes son indice de reproductor 2
        'esto esta mal !
        'por las dudas agrego otro picVideo y listo!
        
        frmIndex.picVideo(IAANext).Visible = False
        frmIndex.picVideo(IAA).Visible = False
        
        frmIndex.picKAR.Visible = False
        frmVIDEO.picKAR_V.Visible = False
        tERR.Anotar "003-0036"
        'volver a PasarHoja a su estado original
        PasarHoja = LeerConfig("PasarHoja", "1")
    End If
    
    
    '*****************************************
    'SACAR DE LA LISTA DE CANCIONES A EJECUTAR EN EL REINICIO
    
    'Valores de ReIni FULL=tema ejecutando y lista LISTA=solo lista NADA=arranca de cero
    'si corresponde graba en reini.tbr la lista de temas por sis se corta la luz
    'graba en reini.tbr los datos que correspondan por si se corta la luz
    tERR.Anotar "003-0037"
    CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
    tERR.Anotar "003-0038"
    'reiniciar reloj de tiempo sin uso
    frmIndex.Timer1.Interval = 0
    tERR.Anotar "003-0039"
    
    SecSinUso = 0
    'lo pongo al ultimo para que tenga tiempo de cargar el tema encargado
    'si lo pongo a donde estaba pasa un pedazito del tema anterior
    tERR.Anotar "003-0040"
    'ya no se cierra!!!
    'para juan carlos BsAs
    'Unload frmTemasDeDisco
    tERR.Anotar "003-0043", TEMA, nombreTEMA, nombreDISCO
    
    '*****************************************
    'contabilizar para el ranking solo si lo pide

    If SumaRanking Then TOP10 TEMA, nombreTEMA, nombreDISCO
    'mostrar el puesto que esta en el ranking
    tERR.Anotar "003-0044", TEMA
    
    '*****************************************
    'EMPEZAR A CARGAR EN EL COMPONENTE REPRODUCTOR
    
    With frmIndex.MP3
        .FileName(IAANext) = TEMA
        tERR.Anotar "003-0047", EsVideo, Salida2
        If EsVideo Then
            If Salida2 Then
                'ver si es karaoke o video comun
                If Left(LCase(fso.GetExtensionName(TEMA)), 2) = "mn" Then
                    EsKar = True
                    
                    frmIndex.WaitOk TEMA
                    
                    'NO HAY QUE SEGUIR!!!
                    Exit Sub
                Else 'no es karaoke
                    EsKar = False
                    'ESCONDER LAS PUBLICIDADES EN LA SALIDA DE tv!!!!!
                    frmVIDEO.picBigImg.Visible = False
                    
                    tERR.Anotar "003-0048b", IAANext
                    .DoOpenVideo "child", frmVIDEO.picVideo.HWND, 0, 0, _
                        (frmVIDEO.picVideo.Width / 15), (frmVIDEO.picVideo.Height / 15), IAANext
                    frmIndex.picVideo(IAANext).Visible = False
                    frmIndex.picKAR.Visible = False
                    frmVIDEO.picKAR_V.Visible = False
                    frmVIDEO.picVideo.Visible = True
                End If
            Else 'va por el monitor
                'ver si es karaoke o video comun
                If Left(LCase(fso.GetExtensionName(TEMA)), 2) = "mn" Then
                    EsKar = True
                    frmIndex.UpdateVista 'esta por empezar un karaoke en el monitor
    
                    frmIndex.WaitOk TEMA
                    'NO HAY QUE SEGUIR!!!
                    Exit Sub
                Else
                    EsKar = False
                    frmIndex.UpdateVista 'esta por empezar un video en el monitor
                    
                    tERR.Anotar "003-0048", IAANext
                    
                    Dim R2 As Long
                    R2 = .DoOpenVideo("child", frmIndex.picVideo(IAANext).HWND, 0, 0, _
                        (frmIndex.picVideo(IAANext).Width / 15), _
                        (frmIndex.picVideo(IAANext).Height / 15), IAANext)
                    '**************************************************
                    'overlapped me saca como una ventana nueva
                    'popup es como overlapped pero sin barra de titulo
                    '**************************************************
                    
                    Select Case R2
                        '0: ok
                        '1: no existe el archivo
                        '3 al mandar el Mci Open fallo !
                        '4: MCIERR_NO_WINDOW
                        '5: otros errores <> 4 que se presentan al pegar el video a una ventana
                        Case 1, 3, 4
                            tERR.AppendLog "guyaby22", CStr(R2)
                            'PASAR AL QUE SIGUE!
                            EMPEZAR_SIGUIENTE 4
                            Exit Sub '!!!!!!!!!!!!!!!!
                    End Select
                    
                    frmIndex.picVideo(IAANext).Visible = True
                    frmIndex.picKAR.Visible = False
                    frmVIDEO.picKAR_V.Visible = False
                End If
            End If
        Else 'no es un video
            tERR.Anotar "003-0049"
            Dim R As Long
            R = .DoOpen(IAANext)
            Select Case R
                Case 1
                    'ya manejo esto antes!
                Case 2
                    MsgBox TR.Trad("No se ha podido abrir el fichero debido a " + _
                        "un problema existente en Windows. " + vbCrLf + _
                        "Revise que el reproductor multimedia de Windows este " + _
                        "instalado y funcione correctamente." + vbCrLf + _
                        "Notifique a tbrSoft de esto para más detalles%99%")
                    Exit Sub '!!!!!!!!!!!!!!!!
                Case 3 'mci dio error
                    tERR.AppendLog "guyaby"
                    'PASAR AL QUE SIGUE!
                    EMPEZAR_SIGUIENTE 4
                    
                    Exit Sub '!!!!!!!!!!!!!!!!
            End Select
        End If
        
        'apenas se abre lo mido
        '*****************************************
        'DARLE AL PLAY viendo desde y hasta donde
        
        tERR.Anotar "003-0049b", iAlias, NOMUSIC, ShowDemoMusic
        
        If NOMUSIC And ShowDemoMusic Then
            If .LengthInSec(IAANext) >= 50 Then
                .SeekTo 30000, IAANext
                TotalTema(IAANext) = 50
                .HastaTema(IAANext) = 50
                varSecPlay = 30
            Else
                TotalTema(IAANext) = .LengthInSec(IAANext)
                .HastaTema(IAANext) = .LengthInSec(IAANext)
                varSecPlay = 0
            End If
        Else
            TotalTema(IAANext) = .LengthInSec(IAANext)
            .HastaTema(IAANext) = .LengthInSec(IAANext)
        End If
        
        'UpdateHastaTema IAANext se cargo en el doopen
        tERR.Anotar "003-0049c", .HastaTema(IAANext)
        
        .Volumen(IAANext) = 0 'sube si o si en los primeros segundos
        tERR.Anotar "003-0051", IAANext
        
        R = .DoPlay(IAANext)
        
        If R = 1 Then
            EMPEZAR_SIGUIENTE 4 'MsgBox "falla play!"
        End If
            
    End With
    tERR.Anotar "003-0052", HabilitarVUMetro
    
    If HabilitarVUMetro Then frmIndex.VU.CarFantastic = False
    
    Exit Sub
ErrEjecutarTema:
    tERR.AppendLog tERR.ErrToTXT(Err), "MP3Andres.BAS" + ".acpo"
    If (frmIndex.MP3.IsPlaying(0) = False And frmIndex.MP3.IsPlaying(1) = False And frmIndex.MP3.IsPlaying(2)) Then
        EMPEZAR_SIGUIENTE 4
    Else
        Resume Next
    End If
End Sub

Public Function EMPEZAR_SIGUIENTE(DesdeDonde As Long) As Long
        
    ContEmpezSig = ContEmpezSig + 1
    frmIndex.List1.List(17) = "EmpezarSig(" + CStr(DesdeDonde) + ")=? = " + CStr(ContEmpezSig)
    
    tERR.Anotar "EmpSig01", DesdeDonde

    'desde donde indica quien pide comenzar cancion
    '1 es desde una cancion que llego a sus ultimos segundos
    '2 desde la tecla B de cancion siguiente
    '3 al inicio del sistema
    '4 el tema siguiente no existe o dio error paso al otro
    '5 una cancion llego a cero! ver que no sea por que se paso de largo
    'con el FF el momento justo (x seg antes de que termine la anterior) que comienzan
    'las canciones con fade in. Esto implica revisar si algo se esta ejecutando.
    'Puede pasar tambien que se cargue una cancion a la lista en estos segundos
    'restantes que ya paso la busqueda del tema siguiente
    '6 desde un tema al azar NUEVO!!(nov 2006)
    '7 cuando enableNextMusic no esta o sea en el inicio de una cancion!
    
    'la funcion devuelve:
    '1: desdedonde=2, adelanta la finalizacion (tambien para cortar gratuitos)
    '2: desdedonde=5
    '3: habia uno en la lista y se ejecuto (audio)
    '4: habia uno en la lista y se ejecuto (video)
    '5: habia uno en la lista y se ejecuto (otro)
    '6: no hay nada en la lista de espera
    
    
    'puede pasar que la cancion que sigue es un video, empieza ok pero al hacer el endplay
    'la cancion anterior viene de nuevo aca y saca el esvideo !!
    If DesdeDonde <> 5 Then
        EstoyEnModoVideoMiniSelDisco = False
        'volver a PasarHoja a su estado original3
        PasarHoja = LeerConfig("PasarHoja", "1")
        EsVideo = False 'no estamos rep video
        EsKar = False
    End If
    '*******************************************************
    'en caso desde la "B" debo poner como tiempo de finalizacion del actual _
        para que no siga (si esperaria hasta el final para terminar normalmente)
    If DesdeDonde = 2 Then ' Or DesdeDonde = 6 Then
        ThisFade = SegFadeB 'que pase rapido por esta vez, despues se acomoda
        
        'ver si es un karaoke !!!
        If frmIndex.MP3.IsPlaying(2) = False Then
            TotalTema(IAA) = frmIndex.MP3.PositionInSec(IAA) + ThisFade
            'ponen en la variable hastatema lo que dice en totaltema
            UpdateHastaTema IAA 'AQUI SIIII
        Else
            frmIndex.MP3.DoStopKar
        End If
        
        EMPEZAR_SIGUIENTE = 1
        frmIndex.List1.List(17) = "EmpezarSig(" + CStr(DesdeDonde) + ")=1 = " + CStr(ContEmpezSig)
        Exit Function
    End If
    '*******************************************************
    
    If DesdeDonde = 7 Then
        Dim TMP As Long: TMP = IAANext: IAANext = IAA: IAA = TMP
    End If
    
    tERR.Anotar "003-0054b", DesdeDonde, IAA, TotalTema(IAA), ThisFade
    
    'CMP
    'si ya hay uno ejecutandose que seria lo normal antes de que termine
    'el anterior
    
    'todo para depurar en rocchio
    Dim J(7) As Boolean
    J(0) = False
    J(1) = False
    J(2) = False
    J(3) = False
    J(5) = False
    J(6) = False
    J(7) = False
    'el (5) es el acumulador
    If DesdeDonde = 5 Then '5 una cancion llego a cero!
        'ver que no sea por que se paso de largo con el ff
        'ver que el 0 y el 1 esten apagados!
        If frmIndex.MP3.IsPlaying(0) Then J(0) = True: J(5) = True
        frmIndex.List1.List(13) = "IsPlaying(0)=" + CStr(J(0))
        'en algunas pcs el anterior aparentemente da falso!!!!
        'por eso agregue esto!!!
        If frmIndex.MP3.isPlayingClock(0) Then J(1) = True: J(5) = True
        frmIndex.List1.List(14) = "IsPlayingClock(0)=" + CStr(J(1))
        
        If frmIndex.MP3.IsPlaying(1) Then J(2) = True: J(5) = True
        frmIndex.List1.List(15) = "IsPlaying(1)=" + CStr(J(2))
        
        If frmIndex.MP3.isPlayingClock(1) Then J(3) = True: J(5) = True
        frmIndex.List1.List(16) = "IsPlayingClock(1)=" + CStr(J(3))
        
        If frmIndex.MP3.IsPlaying(2) Then J(6) = True: J(5) = True
        If frmIndex.MP3.isPlayingClock(2) Then J(7) = True: J(5) = True
        
    End If
    If J(5) Then 'cualquiera sea verdadero
        EMPEZAR_SIGUIENTE = 2
        frmIndex.List1.List(17) = "EmpezarSig(" + CStr(DesdeDonde) + ")=2 = " + CStr(ContEmpezSig)
        Exit Function
    End If
    
    On Local Error GoTo ErrEmpSig
    tERR.Anotar "003-0054", tLST.GetLastIndex
    With frmIndex
        'generar el endplay si o si
        'si hay algun elemento en la lista ejecutarlo
        If tLST.GetLastIndex > 0 Then
            tERR.Anotar "003-0055"
            .RollSONG.ReplaceIndex 0, TR.Trad("Cargando Proximo Tema...%99%")
            
            Dim TemaDeMatriz As String
            tERR.Anotar "003-0057"
            TemaDeMatriz = tLST.GetElementListaPath(1) 'el proximo que sigue!
            'reacomodar la matriz para quitar el primer elemento y ver si no hay mas
            tERR.Anotar "003-0058"
            If tLST.ListaKillElement = 0 Then
                .RollSONG.ReplaceIndex 1, TR.Trad("No hay " + vbCrLf + _
                    "mas selecciones%99%")
            End If
            
            tERR.Anotar "003-0063"
            frmIndex.MP3.EsGratis(IAANext) = False
            CORTAR_TEMA(IAANext) = False 'este tema va entero ya que lo eligio el usuario
            tERR.Anotar "003-0064"
            '*******************************
            'ver si es audio o video o que es
            Dim SP9() As String, sExt As String
            SP9 = Split(TemaDeMatriz, ".")
            sExt = LCase(SP9(UBound(SP9))) 'extencion del archivo
            Select Case sExt
                Case "mp3", "wma"
                    EMPEZAR_SIGUIENTE = 3
                    frmIndex.List1.List(17) = "EmpezarSig(" + CStr(DesdeDonde) + ")=3 = " + CStr(ContEmpezSig)
                Case "avi", "wmv", "mpg", "dat", "mpeg", "vob", "mn0", "mn1"
                    EMPEZAR_SIGUIENTE = 4
                    frmIndex.List1.List(17) = "EmpezarSig(" + CStr(DesdeDonde) + ")=4 = " + CStr(ContEmpezSig)
                Case Else
                    EMPEZAR_SIGUIENTE = 5
                    frmIndex.List1.List(17) = "EmpezarSig(" + CStr(DesdeDonde) + ")=5 = " + CStr(ContEmpezSig)
            End Select
            'C
            EjecutarTema TemaDeMatriz, True
            '*******************************
            tERR.Anotar "003-0065"
            CargarProximosTemas
            frmIndex.Refresh
        Else 'no hay nada en la lista
            EMPEZAR_SIGUIENTE = 6
            frmIndex.List1.List(17) = "EmpezarSig(" + CStr(DesdeDonde) + ")=6 = " + CStr(ContEmpezSig)
            'frmINDEX.MP3.SongName = "" 'no sirve
            tERR.Anotar "003-0066"
            .Timer1.Interval = 3000
            SecSinUso = 0
            'si no hay temas mostrar la leyenda que lo indica
            'tERR.Anotar "003-0067"
            '.lblTiempoRestante = "Falta: " + "00:00"
            tERR.Anotar "003-0068"
            SetKeyState vbKeyCapital, False
            tERR.Anotar "003-0069"
            .RollSONG.ReplaceIndex 0, TR.Trad("Sin reproduccion actual%99%")
            
            tERR.Anotar "003-0071"
            .RollSONG.ReplaceIndex 1, TR.Trad("No hay" + vbCrLf + _
                "mas selecciones%99%")
            
            TEMA_REPRODUCIENDO = TR.Trad("Sin reproduccion actual%99%")
            tERR.Anotar "003-0075", HabilitarVUMetro
            
            If HabilitarVUMetro Then frmIndex.VU.CarFantastic = True
            
            tERR.Anotar "003-0077"
            'frmIndex.MP3.DoClose IAA
            frmIndex.Refresh
            'CMP cambio a multipista
            'frmIndex.picVideo(IAANext).Visible = False
            'frmIndex.picVideo(IAA).Visible = False
        End If
    End With
    
    Exit Function
ErrEmpSig:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpAnd.B" + ".acpr"
    Resume Next
End Function

Public Sub TOP10(nameARCH As String, nameTEMA As String, nameDISCO As String)
    'On Error GoTo notop
    'ver si existe ranking.tbr
    tERR.Anotar "003-0078", nameARCH
    If fso.FileExists(GPF("rd3_444")) = False Then
        tERR.Anotar "003-0079"
        fso.CreateTextFile GPF("rd3_444"), True
    End If
    tERR.Anotar "003-0080"
    Dim TT As String
    Dim mtxTOP10() As String, Z As Integer
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    Dim Encontrado As Boolean
    Dim PTnuevo As Long 'puntos del elemento nuevo
    Dim DatoNuevoFull As String
    Dim ArchivoNuevo As String
    tERR.Anotar "003-0081"
    Encontrado = False
    'abrir el archivo y ver si ya esta el tema
    tERR.Anotar "003-0082"
    Set TE = fso.OpenTextFile(GPF("rd3_444"), ForReading, False)
    'leerlo cargarlo en matriz y ordenar por mas escuchado
    tERR.Anotar "003-0083"
    Do While Not TE.AtEndOfStream
        'cada linea es "puntos,arch,nombretema,nombredisco"
        tERR.Anotar "003-0084"
        TT = TE.ReadLine
        tERR.Anotar "003-0085", TT
        If TT <> "" Then
            tERR.Anotar "003-0086"
            Z = Z + 1
            tERR.Anotar "003-0087", Z
            ThisPTS = Val(txtInLista(TT, 0, ","))
            ThisArch = txtInLista(TT, 1, ",")
            ThisTEMA = txtInLista(TT, 2, ",")
            ThisDISCO = txtInLista(TT, 3, ",")
            tERR.Anotar "003-0091", ThisDISCO, ThisArch
            ReDim Preserve mtxTOP10(Z)
            'comparar este tema con el elegido actual
            tERR.Anotar "003-0092"
            If UCase(Trim(nameARCH)) = UCase(Trim(ThisArch)) Then
                'sumarle un punto
                tERR.Anotar "003-0093"
                ThisPTS = ThisPTS + 1
                'marcar esta cantidad de puntos como referencai futura para
                'agregar el nuevo dato al ranking
                tERR.Anotar "003-0094"
                PTnuevo = ThisPTS
                tERR.Anotar "003-0095"
                TT = CStr(ThisPTS) + "," + ThisArch + "," + ThisTEMA + "," + ThisDISCO
                tERR.Anotar "003-0096"
                DatoNuevoFull = TT
                tERR.Anotar "003-0097"
                ArchivoNuevo = ThisArch
                tERR.Anotar "003-0098"
                Encontrado = True
            End If
            tERR.Anotar "003-0099"
            mtxTOP10(Z) = TT
        End If
    Loop
    tERR.Anotar "003-0100"
    TE.Close
    'ver si el archivo habia sido votado
    tERR.Anotar "003-0101"
    If Encontrado = False Then
        tERR.Anotar "003-0102"
        TT = "1," + Trim(nameARCH) + "," + Trim(nameTEMA) + "," + Trim(nameDISCO)
        tERR.Anotar "003-0103"
        ReDim Preserve mtxTOP10(Z + 1)
        tERR.Anotar "003-0104"
        mtxTOP10(Z + 1) = TT
        tERR.Anotar "003-0105"
        PTnuevo = 1
        tERR.Anotar "003-0106"
        DatoNuevoFull = TT
        tERR.Anotar "003-0107"
        ArchivoNuevo = nameARCH
    End If
    'cargar todos y sacar la primera columna de las zetas
    tERR.Anotar "003-0108"
    Dim MTXsort() As String
    tERR.Anotar "003-0109"
    Set TE = fso.CreateTextFile(GPF("rd3_444"), True)
    Dim PTactual As Long
    Dim YaSeEscribioDatoNuevo As Boolean
    Dim VarMTX As Long 'variacion del indice de la matriz
    tERR.Anotar "003-0110"
    YaSeEscribioDatoNuevo = False
    VarMTX = 0
    tERR.Anotar "003-0111"
    For mtx = 1 To UBound(mtxTOP10)
        tERR.Anotar "003-0112", mtx
        ReDim Preserve MTXsort(mtx + 1)
        tERR.Anotar "003-0113"
        PTactual = txtInLista(mtxTOP10(mtx), 0, ",")
        tERR.Anotar "003-0114", PTactual, PTnuevo
        If PTactual = PTnuevo And YaSeEscribioDatoNuevo = False Then
            tERR.Anotar "003-0115"
            MTXsort(mtx) = DatoNuevoFull
            tERR.Anotar "003-0116"
            TE.WriteLine MTXsort(mtx)
            tERR.Anotar "003-0117"
            YaSeEscribioDatoNuevo = True
            tERR.Anotar "003-0118"
            mtx = mtx - 1
            tERR.Anotar "003-0119"
            VarMTX = 1
        Else
            tERR.Anotar "003-0120"
            If Trim(UCase(ArchivoNuevo)) = Trim(UCase(txtInLista(mtxTOP10(mtx), 1, ","))) Then
                tERR.Anotar "003-0121"
                VarMTX = 0
                tERR.Anotar "003-0122"
                GoTo sig
            End If
            tERR.Anotar "003-0123"
            MTXsort(mtx + VarMTX) = CStr(PTactual) + "," + _
                txtInLista(mtxTOP10(mtx), 1, ",") + "," + _
                txtInLista(mtxTOP10(mtx), 2, ",") + "," + _
                txtInLista(mtxTOP10(mtx), 3, ",")
            tERR.Anotar "003-0124"
            TE.WriteLine MTXsort(mtx + VarMTX)
        End If
sig:
    Next
    tERR.Anotar "003-0125"
    TE.Close
    Exit Sub
notop:
    MsgBox Err.Description
End Sub

Public Sub SumarContadorCreditos(valorSUMAR As Long)

    'son 4 archivos
    '2 para el contador normal y dos el contador historico
    'todos los archivos deben guardar numero diferentes al real
    'para que no puedan buscar por texto
    
    'cc891.dll; cc892.dll para el contador reiniciable
    'cc893.dll; cc894.dll para el historico
    '-----------------------------------------
    'el historico nunca debe restar.
    'cuando se pone en cero el reiniciable no restar al historico!!!!!!!!!!!
    '-----------------------------------------
    
    Dim TMP1 As Long, TMP2 As Long
    Dim TMP3 As Long, TMP4 As Long
    TMP1 = GetNumberArchCredit(GPF("chdc01"))
    TMP2 = GetNumberArchCredit(GPF("chdc02"))
    TMP3 = GetNumberArchCredit(GPF("chdc03"))
    TMP4 = GetNumberArchCredit(GPF("chdc04"))
    'el tmp1 esta multiplicado por 11 y el 2 por 9 (reiniciable)
    'el tmp3 esta multiplicado por 2 y el 4 por 3 (historico)
    
    'comparar el reiniciable
    Dim Res As Long
    Res = (TMP1 / 11) - (TMP2 / 9)
    Dim NewVal As Long
    Select Case Res
        Case 0
            'todo joia
            NewVal = TMP1 / 11 'podria ser el 2 / 9
        Case Is > 0
            'bajaron el tmp2
            NewVal = TMP1 / 11 'el mayor de los dos
        Case Is < 0
            'bajaron el tmp1
            NewVal = TMP2 / 9 'el mayor de los dos
    End Select
    '-----------------SUMAR!!
    NewVal = NewVal + valorSUMAR
    '----------------
    
    'comparara el historico
    Dim Res2 As Long
    Res2 = TMP3 / 2 - TMP4 / 3
    Dim NewVal2 As Long
    Select Case Res2
        Case 0
            'todo joia
            NewVal2 = TMP3 / 2 'podria ser el 4
        Case Is > 0
            'bajaron el tmp2
            NewVal2 = TMP3 / 2 'el mayor de los dos
        Case Is < 0
            'bajaron el tmp1
            NewVal2 = TMP4 / 3 'el mayor de los dos
    End Select
    '-----------------SUMAR si es que hay sumar!!
    'si es menor que cero esta reiniciando el reiniciable!!
    'si es cero es solo para cargar las variables CONTADOR Y CONTADOR2
    If valorSUMAR > 0 Then
        NewVal2 = NewVal2 + valorSUMAR
    End If
    'escribir los dos reiniciables
    PutNumberArchCredit GPF("chdc01"), NewVal * 11
    PutNumberArchCredit GPF("chdc02"), NewVal * 9
    'escribir los dos historicos
    PutNumberArchCredit GPF("chdc03"), NewVal2 * 2
    PutNumberArchCredit GPF("chdc04"), NewVal2 * 3
       
    CONTADOR = NewVal
    CONTADOR2 = NewVal2
End Sub

Public Sub SumarContadorCarrito(valorSUMAR As Long)

    'son 4 archivos
    '2 para el contador normal y dos el contador historico
    'todos los archivos deben guardar numero diferentes al real
    'para que no puedan buscar por texto
    
    'cc895.dll; cc896.dll para el contador reiniciable
    'cc897.dll; cc898.dll para el historico
    '-----------------------------------------
    'el historico nunca debe restar.
    'cuando se pone en cero el reiniciable no restar al historico!!!!!!!!!!!
    '-----------------------------------------
    
    Dim TMP1 As Long, TMP2 As Long
    Dim TMP3 As Long, TMP4 As Long
    TMP1 = GetNumberArchCredit(GPF("chdc05"))
    TMP2 = GetNumberArchCredit(GPF("chdc06"))
    TMP3 = GetNumberArchCredit(GPF("chdc07"))
    TMP4 = GetNumberArchCredit(GPF("chdc08"))
    'el tmp1 esta multiplicado por 7
    'el 2 por 6 (reiniciable)
    'el tmp3 esta multiplicado por 5
    'el 4 por 4 (historico)
    
    'comparar el reiniciable
    Dim Res As Long
    Res = (TMP1 / 7) - (TMP2 / 6)
    Dim NewVal As Long
    Select Case Res
        Case 0
            'todo joia
            NewVal = TMP1 / 7 'podria ser el 2 / 9
        Case Is > 0
            'bajaron el tmp2
            NewVal = TMP1 / 7 'el mayor de los dos
        Case Is < 0
            'bajaron el tmp1
            NewVal = TMP2 / 6 'el mayor de los dos
    End Select
    '-----------------SUMAR!!
    NewVal = NewVal + valorSUMAR
    '----------------
    
    'comparara el historico
    Dim Res2 As Long
    Res2 = TMP3 / 5 - TMP4 / 4
    Dim NewVal2 As Long
    Select Case Res2
        Case 0
            'todo joia
            NewVal2 = TMP3 / 5 'podria ser el 4
        Case Is > 0
            'bajaron el tmp2
            NewVal2 = TMP3 / 5 'el mayor de los dos
        Case Is < 0
            'bajaron el tmp1
            NewVal2 = TMP4 / 4 'el mayor de los dos
    End Select
    '-----------------SUMAR si es que hay sumar!!
    'si es menor que cero esta reiniciando el reiniciable!!
    'si es cero es solo para cargar las variables CONTADOR Y CONTADOR2
    If valorSUMAR > 0 Then
        NewVal2 = NewVal2 + valorSUMAR
    End If
    'escribir los dos reiniciables
    PutNumberArchCredit GPF("chdc05"), NewVal * 7
    PutNumberArchCredit GPF("chdc06"), NewVal * 6
    'escribir los dos historicos
    PutNumberArchCredit GPF("chdc07"), NewVal2 * 5
    PutNumberArchCredit GPF("chdc08"), NewVal2 * 4
       
    Contador_Cart = NewVal
    CONTADOR2_Cart = NewVal2
End Sub



'leer los datos de algun archivo de coins
Private Function GetNumberArchCredit(Arch As String) As Long
    Dim TE8 As TextStream
    tERR.Anotar "003-0126"
    Dim CONTw As Long
    If fso.FileExists(Arch) Then
        tERR.Anotar "003-0129"
        Set TE8 = fso.OpenTextFile(Arch, ForReading, False)
        tERR.Anotar "003-0130"
        CONTw = Val(TE8.ReadLine)
        tERR.Anotar "003-0131"
        TE8.Close
    Else
        tERR.Anotar "003-0132"
        Set TE8 = fso.CreateTextFile(Arch, True)
        tERR.Anotar "003-0133"
        TE8.WriteLine "0"
        tERR.Anotar "003-0134"
        TE8.Close
        CONTw = 0
    End If
    
    GetNumberArchCredit = CONTw
    
End Function

'escribir los datos del archivos de coins
Private Sub PutNumberArchCredit(Arch As String, Valor As Long)
    Dim TE9 As TextStream
    tERR.Anotar "003-0152"
    Set TE9 = fso.CreateTextFile(Arch, True)
    tERR.Anotar "003-0153"
    TE9.WriteLine CStr(Valor)
    tERR.Anotar "003-0154"
    TE9.Close
    tERR.Anotar "003-0155"
End Sub

Public Function GetPuestoN(nOrden As Long) As String
    'leer ranking.tbr y buscar el tema
    Dim tmpTema As String
    tmpTema = ""
    tERR.Anotar "003-0159b", nOrden
    If fso.FileExists(GPF("rd3_444")) Then
        Dim TE661 As TextStream
        Set TE661 = fso.OpenTextFile(GPF("rd3_444"), ForReading, False)
            If TE661.AtEndOfStream Then GoTo fin661
            tmpTema = TE661.ReadLine
            tmpTema = txtInLista(tmpTema, 1, ",")
'            tERR.Anotar "003-0169"
'            ThisTEMA = txtInLista(TT, 2, ",")
'            tERR.Anotar "003-0170"
'            ThisDISCO = txtInLista(TT, 3, ",")
'            tERR.Anotar "003-0171"
        TE661.Close
        Set TE661 = Nothing
    End If
fin661:
    tERR.Anotar "003-0163b", tmpTema
    GetPuestoN = tmpTema
End Function

Public Function PuestoN(TemaBuscado As String) As String
    'leer ranking.tbr y buscar el tema
    tERR.Anotar "003-0159"
    If fso.FileExists(GPF("rd3_444")) = False Then
        'esto no deberia pasar nunca ya que entra despues de que el tema se carga en el ranking
        tERR.Anotar "003-0160"
        fso.CreateTextFile GPF("rd3_444"), True
        tERR.Anotar "003-0161"
        PuestoN = 1
        tERR.Anotar "003-0162"
        Exit Function
    End If
    
    tERR.Anotar "003-0163"
    Set TE = fso.OpenTextFile(GPF("rd3_444"), ForReading, False)
    tERR.Anotar "003-0164"
    Dim TT As String
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    
    Dim PuestoActual As Long
    PuestoActual = 0
    tERR.Anotar "003-0165"
    'XXXXX cuando lee los origenes en red local se pone muy pesado
    'por mas que uno saque el origen sigue siendo lento porque no salen del ranking
    
    Do While Not TE.AtEndOfStream
        tERR.Anotar "003-0166"
        TT = TE.ReadLine
        tERR.Anotar "003-0167", TT
        ThisPTS = Val(txtInLista(TT, 0, ","))
        tERR.Anotar "003-0168", ThisPTS
        ThisArch = txtInLista(TT, 1, ",")
        tERR.Anotar "003-0169", ThisArch
        ThisTEMA = txtInLista(TT, 2, ",")
        tERR.Anotar "003-0170", ThisTEMA
        ThisDISCO = txtInLista(TT, 3, ",")
        tERR.Anotar "003-0171", ThisDISCO
        'If fso.FileExists(ThisArch) Then
            tERR.Anotar "003-0172"
            PuestoActual = PuestoActual + 1
            tERR.Anotar "003-0173", PuestoActual
            If UCase(ThisArch) = UCase(TemaBuscado) Then
                tERR.Anotar "003-0174"
                PuestoN = Trim(CStr(PuestoActual))
                Exit Function
            End If
        'End If
    Loop
    tERR.Anotar "003-0175"
    TE.Close
    tERR.Anotar "003-0176"
    PuestoN = "000" 'era no rank
End Function
