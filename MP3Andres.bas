Attribute VB_Name = "MP3Andres"
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public CONTADOR As Long
Public CONTADOR2 As Long
Public EsVideo As Boolean 'saber si el tema en ejecucion es video

Public Sub EjecutarTema(tema As String, SumaRanking As Boolean)
    EstoyEnModoVideoMiniSelDisco = False
    'si el tema es una publicidad then descuenta de la lista de temas pendientes
    'ademas no va al ranking
    Dim Carp As String
    Carp = txtInLista(tema, 99998, "\")
    If LCase(Carp) = "pub" Then
        PUBs.PubsEnLista = PUBs.PubsEnLista - 1
        'tampoco sumar al ranking!!!!
        SumaRanking = False
    End If
    
    tERR.Anotar "003-0001"
    If FSO.FileExists(tema) = False Then
        tERR.Anotar "003-0002"
        frmIndex.lblTemaSonando = "No se encontro el tema"
        frmIndex.lblTemaSonando2 = "No se encontro el tema"
        tERR.Anotar "003-0003"
        EMPEZAR_SIGUIENTE
    End If
    tERR.Anotar "003-0004"
     OnOffCAPS vbKeyCapital, True
    ' Tocar el fichero
    On Local Error GoTo ErrEjecutarTema
    ' El valor de cada paso del HScrollPos
    tERR.Anotar "003-0005"
    TEMA_REPRODUCIENDO = tema
    Dim nombreTEMA As String, nombreDISCO As String
    tERR.Anotar "003-0006"
    nombreTEMA = FSO.GetBaseName(tema)
    tERR.Anotar "003-0007"
    nombreDISCO = FSO.GetBaseName(FSO.GetParentFolderName(tema))
    tERR.Anotar "003-0008"
    frmIndex.lblTemaSonando = QuitarNumeroDeTema(nombreTEMA) + " / " + nombreDISCO
    frmIndex.lblTemaSonando2 = QuitarNumeroDeTema(nombreTEMA) + " / " + nombreDISCO
    
    tERR.Anotar "003-0009"
    If UCase(FSO.GetExtensionName(tema)) <> "MP3" And UCase(FSO.GetExtensionName(tema)) <> "WMA" Then '''And UCase(FSO.GetExtensionName(tema)) <> "MP4" Then
        EsVideo = True
        tERR.Anotar "003-0010", vidFullScreen, Salida2, HabilitarVUMetro, Is3pmExclusivo
        'cerrar el protector si estaba activo
        Unload frmProtect
        'acomodar los controles en modo video
        'modo texto pata elegir los discos
        With frmIndex
            'ver si es fullscreen o no!!!!!!!
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            If vidFullScreen Then
                ' si estan inhabilitados siempre o solo para videos
                If Salida2 Then
                    'si esta habilitada la salida 2 no me importa nada, dejo
                    'todo como esta en la salida 1
                    'si lo debere habilitar para que siga cargando creditos
                    GoTo NoLeerOtros
                    
                    'para salida doble!!
                    '-------------------
                    '.WindowState = 0 'vbNormal
                    '.Left = 0
                    '.Top = 0
                    '.Width = frmIndex.Width * 2
                    '.Height = Screen.Height
                    '.Refresh
                    '.picVideo.Left = frmIndex.Width / 2
                    '.picVideo.Width = frmIndex.Width / 2 - 50
                    '.picVideo.Top = 0
                    '.picVideo.Height = Screen.Height
                End If
                    
                If HabilitarVUMetro And Is3pmExclusivo = False Then
                    If NoVumVID Then
                        .picVideo.Top = 0
                        .picVideo.Left = 0
                        .picVideo.Width = Screen.Width
                        .picVideo.Height = Screen.Height
                    Else
                        .picVideo.Top = 0
                        .picVideo.Left = .VU1.AnchoBarra
                        .picVideo.Width = .VU1.Width - (.VU1.AnchoBarra * 2)
                        .picVideo.Height = Screen.Height
                        .VU1.Height = .picVideo.Height
                    End If
                Else
                    .picVideo.Top = 0
                    .picVideo.Left = 0
                    .picVideo.Width = Screen.Width
                    .picVideo.Height = Screen.Height
                    .picVideo.ZOrder
                End If
            Else
                '--------------------------------
                'si es salida de TV no volver!!!!
                If Salida2 Then GoTo NoLeerOtros
                '--------------------------------
                EstoyEnModoVideoMiniSelDisco = True
                'quita el fullscreen!!!!
                '.frDISCOS.Height = .picFondo.Top
                '.VU1.Height = .picFondo.Top
                '!!!!!!
                .frModoVideo.Left = Screen.Width - .frModoVideo.Width
                .frTEMAS.Left = Screen.Width - .frTEMAS.Width
                'en principio los discos ocupan todo
                .frModoVideo.Height = .frDISCOS.Height - .lblModoVideo.Height
                .frModoVideo.Visible = True
                .lblModoVideo.Visible = True
                .VU1.Width = Screen.Width - .frModoVideo.Width
                'tener en cuenta si es exclusivo!!!
                If HabilitarVUMetro And Is3pmExclusivo = False Then
                    .frDISCOS.Width = .VU1.Width - (.VU1.AnchoBarra * 2) - 50
                    .picVideo.Width = .VU1.Width - (.VU1.AnchoBarra * 2)
                    .picVideo.Left = .VU1.AnchoBarra
                Else
                    .frDISCOS.Width = .VU1.Width
                    .picVideo.Width = .VU1.Width
                    .picVideo.Left = 0
                End If
                .picFondoDisco.Top = 0
                .picFondoDisco.Left = 0
                
                .picVideo.Top = 0
                .picVideo.Height = .picFondo.Top
            End If
            
'aqui vengo si es fullscreen y no me importa mover nada
NoLeerOtros:
            'si no hago esto el video no se ve (ya que esta adentro)
            '.picFondoDisco.Height = .frDISCOS.Height
            '.picFondoDisco.Width = .frDISCOS.Width
            
            tERR.Anotar "003-0026"
        End With
        'habilitar pasar las paginas con teclas simples
        'por que en el modo texto la lista no
        'tiene paginas
        tERR.Anotar "003-0027"
        PasarHoja = True
    Else
        EsVideo = False
        'acomodar los controles en modo normal
        With frmIndex
            'quita el fullscreen!!!!
            '.frDISCOS.Height = .picFondo.Top
            .VU1.Height = .picFondo.Top
            '!!!!!!
            tERR.Anotar "003-0028"
            .VU1.Width = Screen.Width
            tERR.Anotar "003-0029"
            If HabilitarVUMetro And Is3pmExclusivo = False Then
                .frDISCOS.Left = .VU1.AnchoBarra + 25 ' .VU1.Width
                tERR.Anotar "003-0030"
                .frDISCOS.Width = .VU1.Width - (.VU1.AnchoBarra * 2) - 50
                '.frDISCOS.Width = Screen.Width - .VU1.Width
                'vu no se mueve         .VU1.Top = 0                '.VU1.Height = Screen.Height
            Else
                'si viene de un video se tiene que ensanchar
                tERR.Anotar "003-0031"
                .frDISCOS.Width = .VU1.Width ' Screen.Width
                .frDISCOS.Left = 0
            End If
            tERR.Anotar "003-0032"
            .picFondoDisco.Height = .frDISCOS.Height
            .picFondoDisco.Width = .frDISCOS.Width
            tERR.Anotar "003-0033"
            .picFondoDisco.Top = 0
            .picFondoDisco.Left = 0
            tERR.Anotar "003-0034"
            .frModoVideo.Visible = False
            .lblModoVideo.Visible = False
            tERR.Anotar "003-0035"
            .frTEMAS.Visible = False
            .lblTEMAS.Visible = False
            .picVideo.Visible = False
        End With
        tERR.Anotar "003-0036"
        'volver a PasarHoja a su estado original
        PasarHoja = LeerConfig("PasarHoja", "1")
    End If
    'Valores de ReIni FULL=tema ejecutando y lista LISTA=solo lista NADA=arranca de cero
    'si corresponde graba en reini.tbr la lista de temas por sis se corta la luz
    'graba en reini.tbr los datos que correspondan por si se corta la luz
    tERR.Anotar "003-0037"
    CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
    tERR.Anotar "003-0038"
    'reiniciar reloj de tiempo sin uso
    frmIndex.Timer1.Interval = 0
    tERR.Anotar "003-0039"
    frmIndex.lblNoUSO = "0"
    SecSinUso = 0
    'lo pongo al ultimo para que tenga tiempo de cargar el tema encargado
    'si lo pongo a donde estaba pasa un pedazito del tema anterior
    tERR.Anotar "003-0040"
    'ya no se cierra!!!
    'para juan carlos BsAs
    'Unload frmTemasDeDisco
    tERR.Anotar "003-0041"
    frmIndex.Refresh
    
    tERR.Anotar "003-0042"
    frmIndex.lblPuesto = "Calculando..."
    frmIndex.lblPuesto2 = "Calculando..."
    tERR.Anotar "003-0043", tema, nombreTEMA, nombreDISCO
    'contabilizar para el ranking solo si lo pide
    If SumaRanking Then TOP10 tema, nombreTEMA, nombreDISCO
    'mostrar el puesto que esta en el ranking
    tERR.Anotar "003-0044", tema
    frmIndex.lblPuesto = "Rank # " + PuestoN(tema)
    frmIndex.lblPuesto2 = "Rank # " + PuestoN(tema)
    
    tERR.Anotar "003-0045"
    With frmIndex.MP3
        tERR.Anotar "003-0046"
        .FileName = tema
        tERR.Anotar "003-0047", EsVideo, Salida2
        If EsVideo Then
            If Salida2 Then
                'ESCONDER LAS PUBLICIDADES EN LA SALIDA DE tv!!!!!
                frmVIDEO.picBigImg.Visible = False
                
                tERR.Anotar "003-0048b"
                .DoOpenVideo "child", frmVIDEO.hwnd, 0, 0, _
                    (frmVIDEO.Width / 15), (frmVIDEO.Height / 15)
                frmIndex.picVideo.Visible = False
            Else
                tERR.Anotar "003-0048"
                .DoOpenVideo "child", frmIndex.picVideo.hwnd, 0, 0, _
                    (frmIndex.picVideo.Width / 15), (frmIndex.picVideo.Height / 15)
                '**************************************************
                'overlapped me saca como una ventana nueva
                'popup es como overlapped pero sin barra de titulo
                '**************************************************
                frmIndex.picVideo.Visible = True
            End If
        Else
            tERR.Anotar "003-0049"
            .DoOpen
        End If
        tERR.Anotar "003-0050"
        'si es un tema al azar usar otro volumen
        If CORTAR_TEMA Then
            .Volumen = VolumenIni2 'el dos es un volumen para temas gratuitos
        Else
            .Volumen = VolumenIni
        End If
        tERR.Anotar "003-0051"
        .DoPlay
    End With
    tERR.Anotar "003-0052"
    If HabilitarVUMetro Then
        If Is3pmExclusivo Then
            frmIndex.VU21.CarFantastic = False
        Else
            frmIndex.VU1.CarFantastic = False
        End If
    End If
    If EsVideo Then
        On Local Error GoTo ErrorPavo
        tERR.Anotar "003-0053"
        'para que tome de nuevo el control del teclado
        frmIndex.SetFocus 'JOYA JOYA!!! en mp3 da error, no usar
    End If
    'para que tome de nuevo el control del teclado
    
    Exit Sub
ErrEjecutarTema:
    tERR.AppendLog tERR.ErrToTXT(Err), "MP3Andres.BAS" + ".acpo"
    'WriteTBRLog "ERROR EN EJECUTAR TEMA. " + frmIndex.MP3.FileName + vbCrLf + _
        "Descripcion: " + Err.Description, True
    If frmIndex.MP3.IsPlaying = False Then EMPEZAR_SIGUIENTE
        
    Exit Sub
    
ErrorPavo:
    'tERR.AppendLog tERR.ErrToTXT(Err), "MpAnd.BAS.SetFocus" + ".acpr"
    'WriteTBRLog "ERROR EN EJECUTAR VIDEO SETFOCUS. ERROR OMITIDO. Descripcion: " + vbCrLf + _
        Err.Description + frmIndex.MP3.FileName, True
    Resume Next
End Sub

Public Sub EMPEZAR_SIGUIENTE()
    On Local Error GoTo ErrEmpSig
    tERR.Anotar "003-0054", UBound(MATRIZ_LISTA)
    With frmIndex
        'generar el endplay si o si
        'si hay algun elemento en la lista ejecutarlo
        If UBound(MATRIZ_LISTA) > 0 Then
            
            tERR.Anotar "003-0055"
            .lblTemaSonando = "Cargando Proximo Tema..."
            .lblTemaSonando2 = "Cargando Proximo Tema..."
            tERR.Anotar "003-0056"
            .lblTemaSonando.Refresh
            .lblTemaSonando2.Refresh
            Dim TemaDeMatriz As String
            tERR.Anotar "003-0057"
            TemaDeMatriz = txtInLista(MATRIZ_LISTA(1), 0, ",")
            'reacomodar la matriz para quitar el primer elemento
            Dim c As Long
            tERR.Anotar "003-0058"
            For c = 1 To UBound(MATRIZ_LISTA)
                tERR.Anotar "003-0059"
                If c < UBound(MATRIZ_LISTA) Then
                    'cuando sea cualquiera menos el ultimo
                    tERR.Anotar "003-0060"
                    MATRIZ_LISTA(c) = MATRIZ_LISTA(c + 1)
                Else
                    'cuando sea el ultimo
                    'redefinir la matriz con un indice menos
                    tERR.Anotar "003-0061"
                    ReDim Preserve MATRIZ_LISTA(c - 1)
                    '.lstProximos.Clear
                    '.lstProximos.AddItem "No hay próximo tema"
                    tERR.Anotar "003-0062"
                    .lstProximos = "No hay próximo tema"
                End If
            Next
            tERR.Anotar "003-0063"
            CORTAR_TEMA = False 'este tema va entero ya que lo eligio el usuario
            tERR.Anotar "003-0064"
            EjecutarTema TemaDeMatriz, True
            tERR.Anotar "003-0065"
            CargarProximosTemas
        Else
            .lblREP = ""
            'frmINDEX.MP3.SongName = "" 'no sirve
            tERR.Anotar "003-0066"
            .Timer1.Interval = 3000
            SecSinUso = 0
            'si no hay temas mostrar la leyenda que lo indica
            tERR.Anotar "003-0067"
            .lblTiempoRestante = "Falta: " + "00:00"
            tERR.Anotar "003-0068"
            OnOffCAPS vbKeyCapital, False
            tERR.Anotar "003-0069"
            .lblTemaSonando = "Sin reproduccion actual"
            .lblTemaSonando2 = "Sin reproduccion actual"
            
            tERR.Anotar "003-0070"
            .lblPuesto = "No Rank"
            .lblPuesto2 = "No Rank"
            '.lstProximos.Clear
            '.lstProximos.AddItem "No hay próximo tema"
            tERR.Anotar "003-0071"
            .lstProximos = "No hay próximo tema"
            tERR.Anotar "003-0072"
            .lblTiempoRestante = "Falta: " + "00:00"
            tERR.Anotar "003-0073"
            
            
            TEMA_REPRODUCIENDO = "Sin reproduccion actual"
            tERR.Anotar "003-0075"
            If HabilitarVUMetro Then
                If Is3pmExclusivo Then
                    frmIndex.VU21.CarFantastic = True
                Else
                    frmIndex.VU1.CarFantastic = True
                End If
            End If
            tERR.Anotar "003-0076"
            EsVideo = False 'no estamos rep video
            tERR.Anotar "003-0077"
            frmIndex.MP3.DoClose
            frmIndex.Refresh
            frmIndex.picVideo.Visible = False
        End If
    End With
    
    Exit Sub
ErrEmpSig:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpAnd.B" + ".acpr"
    Resume Next
End Sub

Public Sub TOP10(nameARCH As String, nameTEMA As String, nameDISCO As String)
    'On Error GoTo notop
    'ver si existe ranking.tbr
    tERR.Anotar "003-0078", nameARCH
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        tERR.Anotar "003-0079"
        FSO.CreateTextFile AP + "ranking.tbr", True
    End If
    tERR.Anotar "003-0080"
    Dim TT As String
    Dim mtxTOP10() As String, z As Integer
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
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    'leerlo cargarlo en matriz y ordenar por mas escuchado
    tERR.Anotar "003-0083"
    Do While Not TE.AtEndOfStream
        'cada linea es "puntos,arch,nombretema,nombredisco"
        tERR.Anotar "003-0084"
        TT = TE.ReadLine
        tERR.Anotar "003-0085", TT
        If TT <> "" Then
            tERR.Anotar "003-0086"
            z = z + 1
            tERR.Anotar "003-0087", z
            ThisPTS = Val(txtInLista(TT, 0, ","))
            tERR.Anotar "003-0088", ThisPTS
            ThisArch = txtInLista(TT, 1, ",")
            tERR.Anotar "003-0089", ThisArch
            ThisTEMA = txtInLista(TT, 2, ",")
            tERR.Anotar "003-0090", ThisTEMA
            ThisDISCO = txtInLista(TT, 3, ",")
            tERR.Anotar "003-0091", ThisDISCO
            ReDim Preserve mtxTOP10(z)
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
                TT = Str(ThisPTS) + "," + ThisArch + "," + ThisTEMA + "," + ThisDISCO
                tERR.Anotar "003-0096"
                DatoNuevoFull = TT
                tERR.Anotar "003-0097"
                ArchivoNuevo = ThisArch
                tERR.Anotar "003-0098"
                Encontrado = True
            End If
            tERR.Anotar "003-0099"
            mtxTOP10(z) = TT
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
        ReDim Preserve mtxTOP10(z + 1)
        tERR.Anotar "003-0104"
        mtxTOP10(z + 1) = TT
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
    Set TE = FSO.CreateTextFile(AP + "ranking.tbr", True)
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
    TMP1 = GetNumberArchCredit(SYSfolder + "cc891.dll")
    TMP2 = GetNumberArchCredit(SYSfolder + "cc892.dll")
    TMP3 = GetNumberArchCredit(SYSfolder + "cc893.dll")
    TMP4 = GetNumberArchCredit(SYSfolder + "cc894.dll")
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
    PutNumberArchCredit SYSfolder + "cc891.dll", NewVal * 11
    PutNumberArchCredit SYSfolder + "cc892.dll", NewVal * 9
    'escribir los dos historicos
    PutNumberArchCredit SYSfolder + "cc893.dll", NewVal2 * 2
    PutNumberArchCredit SYSfolder + "cc894.dll", NewVal2 * 3
       
    CONTADOR = NewVal
    CONTADOR2 = NewVal2
End Sub
'leer los datos de algun archivo de coins
Private Function GetNumberArchCredit(Arch As String) As Long
    Dim TE8 As TextStream
    tERR.Anotar "003-0126"
    Dim CONTw As Long
    If FSO.FileExists(Arch) Then
        tERR.Anotar "003-0129"
        Set TE8 = FSO.OpenTextFile(Arch, ForReading, False)
        tERR.Anotar "003-0130"
        CONTw = Val(TE8.ReadLine)
        tERR.Anotar "003-0131"
        TE8.Close
    Else
        tERR.Anotar "003-0132"
        Set TE8 = FSO.CreateTextFile(Arch, True)
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
    Set TE9 = FSO.CreateTextFile(Arch, True)
    tERR.Anotar "003-0153"
    TE9.WriteLine CStr(Valor)
    tERR.Anotar "003-0154"
    TE9.Close
    tERR.Anotar "003-0155"
End Sub

Public Function PuestoN(TemaBuscado As String) As String
    'leer ranking.tbr y buscar el tema
    tERR.Anotar "003-0159"
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        'esto no deberia pasar nunca ya que entra despues de que el tema se carga en el ranking
        tERR.Anotar "003-0160"
        FSO.CreateTextFile AP + "ranking.tbr", True
        tERR.Anotar "003-0161"
        PuestoN = 1
        tERR.Anotar "003-0162"
        Exit Function
    End If
    tERR.Anotar "003-0163"
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    tERR.Anotar "003-0164"
    Dim TT As String
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    
    Dim PuestoActual As Long
    PuestoActual = 0
    tERR.Anotar "003-0165"
    Do While Not TE.AtEndOfStream
        tERR.Anotar "003-0166"
        TT = TE.ReadLine
        tERR.Anotar "003-0167"
        ThisPTS = Val(txtInLista(TT, 0, ","))
        tERR.Anotar "003-0168"
        ThisArch = txtInLista(TT, 1, ",")
        tERR.Anotar "003-0169"
        ThisTEMA = txtInLista(TT, 2, ",")
        tERR.Anotar "003-0170"
        ThisDISCO = txtInLista(TT, 3, ",")
        tERR.Anotar "003-0171"
        If FSO.FileExists(ThisArch) Then
            tERR.Anotar "003-0172"
            PuestoActual = PuestoActual + 1
            tERR.Anotar "003-0173"
            If UCase(ThisArch) = UCase(TemaBuscado) Then
                tERR.Anotar "003-0174"
                PuestoN = Trim(Str(PuestoActual))
                Exit Function
            End If
        End If
    Loop
    tERR.Anotar "003-0175"
    TE.Close
    tERR.Anotar "003-0176"
    PuestoN = "000" 'era no rank
End Function
