Attribute VB_Name = "MP3Andres"
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public CONTADOR As Long
Public EsVideo As Boolean 'saber si el tema en ejecucion es video

Public Sub EjecutarTema(tema As String, SumaRanking As Boolean)
    'si el tema es una publicidad then descuenta de la lista de temas pendientes
    'ademas no va al ranking
    Dim Carp As String
    Carp = txtInLista(tema, 99998, "\")
    If LCase(Carp) = "pub" Then
        PUBs.PubsEnLista = PUBs.PubsEnLista - 1
        'tampoco sumar al ranking!!!!
        SumaRanking = False
    End If
    
    LineaError = "003-0001"
    If FSO.FileExists(tema) = False Then
        LineaError = "003-0002"
        frmIndex.lblTemaSonando = "No se encontro el tema"
        LineaError = "003-0003"
        EMPEZAR_SIGUIENTE
    End If
    LineaError = "003-0004"
     OnOffCAPS vbKeyCapital, True
    ' Tocar el fichero
    On Local Error GoTo ErrEjecutarTema
    ' El valor de cada paso del HScrollPos
    LineaError = "003-0005"
    TEMA_REPRODUCIENDO = tema
    Dim nombreTEMA As String, nombreDISCO As String
    LineaError = "003-0006"
    nombreTEMA = FSO.GetBaseName(tema)
    LineaError = "003-0007"
    nombreDISCO = FSO.GetBaseName(FSO.GetParentFolderName(tema))
    LineaError = "003-0008"
    frmIndex.lblTemaSonando = QuitarNumeroDeTema(nombreTEMA) + " / " + nombreDISCO
    
    LineaError = "003-0009"
    If UCase(FSO.GetExtensionName(tema)) <> "MP3" Then
        EsVideo = True
        LineaError = "003-0010"
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
                    
                If HabilitarVUMetro Then
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
                End If
            Else
                '--------------------------------
                'si es salida de TV no volver!!!!
                If Salida2 Then GoTo NoLeerOtros
                '--------------------------------
                
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
                If HabilitarVUMetro Then
                    .frDISCOS.Width = .VU1.Width - (.VU1.AnchoBarra * 2) - 50
                Else
                    .frDISCOS.Width = .VU1.Width
                End If
                .picFondoDisco.Top = 0
                .picFondoDisco.Left = 0
                
                .picVideo.Top = 0
                .picVideo.Left = .VU1.AnchoBarra
                .picVideo.Width = .VU1.Width - (.VU1.AnchoBarra * 2)
                .picVideo.Height = .picFondo.Top
            End If
            
            
            
'aqui vengo si es fullscreen y no me importa mover nada
NoLeerOtros:
            'si no hago esto el video no se ve (ya que esta adentro)
            '.picFondoDisco.Height = .frDISCOS.Height
            '.picFondoDisco.Width = .frDISCOS.Width
            
            LineaError = "003-0026"
        End With
        'habilitar pasar las paginas con teclas simples
        'por que en el modo texto la lista no
        'tiene paginas
        LineaError = "003-0027"
        PasarHoja = True
    Else
        EsVideo = False
        'acomodar los controles en modo normal
        With frmIndex
            'quita el fullscreen!!!!
            '.frDISCOS.Height = .picFondo.Top
            .VU1.Height = .picFondo.Top
            '!!!!!!
            LineaError = "003-0028"
            .VU1.Width = Screen.Width
            LineaError = "003-0029"
            If HabilitarVUMetro Then
                .frDISCOS.Left = .VU1.AnchoBarra + 25 ' .VU1.Width
                LineaError = "003-0030"
                .frDISCOS.Width = .VU1.Width - (.VU1.AnchoBarra * 2) - 50
                '.frDISCOS.Width = Screen.Width - .VU1.Width
                'vu no se mueve         .VU1.Top = 0                '.VU1.Height = Screen.Height
            Else
                'si viene de un video se tiene que ensanchar
                LineaError = "003-0031"
                .frDISCOS.Width = .VU1.Width ' Screen.Width
                .frDISCOS.Left = 0
            End If
            LineaError = "003-0032"
            .picFondoDisco.Height = .frDISCOS.Height
            .picFondoDisco.Width = .frDISCOS.Width
            LineaError = "003-0033"
            .picFondoDisco.Top = 0
            .picFondoDisco.Left = 0
            LineaError = "003-0034"
            .frModoVideo.Visible = False
            .lblModoVideo.Visible = False
            LineaError = "003-0035"
            .frTEMAS.Visible = False
            .lblTEMAS.Visible = False
            .picVideo.Visible = False
        End With
        LineaError = "003-0036"
        'volver a PasarHoja a su estado original
        PasarHoja = LeerConfig("PasarHoja", "1")
    End If
    'Valores de ReIni FULL=tema ejecutando y lista LISTA=solo lista NADA=arranca de cero
    'si corresponde graba en reini.tbr la lista de temas por sis se corta la luz
    'graba en reini.tbr los datos que correspondan por si se corta la luz
    LineaError = "003-0037"
    CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
    LineaError = "003-0038"
    'reiniciar reloj de tiempo sin uso
    frmIndex.Timer1.Interval = 0
    LineaError = "003-0039"
    frmIndex.lblNoUSO = "0"
    SecSinUso = 0
    'lo pongo al ultimo para que tenga tiempo de cargar el tema encargado
    'si lo pongo a donde estaba pasa un pedazito del tema anterior
    LineaError = "003-0040"
    
    Unload frmTemasDeDisco
    LineaError = "003-0041"
    frmIndex.Refresh
    
    LineaError = "003-0042"
    frmIndex.lblPuesto = "Calculando..."
    LineaError = "003-0043"
    'contabilizar para el ranking solo si lo pide
    If SumaRanking Then TOP10 tema, nombreTEMA, nombreDISCO
    'mostrar el puesto que esta en el ranking
    LineaError = "003-0044"
    frmIndex.lblPuesto = "Rank # " + PuestoN(tema)
    
    LineaError = "003-0045"
    With frmIndex.MP3
        LineaError = "003-0046"
        .FileName = tema
        LineaError = "003-0047"
        If EsVideo Then
            If Salida2 Then
                'ESCONDER LAS PUBLICIDADES EN LA SALIDA DE tv!!!!!
                frmVIDEO.picBigImg.Visible = False
                
                LineaError = "003-0048b"
                .DoOpenVideo "child", frmVIDEO.hWnd, 0, 0, _
                    (frmVIDEO.Width / 15), (frmVIDEO.Height / 15)
                frmIndex.picVideo.Visible = False
            Else
                LineaError = "003-0048"
                .DoOpenVideo "child", frmIndex.picVideo.hWnd, 0, 0, _
                    (frmIndex.picVideo.Width / 15), (frmIndex.picVideo.Height / 15)
                frmIndex.picVideo.Visible = True
            End If
        Else
            LineaError = "003-0049"
            .DoOpen
        End If
        LineaError = "003-0050"
        .Volumen = VolumenIni
        LineaError = "003-0051"
        .DoPlay
    End With
    LineaError = "003-0052"
    If HabilitarVUMetro Then frmIndex.VU1.CarFantastic = False
    If EsVideo Then
        LineaError = "003-0053"
        'para que tome de nuevo el control del teclado
        frmIndex.SetFocus 'JOYA JOYA!!! en mp3 da error, no usar
    End If
    'para qyue tome de nuevo el control del teclado
    Exit Sub
ErrEjecutarTema:
    WriteTBRLog "ERROR EN EJECUTAR TEMA. " + frmIndex.MP3.FileName + vbCrLf + _
        "Descripcion: " + Err.Description, True
    If frmIndex.MP3.IsPlaying = False Then EMPEZAR_SIGUIENTE
End Sub

Public Sub EMPEZAR_SIGUIENTE()
    LineaError = "003-0054"
    With frmIndex
        'generar el endplay si o si
        'si hay algun elemento en la lista ejecutarlo
        If UBound(MATRIZ_LISTA) > 0 Then
            
            LineaError = "003-0055"
            .lblTemaSonando = "Cargando Proximo Tema..."
            LineaError = "003-0056"
            .lblTemaSonando.Refresh
            Dim TemaDeMatriz As String
            LineaError = "003-0057"
            TemaDeMatriz = txtInLista(MATRIZ_LISTA(1), 0, ",")
            'reacomodar la matriz para quitar el primer elemento
            Dim c As Long
            LineaError = "003-0058"
            For c = 1 To UBound(MATRIZ_LISTA)
                LineaError = "003-0059"
                If c < UBound(MATRIZ_LISTA) Then
                    'cuando sea cualquiera menos el ultimo
                    LineaError = "003-0060"
                    MATRIZ_LISTA(c) = MATRIZ_LISTA(c + 1)
                Else
                    'cuando sea el ultimo
                    'redefinir la matriz con un indice menos
                    LineaError = "003-0061"
                    ReDim Preserve MATRIZ_LISTA(c - 1)
                    '.lstProximos.Clear
                    '.lstProximos.AddItem "No hay próximo tema"
                    LineaError = "003-0062"
                    .lstProximos = "No hay próximo tema"
                End If
            Next
            LineaError = "003-0063"
            CORTAR_TEMA = False 'este tema va entero ya que lo eligio el usuario
            LineaError = "003-0064"
            EjecutarTema TemaDeMatriz, True
            LineaError = "003-0065"
            CargarProximosTemas
        Else
            'frmINDEX.MP3.SongName = "" 'no sirve
            LineaError = "003-0066"
            .Timer1.Interval = 10000
            SecSinUso = 0
            'si no hay temas mostrar la leyenda que lo indica
            LineaError = "003-0067"
            .lblTiempoRestante = "FALTA: " + "00:00"
            LineaError = "003-0068"
            OnOffCAPS vbKeyCapital, False
            LineaError = "003-0069"
            .lblTemaSonando = "Sin reproduccion actual"
            
            LineaError = "003-0070"
            .lblPuesto = "No Rank"
            '.lstProximos.Clear
            '.lstProximos.AddItem "No hay próximo tema"
            LineaError = "003-0071"
            .lstProximos = "No hay próximo tema"
            LineaError = "003-0072"
            .lblTiempoRestante = "FALTA: " + "00:00"
            LineaError = "003-0073"
            .LBLpORCtEMA.Width = .lblTemaSonando.Width
            LineaError = "003-0074"
            TEMA_REPRODUCIENDO = "Sin reproduccion actual"
            LineaError = "003-0075"
            If HabilitarVUMetro Then frmIndex.VU1.CarFantastic = True
            LineaError = "003-0076"
            EsVideo = False 'no estamos rep video
            LineaError = "003-0077"
            frmIndex.MP3.DoClose
            frmIndex.Refresh
            frmIndex.picVideo.Visible = False
        End If
    End With
End Sub

Public Sub TOP10(nameARCH As String, nameTEMA As String, nameDISCO As String)
    'On Error GoTo notop
    'ver si existe ranking.tbr
    LineaError = "003-0078"
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        LineaError = "003-0079"
        FSO.CreateTextFile AP + "ranking.tbr", True
    End If
    LineaError = "003-0080"
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
    LineaError = "003-0081"
    Encontrado = False
    'abrir el archivo y ver si ya esta el tema
    LineaError = "003-0082"
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    'leerlo cargarlo en matriz y ordenar por mas escuchado
    LineaError = "003-0083"
    Do While Not TE.AtEndOfStream
        'cada linea es "puntos,arch,nombretema,nombredisco"
        LineaError = "003-0084"
        TT = TE.ReadLine
        LineaError = "003-0085"
        If TT <> "" Then
            LineaError = "003-0086"
            z = z + 1
            LineaError = "003-0087"
            ThisPTS = Val(txtInLista(TT, 0, ","))
            LineaError = "003-0088"
            ThisArch = txtInLista(TT, 1, ",")
            LineaError = "003-0089"
            ThisTEMA = txtInLista(TT, 2, ",")
            LineaError = "003-0090"
            ThisDISCO = txtInLista(TT, 3, ",")
            LineaError = "003-0091"
            ReDim Preserve mtxTOP10(z)
            'comparar este tema con el elegido actual
            LineaError = "003-0092"
            If UCase(Trim(nameARCH)) = UCase(Trim(ThisArch)) Then
                'sumarle un punto
                LineaError = "003-0093"
                ThisPTS = ThisPTS + 1
                'marcar esta cantidad de puntos como referencai futura para
                'agregar el nuevo dato al ranking
                LineaError = "003-0094"
                PTnuevo = ThisPTS
                LineaError = "003-0095"
                TT = Str(ThisPTS) + "," + ThisArch + "," + ThisTEMA + "," + ThisDISCO
                LineaError = "003-0096"
                DatoNuevoFull = TT
                LineaError = "003-0097"
                ArchivoNuevo = ThisArch
                LineaError = "003-0098"
                Encontrado = True
            End If
            LineaError = "003-0099"
            mtxTOP10(z) = TT
        End If
    Loop
    LineaError = "003-0100"
    TE.Close
    'ver si el archivo habia sido votado
    LineaError = "003-0101"
    If Encontrado = False Then
        LineaError = "003-0102"
        TT = "1," + Trim(nameARCH) + "," + Trim(nameTEMA) + "," + Trim(nameDISCO)
        LineaError = "003-0103"
        ReDim Preserve mtxTOP10(z + 1)
        LineaError = "003-0104"
        mtxTOP10(z + 1) = TT
        LineaError = "003-0105"
        PTnuevo = 1
        LineaError = "003-0106"
        DatoNuevoFull = TT
        LineaError = "003-0107"
        ArchivoNuevo = nameARCH
    End If
    'cargar todos y sacar la primera columna de las zetas
    LineaError = "003-0108"
    Dim MTXsort() As String
    LineaError = "003-0109"
    Set TE = FSO.CreateTextFile(AP + "ranking.tbr", True)
    Dim PTactual As Long
    Dim YaSeEscribioDatoNuevo As Boolean
    Dim VarMTX As Long 'variacion del indice de la matriz
    LineaError = "003-0110"
    YaSeEscribioDatoNuevo = False
    VarMTX = 0
    LineaError = "003-0111"
    For mtx = 1 To UBound(mtxTOP10)
        LineaError = "003-0112"
        ReDim Preserve MTXsort(mtx + 1)
        LineaError = "003-0113"
        PTactual = txtInLista(mtxTOP10(mtx), 0, ",")
        LineaError = "003-0114"
        If PTactual = PTnuevo And YaSeEscribioDatoNuevo = False Then
            LineaError = "003-0115"
            MTXsort(mtx) = DatoNuevoFull
            LineaError = "003-0116"
            TE.WriteLine MTXsort(mtx)
            LineaError = "003-0117"
            YaSeEscribioDatoNuevo = True
            LineaError = "003-0118"
            mtx = mtx - 1
            LineaError = "003-0119"
            VarMTX = 1
        Else
            LineaError = "003-0120"
            If Trim(UCase(ArchivoNuevo)) = Trim(UCase(txtInLista(mtxTOP10(mtx), 1, ","))) Then
                LineaError = "003-0121"
                VarMTX = 0
                LineaError = "003-0122"
                GoTo SIG
            End If
            LineaError = "003-0123"
            MTXsort(mtx + VarMTX) = CStr(PTactual) + "," + _
                txtInLista(mtxTOP10(mtx), 1, ",") + "," + _
                txtInLista(mtxTOP10(mtx), 2, ",") + "," + _
                txtInLista(mtxTOP10(mtx), 3, ",")
            LineaError = "003-0124"
            TE.WriteLine MTXsort(mtx + VarMTX)
        End If
SIG:
    Next
    LineaError = "003-0125"
    TE.Close
    Exit Sub
notop:
    MsgBox Err.Description
End Sub

Public Sub SumarContadorCreditos(valorSUMAR As Long)
    LineaError = "003-0126"
    Dim ARCHcont As String
    'ver el valor en win
    LineaError = "003-0127"
    ARCHcont = WINfolder + "nnr.dll"
    LineaError = "003-0128"
    If FSO.FileExists(ARCHcont) Then
        LineaError = "003-0129"
        Set TE = FSO.OpenTextFile(ARCHcont, ForReading, False)
        Dim CONTw As Long
        LineaError = "003-0130"
        CONTw = Val(TE.ReadLine)
        LineaError = "003-0131"
        TE.Close
    Else
        LineaError = "003-0132"
        Set TE = FSO.CreateTextFile(ARCHcont, True)
        LineaError = "003-0133"
        TE.WriteLine "0"
        LineaError = "003-0134"
        TE.Close
        CONTw = 0
    End If
    LineaError = "003-0135"
    'ver el valor en sys
    ARCHcont = SYSfolder + "\nnr.dll"
    LineaError = "003-0136"
    If FSO.FileExists(ARCHcont) Then
        LineaError = "003-0137"
        Set TE = FSO.OpenTextFile(ARCHcont, ForReading, False)
        LineaError = "003-0138"
        Dim CONTs As Long
        CONTs = Val(TE.ReadLine)
        LineaError = "003-0139"
        TE.Close
    Else
        LineaError = "003-0140"
        Set TE = FSO.CreateTextFile(ARCHcont, True)
        LineaError = "003-0141"
        TE.WriteLine "0"
        LineaError = "003-0142"
        TE.Close
        LineaError = "003-0143"
        CONTs = 0
    End If
    LineaError = "003-0144"
    Dim ContFinal As Long
    If CONTw <> CONTs Then
        LineaError = "003-0145"
        If CONTw > CONTs Then
            LineaError = "003-0146"
            ContFinal = CONTw
        Else
            LineaError = "003-0147"
            ContFinal = CONTs
        End If
    Else
        LineaError = "003-0148"
        'aqui vale cualquiera de los dos por que son iguales
        ContFinal = CONTs
    End If
    LineaError = "003-0149"
    'sumar lo que corresponde
    ContFinal = ContFinal + valorSUMAR
    'asignarlo a la variable global
    LineaError = "003-0150"
    CONTADOR = ContFinal
    LineaError = "003-0151"
    'actualizar el valor en win
    ARCHcont = WINfolder + "nnr.dll"
    LineaError = "003-0152"
    Set TE = FSO.CreateTextFile(ARCHcont, True)
    LineaError = "003-0153"
    TE.WriteLine Trim(Str(CONTADOR))
    LineaError = "003-0154"
    TE.Close
    LineaError = "003-0155"
    'actualizar el valor en sys
    ARCHcont = SYSfolder + "\nnr.dll"
    LineaError = "003-0156"
    Set TE = FSO.CreateTextFile(ARCHcont, True)
    LineaError = "003-0157"
    TE.WriteLine Trim(Str(CONTADOR))
    LineaError = "003-0158"
    TE.Close
    
End Sub

Public Function PuestoN(TemaBuscado As String) As String
    'leer ranking.tbr y buscar el tema
    LineaError = "003-0159"
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        'esto no deberia pasar nunca ya que entra despues de que el tema se carga en el ranking
        LineaError = "003-0160"
        FSO.CreateTextFile AP + "ranking.tbr", True
        LineaError = "003-0161"
        PuestoN = 1
        LineaError = "003-0162"
        Exit Function
    End If
    LineaError = "003-0163"
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    LineaError = "003-0164"
    Dim TT As String
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    
    Dim PuestoActual As Long
    PuestoActual = 0
    LineaError = "003-0165"
    Do While Not TE.AtEndOfStream
        LineaError = "003-0166"
        TT = TE.ReadLine
        LineaError = "003-0167"
        ThisPTS = Val(txtInLista(TT, 0, ","))
        LineaError = "003-0168"
        ThisArch = txtInLista(TT, 1, ",")
        LineaError = "003-0169"
        ThisTEMA = txtInLista(TT, 2, ",")
        LineaError = "003-0170"
        ThisDISCO = txtInLista(TT, 3, ",")
        LineaError = "003-0171"
        If FSO.FileExists(ThisArch) Then
            LineaError = "003-0172"
            PuestoActual = PuestoActual + 1
            LineaError = "003-0173"
            If UCase(ThisArch) = UCase(TemaBuscado) Then
                LineaError = "003-0174"
                PuestoN = Trim(Str(PuestoActual))
                Exit Function
            End If
        End If
    Loop
    LineaError = "003-0175"
    TE.Close
    LineaError = "003-0176"
    PuestoN = "No Rank"
End Function
