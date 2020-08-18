Attribute VB_Name = "MP3Andres"
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public CONTADOR As Long
Public EsVideo As Boolean 'saber si el tema en ejecucion es video

Public Sub EjecutarTema(tema As String, SumaRanking As Boolean)
    If FSO.FileExists(tema) = False Then
        frmIndex.lblTemaSonando = "No se encontro el tema"
        EMPEZAR_SIGUIENTE
    End If
     OnOffCAPS vbKeyCapital, True
    ' Tocar el fichero
    On Local Error GoTo ErrEjecutarTema
    ' El valor de cada paso del HScrollPos
    TEMA_REPRODUCIENDO = tema
    Dim nombreTEMA As String, nombreDISCO As String
    nombreTEMA = FSO.GetBaseName(tema)
    nombreDISCO = FSO.GetBaseName(FSO.GetParentFolderName(tema))
    frmIndex.lblTemaSonando = QuitarNumeroDeTema(nombreTEMA) + " / " + nombreDISCO
    
    If UCase(FSO.GetExtensionName(tema)) <> "MP3" Then
        EsVideo = True
        'cerrar el protector si estaba activo
        Unload frmProtect
        'acomodar los controles en modo video
        'modo texto pata elegir los discos
        With frmIndex
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
                .frDISCOS.Width = .VU1.Width ' Screen.Width - .frModoVideo.Width
            End If
            'vu ahora no se cambia        .VU1.Top = .frDISCOS.Height            '.VU1.Height = Screen.Height - .frDISCOS.Height
            .picVideo.Top = 0
            .picVideo.Left = 0
            .picVideo.Width = .frDISCOS.Width
            .picVideo.Height = .frDISCOS.Height
            .picVideo.Visible = True
        End With
        'habilitar pasar las paginas con teclas simples
        'por que en el modo texto la lista no
        'tiene paginas
        PasarHoja = True
    Else
        EsVideo = False
        'acomodar los controles en modo normal
        With frmIndex
            .VU1.Width = Screen.Width
            If HabilitarVUMetro Then
                .frDISCOS.Left = .VU1.AnchoBarra + 25 ' .VU1.Width
                .frDISCOS.Width = .VU1.Width - (.VU1.AnchoBarra * 2) - 50
                '.frDISCOS.Width = Screen.Width - .VU1.Width
                'vu no se mueve         .VU1.Top = 0                '.VU1.Height = Screen.Height
            Else
                'si viene de un video se tiene que ensanchar
                .frDISCOS.Width = .VU1.Width ' Screen.Width
                .frDISCOS.Left = 0
            End If
            .frModoVideo.Visible = False
            .lblModoVideo.Visible = False
            .frTEMAS.Visible = False
            .lblTEMAS.Visible = False
        End With
        'volver a PasarHoja a su estado original
        PasarHoja = LeerConfig("PasarHoja", "1")
    End If
    'Valores de ReIni FULL=tema ejecutando y lista LISTA=solo lista NADA=arranca de cero
    'si corresponde graba en reini.tbr la lista de temas por sis se corta la luz
   'graba en reini.tbr los datos que correspondan por si se corta la luz
    CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
    
    'reiniciar reloj de tiempo sin uso
    frmIndex.Timer1.Interval = 0
    frmIndex.lblNoUSO = "0"
    SecSinUso = 0
    'lo pongo al ultimo para que tenga tiempo de cargar el tema encargado
    'si lo pongo a donde estaba pasa un pedazito del tema anterior
    
    Unload frmTemasDeDisco
    frmIndex.Refresh
    frmIndex.lblPuesto = "Calculando..."
    'contabilizar para el ranking solo si lo pide
    If SumaRanking Then TOP10 tema, nombreTEMA, nombreDISCO
    'mostrar el puesto que esta en el ranking
    frmIndex.lblPuesto = "Rank # " + PuestoN(tema)
    
    With frmIndex.MP3
        .FileName = tema
        If EsVideo Then
            .DoOpenVideo "child", frmIndex.picVideo.hWnd, 0, 0, (frmIndex.frDISCOS.Width / 15), (frmIndex.lblTemaSonando.Top / 15)
        Else
            .DoOpen
        End If
        .Volumen = VolumenIni
        .DoPlay
    End With
    If HabilitarVUMetro Then frmIndex.VU1.CarFantastic = False
    'If EsVideo Then
        frmIndex.SetFocus 'JOYA JOYA!!! en mp3 da error, no usar
    'End If
    'para qyue tome de nuevo el control del teclado
    Exit Sub
ErrEjecutarTema:
    WriteTBRLog "ERROR EN EJECUTAR TEMA. " + frmIndex.MP3.FileName + "Descripcion: " + Err.Description, True
    If frmIndex.MP3.IsPlaying = False Then EMPEZAR_SIGUIENTE
End Sub

Public Sub EMPEZAR_SIGUIENTE()
    With frmIndex
        'si hay algun elemento en la lista ejecutarlo
        If UBound(MATRIZ_LISTA) > 0 Then
            .lblTemaSonando = "Cargando Proximo Tema..."
            .lblTemaSonando.Refresh
            Dim TemaDeMatriz As String
            TemaDeMatriz = txtInLista(MATRIZ_LISTA(1), 0, ",")
            'reacomodar la matriz para quitar el primer elemento
            Dim c As Long
            For c = 1 To UBound(MATRIZ_LISTA)
                If c < UBound(MATRIZ_LISTA) Then
                    'cuando sea cualquiera menos el ultimo
                    MATRIZ_LISTA(c) = MATRIZ_LISTA(c + 1)
                Else
                    'cuando sea el ultimo
                    'redefinir la matriz con un indice menos
                    ReDim Preserve MATRIZ_LISTA(c - 1)
                    .lstProximos.Clear
                    .lstProximos.AddItem "No hay próximo tema"
                End If
            Next
            CORTAR_TEMA = False 'este tema va entero ya que lo eligio el usuario
            EjecutarTema TemaDeMatriz, True
            CargarProximosTemas
        Else
            'frmINDEX.MP3.SongName = "" 'no sirve
            .Timer1.Interval = 10000
            SecSinUso = 0
            'si no hay temas mostrar la leyenda que lo indica
            .lblTiempoRestante = "FALTA: " + "00:00"
            OnOffCAPS vbKeyCapital, False
            .lblTemaSonando = "Sin reproduccion actual"
            .lblPuesto = "No Rank"
            .lstProximos.Clear
            .lstProximos.AddItem "No hay próximo tema"
            .lblTiempoRestante = "FALTA: " + "00:00"
            .LBLpORCtEMA.Width = .lblTemaSonando.Width
            TEMA_REPRODUCIENDO = "Sin reproduccion actual"
            If HabilitarVUMetro Then frmIndex.VU1.CarFantastic = True
            EsVideo = False 'no estamos rep video
            frmIndex.MP3.DoClose
        End If
    End With
End Sub

Public Sub TOP10(nameARCH As String, nameTEMA As String, nameDISCO As String)
    'On Error GoTo notop
    'ver si existe ranking.tbr
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        FSO.CreateTextFile AP + "ranking.tbr", True
    End If
    
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
    Encontrado = False
    'abrir el archivo y ver si ya esta el tema
    
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    'leerlo cargarlo en matriz y ordenar por mas escuchado
    Do While Not TE.AtEndOfStream
        'cada linea es "puntos,arch,nombretema,nombredisco"
        TT = TE.ReadLine
        If TT <> "" Then
            z = z + 1
            ThisPTS = Val(txtInLista(TT, 0, ","))
            ThisArch = txtInLista(TT, 1, ",")
            ThisTEMA = txtInLista(TT, 2, ",")
            ThisDISCO = txtInLista(TT, 3, ",")
            ReDim Preserve mtxTOP10(z)
            'comparar este tema con el elegido actual
            
            If UCase(Trim(nameARCH)) = UCase(Trim(ThisArch)) Then
                'sumarle un punto
                ThisPTS = ThisPTS + 1
                'marcar esta cantidad de puntos como referencai futura para
                'agregar el nuevo dato al ranking
                PTnuevo = ThisPTS
                TT = Str(ThisPTS) + "," + ThisArch + "," + ThisTEMA + "," + ThisDISCO
                DatoNuevoFull = TT
                ArchivoNuevo = ThisArch
                Encontrado = True
            End If
            mtxTOP10(z) = TT
        End If
    Loop
     TE.Close
    'ver si el archivo habia sido votado
    If Encontrado = False Then
        TT = "1," + Trim(nameARCH) + "," + Trim(nameTEMA) + "," + Trim(nameDISCO)
        ReDim Preserve mtxTOP10(z + 1)
        mtxTOP10(z + 1) = TT
        PTnuevo = 1
        DatoNuevoFull = TT
        ArchivoNuevo = nameARCH
    End If
    'cargar todos y sacar la primera columna de las zetas
    Dim MTXsort() As String
    Set TE = FSO.CreateTextFile(AP + "ranking.tbr", True)
    Dim PTactual As Long
    Dim YaSeEscribioDatoNuevo As Boolean
    Dim VarMTX As Long 'variacion del indice de la matriz
    YaSeEscribioDatoNuevo = False
    VarMTX = 0
    For mtx = 1 To UBound(mtxTOP10)
        ReDim Preserve MTXsort(mtx + 1)
        PTactual = txtInLista(mtxTOP10(mtx), 0, ",")
        If PTactual = PTnuevo And YaSeEscribioDatoNuevo = False Then
            MTXsort(mtx) = DatoNuevoFull
            TE.WriteLine MTXsort(mtx)
            YaSeEscribioDatoNuevo = True
            mtx = mtx - 1
            VarMTX = 1
        Else
            If Trim(UCase(ArchivoNuevo)) = Trim(UCase(txtInLista(mtxTOP10(mtx), 1, ","))) Then
                VarMTX = 0
                GoTo SIG
            End If
            MTXsort(mtx + VarMTX) = CStr(PTactual) + "," + _
                txtInLista(mtxTOP10(mtx), 1, ",") + "," + _
                txtInLista(mtxTOP10(mtx), 2, ",") + "," + _
                txtInLista(mtxTOP10(mtx), 3, ",")
            TE.WriteLine MTXsort(mtx + VarMTX)
        End If
SIG:
    Next
    TE.Close
    Exit Sub
notop:
    MsgBox Err.Description
    
End Sub

Public Sub SumarContadorCreditos(valorSUMAR As Long)
    Dim ARCHcont As String
    'ver el valor en win
    ARCHcont = WINfolder + "\nnr.dll"
    If FSO.FileExists(ARCHcont) Then
        Set TE = FSO.OpenTextFile(ARCHcont, ForReading, False)
        Dim CONTw As Long
        CONTw = Val(TE.ReadLine)
        TE.Close
    Else
        Set TE = FSO.CreateTextFile(ARCHcont, True)
        TE.WriteLine "0"
        TE.Close
        CONTw = 0
    End If
    
    'ver el valor en sys
    ARCHcont = SYSfolder + "\nnr.dll"
    If FSO.FileExists(ARCHcont) Then
        Set TE = FSO.OpenTextFile(ARCHcont, ForReading, False)
        Dim CONTs As Long
        CONTs = Val(TE.ReadLine)
        TE.Close
    Else
        Set TE = FSO.CreateTextFile(ARCHcont, True)
        TE.WriteLine "0"
        TE.Close
        CONTs = 0
    End If
    
    Dim ContFinal As Long
    If CONTw <> CONTs Then
        If CONTw > CONTs Then
            ContFinal = CONTw
        Else
            ContFinal = CONTs
        End If
    Else
        'aqui vale cualquiera de los dos por que son iguales
        ContFinal = CONTs
    End If
    'sumar lo que corresponde
    ContFinal = ContFinal + valorSUMAR
    'asignarlo a la variable global
    CONTADOR = ContFinal
    
    'actualizar el valor en win
    ARCHcont = WINfolder + "\nnr.dll"
    Set TE = FSO.CreateTextFile(ARCHcont, True)
    TE.WriteLine Trim(Str(CONTADOR))
    TE.Close
    
    'actualizar el valor en sys
    ARCHcont = SYSfolder + "\nnr.dll"
    Set TE = FSO.CreateTextFile(ARCHcont, True)
    TE.WriteLine Trim(Str(CONTADOR))
    TE.Close
    
End Sub

Public Function PuestoN(TemaBuscado As String) As String
    'leer ranking.tbr y buscar el tema
    
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        'esto no deberia pasar nunca ya que entra despues de que el tema se carga en el ranking
        FSO.CreateTextFile AP + "ranking.tbr", True
        PuestoN = 1
        Exit Function
    End If
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    Dim TT As String
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    
    Dim PuestoActual As Long
    PuestoActual = 0
    Do While Not TE.AtEndOfStream
        TT = TE.ReadLine
        ThisPTS = Val(txtInLista(TT, 0, ","))
        ThisArch = txtInLista(TT, 1, ",")
        ThisTEMA = txtInLista(TT, 2, ",")
        ThisDISCO = txtInLista(TT, 3, ",")
        If FSO.FileExists(ThisArch) Then
            PuestoActual = PuestoActual + 1
            If UCase(ThisArch) = UCase(TemaBuscado) Then
                PuestoN = Trim(Str(PuestoActual))
                Exit Function
            End If
        End If
    Loop
    PuestoN = "No Rank"
End Function
