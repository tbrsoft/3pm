VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMMCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'carrito de compra de multimedia

Private mListaMM() As String 'lista de todos los paths de cada archivo en el carrito
Private mListaMMFull() As String 'a diferencia de la anterior si busca todas las canciones _
    dentro de un disco cuando se elige completo

Private mTotalAudio As Long
Private mTotalVideo As Long

Private mArchivo As String 'se graba a disco por cortes de luz

'Private mModoCompra As Long
'por si quiere com pra basica sin elegir destino o avanzada eligiendo
'destino y cosas a borrar si hubiera lugar
'simple basico =1
'avanzado=2

'el precios se puede calcular por: _
    MB _
    Canciones mp3 + videos _
    Opcionalmente puede haber canciones mas caras

Private mPricesAudio() As Single
Private mPricesVideo() As Single

'por ejemplo mPricesAudio(1)=3
'            mPricesAudio(6)=10
'            mPricesVideo(1)=4
'            mPricesVideo(8)=24

'quiere decir que 1 cancion cuestan 3 creditos cada una _
    6 o mas canciones cuestan 10/6 CREDITOS cada una
'cada video cuesta 4 creditos pero si son 8 o mas cuesta 3 cada uno

Private CarritoChange As Boolean
Private mGetTotalMB As Single

Public Function SetPricesAudio(XaCantidad As Long, PriceCred As Long) As Long
    
    If XaCantidad > UBound(mPricesAudio) Then
        ReDim Preserve mPricesAudio(XaCantidad)
        mPricesAudio(XaCantidad) = PriceCred
        'devuelve el precio anterior
        SetPricesAudio = 0
    Else
        SetPricesAudio = mPricesAudio(XaCantidad)
        mPricesAudio(XaCantidad) = PriceCred
    End If
    
    CarritoChange = True
End Function

Public Function SetPricesVideo(XaCantidad As Long, PriceCred As Long) As Long
    If XaCantidad > UBound(mPricesVideo) Then
        ReDim Preserve mPricesVideo(XaCantidad)
        mPricesVideo(XaCantidad) = PriceCred
        'devuelve el precio anterior
        SetPricesVideo = 0
    Else
        SetPricesVideo = mPricesVideo(XaCantidad)
        mPricesVideo(XaCantidad) = PriceCred
    End If
    
    CarritoChange = True
End Function

Public Function GetPricesVideo(XaCantidad As Long) As Long
    'busca el precio para la cantidad indicada, si no existe igual baja hasta encontrar
    
    If XaCantidad = 0 Then
        GetPricesVideo = 0
        Exit Function
    End If
    
    Dim Xa As Long
    Xa = XaCantidad
    
    'voy bajando hasta encontrar un precio
    For Xa = XaCantidad To 1 Step -1
        'ver que mida si da la matriz
        If Xa > UBound(mPricesVideo) Then GoTo sig
        
        If mPricesVideo(Xa) > 0 Then
            'ya se cual es el precio unitario por la cantidad que elegi
            'ahora veo el precio que le hago por todo
            Dim SubTot As Single
            SubTot = (mPricesVideo(Xa) / Xa)
            SubTot = SubTot * XaCantidad
            'ahora redondeo para no andar con faking monedas
            GetPricesVideo = CLng(SubTot)
            Exit Function
        End If
sig:
    Next Xa
    
    GetPricesVideo = 0 'ES GRATIS! YO LE COBRO IGUAL ...
    
End Function

Public Sub CleanPrices()
    ReDim mPricesAudio(0)
    ReDim mPricesVideo(0)
    
    CarritoChange = True
    
End Sub

Public Function GetPricesAudio(XaCantidad As Long) As Long
    'busca el precio para la cantidad indicada, si no existe igual baja hasta encontrar
    
    If XaCantidad = 0 Then
        GetPricesAudio = 0
        Exit Function
    End If
    
    Dim Xa As Long
    Xa = XaCantidad
    
    'voy bajando hasta encontrar un precio
    For Xa = XaCantidad To 1 Step -1
        'ver que mida si da la matriz
        If Xa > UBound(mPricesAudio) Then GoTo sig
        
        If mPricesAudio(Xa) > 0 Then
            'ya se cual es el precio unitario por la cantidad que elegi
            'ahora veo el precio que le hago por todo
            Dim SubTot As Single
            SubTot = (mPricesAudio(Xa) / Xa)
            SubTot = SubTot * XaCantidad
            'ahora redondeo para no andar con faking monedas
            GetPricesAudio = CLng(SubTot)
            Exit Function
        End If
sig:
    Next Xa
    
    GetPricesAudio = 0 'ES GRATIS!
End Function


'Public Property Let ModoCompra(NewModo As Long)
'    mModoCompra = NewModo
'End Property
'
'Public Property Get ModoCompra() As Long
'    ModoCompra = mModoCompra
'End Property

Public Function CalculateTotalPrice() As Single

    Dim PrAu As Long 'precio del audio según la cantidad que tiene
    Dim PrVi As Long 'precio del video según la cantidad que tiene

    PrAu = GetPricesAudio(mTotalAudio)
    PrVi = GetPricesVideo(mTotalVideo)
    
    CalculateTotalPrice = PrAu + PrVi 'esta expresado en creditos
End Function

Public Function CleanFile(i As Long) As Long
    'borrar un archivo ni bien se grabo
    If (i > UBound(mListaMMFull)) Or (i = 0) Then
        Exit Function
    End If
    
    Dim H As Long
    mListaMMFull(i) = ""
    'si es el ultimo es una cosa
    If i = UBound(mListaMMFull) Then
        ReDim Preserve mListaMMFull(i - 1)
    Else 'si no corro todo para abajo
        For H = i To UBound(mListaMMFull) - 1
            mListaMMFull(H) = mListaMMFull(H + 1)
        Next H
        
        ReDim Preserve mListaMMFull(UBound(mListaMMFull) - 1)
    End If
    CleanEmptyFolder
    CleanFile = 0
End Function

Public Function CleanEmptyFolder()
    'ver integridad de carpetas elegidas con canciones
    'ver si las carpetas que hayan sido elegidas tengan al menos una cancion
    'estro porque puede ser que se borre el unico archivo de una carpeta
    
    Dim H As Long
    
    'como se cambia el fin del for tengo un do-loop
    Do
INIAGAIN:
        For H = 1 To UBound(mListaMM)
            T = mListaMM(H)
            Dim Dep As Long
            Dim Z As Long
            If Right(T, 1) = "\" Then
                Dep = GetDependent(H)
                If Dep = 0 Then 'carpeta sin nada!
                    mListaMM(H) = "" 'marco para borrar
                    'BORRAR Y RETRAE LA MATRIZ EN 1
                    'si es el ultimo es una cosa
                    If H = UBound(mListaMM) Then
                        ReDim Preserve mListaMM(H - 1)
                    Else 'si no corro todo para abajo
                        For Z = H To UBound(mListaMM) - 1
                            mListaMM(Z) = mListaMM(Z + 1)
                        Next Z
                        ReDim Preserve mListaMM(UBound(mListaMM) - 1)
                    End If
                    GoTo INIAGAIN
                End If
            Else
                'si es un mp3 y se borro su correspondiente en listaFull lo saco tambien!
                Dim Esta As Boolean
                Esta = False
                For Z = 1 To UBound(mListaMMFull)
                    If T = mListaMMFull(Z) Then
                        Esta = True
                        Exit For
                    End If
                Next Z
                
                If Esta = False Then 'lo saco de la lista basica tambien!!
                    If H = UBound(mListaMM) Then
                        ReDim Preserve mListaMM(H - 1)
                    Else 'si no corro todo para abajo
                        For Z = H To UBound(mListaMM) - 1
                            mListaMM(Z) = mListaMM(Z + 1)
                        Next Z
                        
                        ReDim Preserve mListaMM(UBound(mListaMM) - 1)
                    End If
                End If
            End If
        Next H
        'si llego hasta aca es por que termine por fin
        If H >= UBound(mListaMM) Then Exit Do
    Loop
    
    CarritoChange = True
End Function

Public Function CleanSelSoloMarcados() As Long
    Dim H As Long
    'limpiar uno por uno prolijamente
    'Dim Limpio As Boolean
    'Do While Limpio = False
    '    Limpio = False
        For H = 1 To UBound(mListaMM)
            If mListaMM(H) = "" Then
                CleanSel H
                Limpio = True
                Exit For
            End If
       Next H
    'Loop
    
    CarritoChange = True
    
End Function

Public Function CleanFileSoloMarca(i As Long) As Long
    'marcar para eliminar
    If (i > UBound(mListaMMFull)) Or (i = 0) Then
        Exit Function
    End If
    
    CarritoChange = True
    
    mListaMMFull(i) = ""
    
End Function

Public Function CleanMarcados() As Long
    Dim H As Long
    'limpiar uno por uno prolijamente
    'Dim Limpio As Boolean
    'Do While Limpio = False
    '    Limpio = False
        For H = 1 To UBound(mListaMMFull)
            If mListaMMFull(H) = "" Then
                CleanFile H
                Limpio = True
                Exit For
            End If
       Next H
    'Loop
    
    CarritoChange = True
    
End Function

Public Function CleanSel(i As Long) As Long
    'borrar de lista de selecciones
    'opcionalmente eliminar en cascada
    
    If (i > UBound(mListaMM)) Or (i = 0) Then
        Exit Function
    End If
    
    Dim T As String, H As Long
    T = mListaMM(i)
    If Right(T, 1) = "\" Then
        'eliminar sus dependencias
        Dim EP As String
        For H = 1 To UBound(mListaMMFull)
            EP = fso.GetParentFolderName(GetElementFull(H)) + "\"
            If EP = mListaMM(i) Then
                'marco para eliminar despues por que se cambian los indices y es un lio 'CleanFile i
                mListaMMFull(H) = ""
            End If
        Next H
        
        CleanMarcados
    Else
        'eliminar normal
        'ver que numero tiene en la matriz full para usar CleanFile
        Dim Y As Long
        For Y = 1 To UBound(mListaMMFull)
            If mListaMMFull(Y) = mListaMM(i) Then
                CleanFile Y
                
                'borrar de la lista de selecciones para que no se muestre
                
                
                Exit For
            End If
        Next Y
        
    End If
    
    'sacar de la matriz el elemento borrado
    CleanEmptyFolder
    
    CarritoChange = True
    
    CleanSel = 0
End Function

Public Function GetDependent(i As Long) As Long
    'saber cuantos archivos estan enganchados a una carpeta
    If (i > UBound(mListaMM)) Or (i = 0) Then
        Exit Function
    End If
    
    Dim D As Long
    D = 0 'contador temporal
    'ver si es una cancion o una carpeta
    Dim T As String
    T = mListaMM(i)
    If Right(T, 1) = "\" Then
        'ver cuantas canciones tiene asignadas
        Dim J As Long
        For J = 1 To UBound(mListaMMFull)
            If LCase(T) = LCase(GetElementFullListPath(J)) Then
                D = D + 1
            End If
        Next J
    End If

    GetDependent = D

End Function

Public Function GetDependentOf(myFolder As String) As Long
    'ubicar si la carpeta de un archivo es una seleccion o este se eligio suelto
    'si es una carpeta devuelve su indice en la matriz mListaMM, si no es -1
    
    Dim J As Long
    For J = 1 To UBound(mListaMM)
        If LCase(myFolder) = LCase(GetElement(J)) Then
            GetDependentOf = J
            Exit Function
        End If
    Next J

    GetDependentOf = -1
End Function


Public Function GetTotalAudio() As Long
    GetTotalAudio = mTotalAudio
End Function

Public Function GetTotalVideo() As Long
    GetTotalVideo = mTotalVideo
End Function

Public Function AddFile(nFile As String) As Long
    Dim H As Long
    
    'ver si ya existe!!!
    If ExisteFileInCart(nFile) = False Then
        H = UBound(mListaMM) + 1
        ReDim Preserve mListaMM(H)
        mListaMM(H) = nFile
        
        H = UBound(mListaMMFull) + 1
        ReDim Preserve mListaMMFull(H)
        mListaMMFull(H) = nFile
        
        Select Case LCase(fso.GetExtensionName(nFile))
            Case "mp3", "wma"
                mTotalAudio = mTotalAudio + 1
            Case "mpeg", "mpg", "avi", "wmv", "vob", "dat"
                mTotalVideo = mTotalVideo + 1
            Case "mn0", "mn1"
            
        End Select
        CarritoChange = True
    Else
        AddFile = 1 'YA ESTABA!
        Exit Function
    End If
    
    AddFile = 0
End Function

Public Function AddFolder(nFolder As String) As Long
    'agrega toda una carpeta sin rutina recursiva
    
    On Local Error GoTo ER
    
    AddFolder = 1
    If Right(nFolder, 1) <> "\" Then nFolder = nFolder + "\"
    
    If fso.FolderExists(nFolder) = False Then Exit Function
    'marco el disco entero para que se vea en la lista solo eso
    'por si alguien elige muchos discos llegue a ver todo lo que eligio!
    
    'ver si ya existe!!!
    
    Dim H As Long
    
    'buscar toda la multimedia que se puede copiar
    Dim T() As String
    AddFolder = 2
    T = ObtenerArchMM(nFolder, True)
    Dim J As Long
    
    Dim CantSel As Long
    CantSel = 0 'me fijo cuantos realmente no estaban para ver si se cargo todo o ninguno
    Dim nFile As String
    For H = 1 To UBound(T)
        nFile = txtInLista(T(H), 0, "#")
        'nombreTemas = txtInLista(MATRIZ_TEMAS(c), 1, "#")
        'ME ASEGURO QUE NO ESTE ELGIDO YA
        If ExisteFileInCart(nFile) = False Then
            CantSel = CantSel + 1
            J = UBound(mListaMMFull) + 1
            ReDim Preserve mListaMMFull(J)
            mListaMMFull(J) = nFile
            'sumar los contadores para los precios
            Select Case LCase(fso.GetExtensionName(nFile))
                Case "mp3", "wma"
                    mTotalAudio = mTotalAudio + 1
                Case "mpeg", "mpg", "avi", "wmv", "vob", "dat"
                    mTotalVideo = mTotalVideo + 1
                Case "mn0", "mn1"
                
            End Select
            
        End If
    Next H
    
    AddFolder = 3
    If CantSel = 0 Then
        'ya estaban TODAS de antes
        Exit Function
    End If
    
    If CantSel < UBound(T) Then
        'deberia avisar que al menos una cancion ya estaba
    End If
    
    H = UBound(mListaMM) + 1
    ReDim Preserve mListaMM(H)
    mListaMM(H) = nFolder 'como termina en "\" sabre despues que es una carpeta
    
    AddFolder = 0
    CarritoChange = True
    Exit Function
ER:
     
End Function

Public Function GetElement(i As Long) As String
    If (i > UBound(mListaMM)) Or (i = 0) Then
        GetElement = ""
        Exit Function
    End If
    
    GetElement = mListaMM(i)
End Function

Public Function GetPricesAudioBase(i As Long) As Long
    If (i > UBound(mPricesAudio)) Or (i = 0) Then
        GetPricesAudioBase = 0
        Exit Function
    End If
    
    GetPricesAudioBase = mPricesAudio(i)
End Function

Public Function KillPricesAudioBase(i As Long) As Long
    If (i > UBound(mPricesAudio)) Or (i = 0) Then
        KillPricesAudioBase = -1
        Exit Function
    End If
    
    mPricesAudio(i) = 0
    KillPricesAudioBase = 0
    
    CarritoChange = True
End Function

Public Function KillPricesVideoBase(i As Long) As Long
    If (i > UBound(mPricesVideo)) Or (i = 0) Then
        KillPricesVideoBase = -1
        Exit Function
    End If
    
    mPricesVideo(i) = 0
    KillPricesVideoBase = 0
    
    CarritoChange = True
End Function



Public Function GetPricesVideoBase(i As Long) As Long
    If (i > UBound(mPricesVideo)) Or (i = 0) Then
        GetPricesVideoBase = 0
        Exit Function
    End If
    
    GetPricesVideoBase = mPricesVideo(i)
End Function

Public Function GetTotalPricesAudio() As Long
    GetTotalPricesAudio = UBound(mPricesAudio)
End Function

Public Function GetTotalPricesVideo() As Long
    GetTotalPricesVideo = UBound(mPricesVideo)
End Function

Public Function GetElementFull(i As Long) As String
    If (i > UBound(mListaMMFull)) Or (i = 0) Then
        GetElementFull = ""
        Exit Function
    End If
    
    GetElementFull = mListaMMFull(i)
End Function


Public Function GetElementPath(i As Long) As String
    If (i > UBound(mListaMM)) Or (i = 0) Then
        GetElementPath = ""
        Exit Function
    End If
    
    'ver si es una cancion o una carpeta
    Dim T As String
    If Right(mListaMM(i), 1) = "\" Then
        T = mListaMM(i)
    Else
        T = fso.GetParentFolderName(mListaMM(i))
    End If
    
    If Right(T, 1) <> "\" Then T = T + "\"
    
    GetElementPath = T
End Function

Public Function GetElementFullListPath(i As Long) As String
    If (i > UBound(mListaMMFull)) Or (i = 0) Then
        GetElementFullListPath = ""
        Exit Function
    End If
    
    'es una cancion si o si
    T = fso.GetParentFolderName(mListaMMFull(i))
    
    If Right(T, 1) <> "\" Then T = T + "\"
    
    GetElementFullListPath = T
End Function

Public Function GetTotalMB() As Single
    'calcula el peso de todos los archivos
    
    Dim TMP As Single: TMP = 0
    
    'PODRIA VER SI NO HUBO CAMBIOS Y NO LEER TODO DE NUEVO
    If CarritoChange = True Then
        Dim H As Long
        For H = 1 To UBound(mListaMMFull)
            If fso.FileExists(mListaMMFull(H)) Then
                TMP = TMP + ((FileLen(mListaMMFull(H)) / 1024) / 1024)
            End If
        Next H
        mGetTotalMB = Round(TMP, 2)
        GetTotalMB = mGetTotalMB
    Else
        'ver valor anterior
        GetTotalMB = mGetTotalMB
    End If
End Function

Public Function GetElementName(i As Long) As String
    If (i > UBound(mListaMM)) Or (i = 0) Then
        GetElementName = ""
        Exit Function
    End If
    
    Dim T As String, T2 As String
    
    'ver si es una cancion o una carpeta
    If Right(mListaMM(i), 1) = "\" Then
        'rrrrrrrr nueva traducion
        T = TR.Trad("Disco completo:%99%") + " " + fso.GetBaseName(mListaMM(i))
        GetElementName = T
    Else
        T = fso.GetBaseName(fso.GetParentFolderName(mListaMM(i)))
        T2 = fso.GetBaseName(mListaMM(i))
        GetElementName = T2 + " / " + T
    End If
End Function

Public Function GetFileCant() As Long
    Dim H As Long
    H = UBound(mListaMM)
    
    GetFileCant = H
End Function

Public Function GetFileCantFull() As Long
    Dim H As Long
    H = UBound(mListaMMFull)
    GetFileCantFull = H
End Function

Private Sub Class_Initialize()
    ReDim mListaMM(0) 'indice siempre vacio
    ReDim mListaMMFull(0)
    mTotalAudio = 0
    mTotalVideo = 0
    mArchivo = "c:\cart.txt" 'y si no hay "c:\" YA ME PASO!
    'mModoCompra = 0
    CleanPrices
    CarritoChange = False
End Sub

Public Sub SetFileSave(sFile As String) 'elijo donde se va a agrabar
    mArchivo = sFile
End Sub
Public Function LoadCartFromDisk() As Long
    On Local Error GoTo ER
    
    LoadCartFromDisk = 1
    If fso.FileExists(mArchivo) = False Then Exit Function
    ReDim mListaMM(0)
    ReDim mListaMMFull(0)
    
    LoadCartFromDisk = 2
    Dim TE As TextStream, TX As String
    Set TE = fso.OpenTextFile(mArchivo, ForReading, False)
        TX = TE.ReadAll
    TE.Close
    Set TE = Nothing
    
    LoadCartFromDisk = 3
    Dim SP() As String
    SP = Split(TX, Chr(6))
    
    Dim H As Long
    For H = 0 To UBound(SP)
        ReDim Preserve mListaMM(H + 1) 'la matriz inicia en 1 siempre
        mListaMM(H + 1) = SP(H)
    Next H
    
    LoadCartFromDisk = 0
    
    CarritoChange = True
    
    Exit Function
ER:
End Function

Public Function SaveCartToDisk() As Long
    On Local Error GoTo ER
    
    SaveCartToDisk = 1
    If fso.FileExists(mArchivo) Then fso.DeleteFile mArchivo, True
    SaveCartToDisk = 2
    If UBound(mListaMM) = 0 Then Exit Function
    
    Dim TE As TextStream, H As Long
    Set TE = fso.CreateTextFile(mArchivo, True)
    
        For H = 1 To UBound(mListaMM)
            TE.Write mListaMM(H)
            'para que al leer el split no me de un elemento vacio al ultimo
            If H < UBound(mlista) Then TE.Write Chr(6)
        Next H
    
    TE.Close
    Set TE = Nothing
    
    SaveCartToDisk = 0
    Exit Function
    
ER:
End Function

Public Function ClearCart() As Long
    ReDim mListaMM(0)
    ReDim mListaMMFull(0)
    mTotalAudio = 0
    mTotalVideo = 0
    
    CarritoChange = True
End Function

Public Function ExisteFileInCart(nFile As String) As Boolean
    'ver si hay algo!
    Dim H As Long, B As Boolean
    B = False
    
    For H = 1 To UBound(mListaMMFull)
        If LCase(mListaMMFull(H)) = LCase(nFile) Then
            B = True
            Exit For
        End If
    Next H
    
    ExisteFileInCart = B
End Function

Public Function SavePrices(sFile As String)

    If fso.FileExists(sFile) Then fso.DeleteFile sFile
    Dim TePC As TextStream, H As Long
    
    Set TePC = fso.CreateTextFile(sFile, True)
        For H = 1 To GetTotalPricesAudio
            If GetPricesAudioBase(H) > 0 Then
                TePC.Write "A|" + CStr(H) + "|" + CStr(GetPricesAudioBase(H)) + Chr(6)
            End If
        Next H
        
        For H = 1 To GetTotalPricesVideo
            If GetPricesVideoBase(H) > 0 Then
                TePC.Write "V|" + CStr(H) + "|" + CStr(GetPricesVideoBase(H)) + Chr(6)
            End If
        Next H
    TePC.Close

End Function

Public Function LoadPrices(sFile As String)

    'precio del carrito de compras
    Dim TePC As TextStream 'TextStreamPrecioCarrito
VerPromoCart:
    
    If fso.FileExists(sFile) = False Then
        tERR.Anotar "dacb"
        Set TePC = fso.CreateTextFile(sFile, True)
            TePC.Write "A|1|2" + Chr(6) 'Audio 1 cancion 2 creditos
            TePC.Write "V|1|4" + Chr(6) 'Video 1 cancion 4 creditos
            
            TePC.Write "A|15|15" + Chr(6) 'Audio 15 canciones 15 creditos (1 cada uno)
            TePC.Write "V|10|19" + Chr(6) 'Video 10 videos 19 creditos (1.9 cred cada uno)
        TePC.Close
    End If
    
    'leer el que sea
    Dim TMP As String
    Set TePC = fso.OpenTextFile(sFile, ForReading, False)
        If TePC.AtEndOfStream Then
            TePC.Close
            fso.DeleteFile sFile, True
            GoTo VerPromoCart
        Else
            TMP = TePC.ReadAll
        End If
    TePC.Close
    tERR.Anotar "dacc", TMP
    
    Dim SP8() As String
    SP8 = Split(TMP, Chr(6))
    If UBound(SP8) = 0 Then
        tERR.AppendLog "dacd"
        fso.DeleteFile sFile, True
        GoTo VerPromoCart
    Else
        Dim H As Long, SP9() As String
        For H = 0 To UBound(SP8)
            SP9 = Split(SP8(H), "|")
            'VER SI ES ULTIMO VACIO!
            'ACABO DE INVENTARLO MIENTRAS DEPURO!!!!!!!!!!!!!
            If (IsEmpty(SP9)) Then GoTo sig
            If UBound(SP9) = -1 Then GoTo sig
            If UBound(SP9) < 2 Then 'al menos 3 indices tiene que tener
                tERR.AppendLog "dace"
                fso.DeleteFile sFile, True
                GoTo VerPromoCart
            Else
                Select Case UCase(SP9(0))
                    'musica
                    Case "A": SetPricesAudio CLng(SP9(1)), CLng(SP9(2))
                    'videillo
                    Case "V": SetPricesVideo CLng(SP9(1)), CLng(SP9(2))
                    Case Else: tERR.AppendLog "FuckCart:" + SP9(0)
                End Select
            End If
sig:
        Next H
        'si no salio un puto precio ok entonces va de nuevo
        'es casi imposible despues de tantas verificaciones
    End If

    CarritoChange = True
End Function
