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

Private Type CartCancion
    cPath As String
    cPerfil As Long
End Type

Private Type CartElement 'puede ser un disco o un elementro suelto. Depende de las selecciones del cliente
    cPath As String
    cPerfil As String
    cFolder As Boolean '0 si es archivo, 1 si es carpeta
End Type

Private mListaMM() As CartElement 'lista de todos los paths de cada archivo en el carrito
Private mListaMMFull() As CartCancion 'a diferencia de la anterior si busca todas las canciones _
    dentro de un disco cuando se elige completo

Private mTotalAudio As Long
Private mTotalVideo As Long
Private mTotalRingtones As Long
Private mTotalWallpapers As Long
Private mTotalJava As Long
Private mTotalIso As Long 'mp01
Private mTotal3GP As Long
Private mTotalThemes As Long
Private mTotalIsoDVD As Long 'mp01

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
Private mPricesRingtones() As Single
Private mPricesWallpapers() As Single
Private mPricesJava() As Single
Private mPricesISO() As Single 'mm91
Private mPrices3GP() As Single 'mm91
Private mPricesThemes() As Single
Private mPricesIsoDVD() As Single 'mp01

'por ejemplo mPricesAudio(1)=3
'            mPricesAudio(6)=10
'            mPricesVideo(1)=4
'            mPricesVideo(8)=24

'quiere decir que 1 cancion cuestan 3 creditos cada una _
    6 o mas canciones cuestan 10/6 CREDITOS cada una
'cada video cuesta 4 creditos pero si son 8 o mas cuesta 3 cada uno

Private CarritoChange As Boolean
Private mGetTotalMB As Single
Private mGetTotalMinutos As Single

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

Public Function SetPricesRingtones(XaCantidad As Long, PriceCred As Long) As Long
    
    If XaCantidad > UBound(mPricesRingtones) Then
        ReDim Preserve mPricesRingtones(XaCantidad)
        mPricesRingtones(XaCantidad) = PriceCred
        'devuelve el precio anterior
        SetPricesRingtones = 0
    Else
        SetPricesRingtones = mPricesRingtones(XaCantidad)
        mPricesRingtones(XaCantidad) = PriceCred
    End If
    
    CarritoChange = True
End Function

Public Function SetPricesWallpapers(XaCantidad As Long, PriceCred As Long) As Long
    
    If XaCantidad > UBound(mPricesWallpapers) Then
        ReDim Preserve mPricesWallpapers(XaCantidad)
        mPricesWallpapers(XaCantidad) = PriceCred
        'devuelve el precio anterior
        SetPricesWallpapers = 0
    Else
        SetPricesWallpapers = mPricesWallpapers(XaCantidad)
        mPricesWallpapers(XaCantidad) = PriceCred
    End If
    
    CarritoChange = True
End Function

Public Function SetPricesJava(XaCantidad As Long, PriceCred As Long) As Long
    
    If XaCantidad > UBound(mPricesJava) Then
        ReDim Preserve mPricesJava(XaCantidad)
        mPricesJava(XaCantidad) = PriceCred
        'devuelve el precio anterior
        SetPricesJava = 0
    Else
        SetPricesJava = mPricesJava(XaCantidad)
        mPricesJava(XaCantidad) = PriceCred
    End If
    
    CarritoChange = True
End Function

Public Function SetPricesIso(XaCantidad As Long, PriceCred As Long) As Long 'mm91
    
    If XaCantidad > UBound(mPricesISO) Then
        ReDim Preserve mPricesISO(XaCantidad)
        mPricesISO(XaCantidad) = PriceCred
        'devuelve el precio anterior
        SetPricesIso = 0
    Else
        SetPricesIso = mPricesISO(XaCantidad)
        mPricesISO(XaCantidad) = PriceCred
    End If
    
    CarritoChange = True
End Function

'mp01
Public Function SetPricesIsoDVD(XaCantidad As Long, PriceCred As Long) As Long 'mm91
    
    If XaCantidad > UBound(mPricesIsoDVD) Then
        ReDim Preserve mPricesIsoDVD(XaCantidad)
        mPricesIsoDVD(XaCantidad) = PriceCred
        'devuelve el precio anterior
        SetPricesIsoDVD = 0
    Else
        SetPricesIsoDVD = mPricesIsoDVD(XaCantidad)
        mPricesIsoDVD(XaCantidad) = PriceCred
    End If
    
    CarritoChange = True
End Function


Public Function SetPrices3GP(XaCantidad As Long, PriceCred As Long) As Long 'mm91
    
    If XaCantidad > UBound(mPrices3GP) Then
        ReDim Preserve mPrices3GP(XaCantidad)
        mPrices3GP(XaCantidad) = PriceCred
        'devuelve el precio anterior
        SetPrices3GP = 0
    Else
        SetPrices3GP = mPrices3GP(XaCantidad)
        mPrices3GP(XaCantidad) = PriceCred
    End If
    
    CarritoChange = True
End Function

Public Function SetPricesThemes(XaCantidad As Long, PriceCred As Long) As Long 'mm91
    
    If XaCantidad > UBound(mPricesThemes) Then
        ReDim Preserve mPricesThemes(XaCantidad)
        mPricesThemes(XaCantidad) = PriceCred
        'devuelve el precio anterior
        SetPricesThemes = 0
    Else
        SetPricesThemes = mPricesThemes(XaCantidad)
        mPricesThemes(XaCantidad) = PriceCred
    End If
    
    CarritoChange = True
End Function



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

Public Function GetPricesRingtones(XaCantidad As Long) As Long
    'busca el precio para la cantidad indicada, si no existe igual baja hasta encontrar
    
    If XaCantidad = 0 Then
        GetPricesRingtones = 0
        Exit Function
    End If
    
    Dim Xa As Long
    Xa = XaCantidad
    
    'voy bajando hasta encontrar un precio
    For Xa = XaCantidad To 1 Step -1
        'ver que mida si da la matriz
        If Xa > UBound(mPricesRingtones) Then GoTo sig
        
        If mPricesRingtones(Xa) > 0 Then
            'ya se cual es el precio unitario por la cantidad que elegi
            'ahora veo el precio que le hago por todo
            Dim SubTot As Single
            SubTot = (mPricesRingtones(Xa) / Xa)
            SubTot = SubTot * XaCantidad
            'ahora redondeo para no andar con faking monedas
            GetPricesRingtones = CLng(SubTot)
            Exit Function
        End If
sig:
    Next Xa
    
    GetPricesRingtones = 0 'ES GRATIS!
End Function

Public Function GetPricesWallpapers(XaCantidad As Long) As Long
    'busca el precio para la cantidad indicada, si no existe igual baja hasta encontrar
    
    If XaCantidad = 0 Then
        GetPricesWallpapers = 0
        Exit Function
    End If
    
    Dim Xa As Long
    Xa = XaCantidad
    
    'voy bajando hasta encontrar un precio
    For Xa = XaCantidad To 1 Step -1
        'ver que mida si da la matriz
        If Xa > UBound(mPricesWallpapers) Then GoTo sig
        
        If mPricesWallpapers(Xa) > 0 Then
            'ya se cual es el precio unitario por la cantidad que elegi
            'ahora veo el precio que le hago por todo
            Dim SubTot As Single
            SubTot = (mPricesWallpapers(Xa) / Xa)
            SubTot = SubTot * XaCantidad
            'ahora redondeo para no andar con faking monedas
            GetPricesWallpapers = CLng(SubTot)
            Exit Function
        End If
sig:
    Next Xa
    
    GetPricesWallpapers = 0 'ES GRATIS!
End Function

Public Function GetPricesJava(XaCantidad As Long) As Long
    'busca el precio para la cantidad indicada, si no existe igual baja hasta encontrar
    
    If XaCantidad = 0 Then
        GetPricesJava = 0
        Exit Function
    End If
    
    Dim Xa As Long
    Xa = XaCantidad
    
    'voy bajando hasta encontrar un precio
    For Xa = XaCantidad To 1 Step -1
        'ver que mida si da la matriz
        If Xa > UBound(mPricesJava) Then GoTo sig
        
        If mPricesJava(Xa) > 0 Then
            'ya se cual es el precio unitario por la cantidad que elegi
            'ahora veo el precio que le hago por todo
            Dim SubTot As Single
            SubTot = (mPricesJava(Xa) / Xa)
            SubTot = SubTot * XaCantidad
            'ahora redondeo para no andar con faking monedas
            GetPricesJava = CLng(SubTot)
            Exit Function
        End If
sig:
    Next Xa
    
    GetPricesJava = 0 'ES GRATIS!
End Function

Public Function GetPricesIso(XaCantidad As Long) As Long 'mm91
    'busca el precio para la cantidad indicada, si no existe igual baja hasta encontrar
    
    If XaCantidad = 0 Then
        GetPricesIso = 0
        Exit Function
    End If
    
    Dim Xa As Long
    Xa = XaCantidad
    
    'voy bajando hasta encontrar un precio
    For Xa = XaCantidad To 1 Step -1
        'ver que mida si da la matriz
        If Xa > UBound(mPricesISO) Then GoTo sig
        
        If mPricesISO(Xa) > 0 Then
            'ya se cual es el precio unitario por la cantidad que elegi
            'ahora veo el precio que le hago por todo
            Dim SubTot As Single
            SubTot = (mPricesISO(Xa) / Xa)
            SubTot = SubTot * XaCantidad
            'ahora redondeo para no andar con faking monedas
            GetPricesIso = CLng(SubTot)
            Exit Function
        End If
sig:
    Next Xa
    
    GetPricesIso = 0 'ES GRATIS!
End Function

'mp01
Public Function GetPricesIsoDVD(XaCantidad As Long) As Long 'mm91
    'busca el precio para la cantidad indicada, si no existe igual baja hasta encontrar
    
    If XaCantidad = 0 Then
        GetPricesIsoDVD = 0
        Exit Function
    End If
    
    Dim Xa As Long
    Xa = XaCantidad
    
    'voy bajando hasta encontrar un precio
    For Xa = XaCantidad To 1 Step -1
        'ver que mida si da la matriz
        If Xa > UBound(mPricesIsoDVD) Then GoTo sig
        
        If mPricesIsoDVD(Xa) > 0 Then
            'ya se cual es el precio unitario por la cantidad que elegi
            'ahora veo el precio que le hago por todo
            Dim SubTot As Single
            SubTot = (mPricesIsoDVD(Xa) / Xa)
            SubTot = SubTot * XaCantidad
            'ahora redondeo para no andar con faking monedas
            GetPricesIsoDVD = CLng(SubTot)
            Exit Function
        End If
sig:
    Next Xa
    
    GetPricesIsoDVD = 0 'ES GRATIS!
End Function


Public Function GetPrices3GP(XaCantidad As Long) As Long 'mm91
    'busca el precio para la cantidad indicada, si no existe igual baja hasta encontrar
    
    If XaCantidad = 0 Then
        GetPrices3GP = 0
        Exit Function
    End If
    
    Dim Xa As Long
    Xa = XaCantidad
    
    'voy bajando hasta encontrar un precio
    For Xa = XaCantidad To 1 Step -1
        'ver que mida si da la matriz
        If Xa > UBound(mPrices3GP) Then GoTo sig
        
        If mPrices3GP(Xa) > 0 Then
            'ya se cual es el precio unitario por la cantidad que elegi
            'ahora veo el precio que le hago por todo
            Dim SubTot As Single
            SubTot = (mPrices3GP(Xa) / Xa)
            SubTot = SubTot * XaCantidad
            'ahora redondeo para no andar con faking monedas
            GetPrices3GP = CLng(SubTot)
            Exit Function
        End If
sig:
    Next Xa
    
    GetPrices3GP = 0 'ES GRATIS!
End Function

Public Function GetPricesThemes(XaCantidad As Long) As Long 'mm91
    'busca el precio para la cantidad indicada, si no existe igual baja hasta encontrar
    
    If XaCantidad = 0 Then
        GetPricesThemes = 0
        Exit Function
    End If
    
    Dim Xa As Long
    Xa = XaCantidad
    
    'voy bajando hasta encontrar un precio
    For Xa = XaCantidad To 1 Step -1
        'ver que mida si da la matriz
        If Xa > UBound(mPricesThemes) Then GoTo sig
        
        If mPricesThemes(Xa) > 0 Then
            'ya se cual es el precio unitario por la cantidad que elegi
            'ahora veo el precio que le hago por todo
            Dim SubTot As Single
            SubTot = (mPricesThemes(Xa) / Xa)
            SubTot = SubTot * XaCantidad
            'ahora redondeo para no andar con faking monedas
            GetPricesThemes = CLng(SubTot)
            Exit Function
        End If
sig:
    Next Xa
    
    GetPricesThemes = 0 'ES GRATIS!
End Function

Public Sub CleanPrices()
    
    ReDim mPricesAudio(0)
    ReDim mPricesVideo(0)
    ReDim mPricesRingtones(0)
    ReDim mPricesWallpapers(0)
    ReDim mPricesJava(0)
    ReDim mPricesISO(0)
    ReDim mPrices3GP(0)
    ReDim mPricesThemes(0)
    ReDim mPricesIsoDVD(0) 'mp01
    CarritoChange = True
    
End Sub

'Public Property Let ModoCompra(NewModo As Long)
'    mModoCompra = NewModo
'End Property
'
'Public Property Get ModoCompra() As Long
'    ModoCompra = mModoCompra
'End Property

Public Function CalculateTotalPrice() As Single
    
    'si el carrito cambio volver a calcularlo!
    'If CarritoChange Then 'esta variable nunca vuelva a false ... digamos que es ua bosta
        Update_Tipos_Cantidad
    'End If
    
    Dim PrAu As Long 'precio del audio según la cantidad que tiene
    Dim PrVi As Long 'precio del video según la cantidad que tiene
    Dim PrRin As Long 'precio de rigtones según la cantidad que tiene
    Dim PrWal As Long 'precio del wallpapers según la cantidad que tiene
    Dim PrJa As Long 'precio del java según la cantidad que tiene
    Dim PrIso As Long 'precio segun la cantidad de imagenes ISO 'MM91
    Dim Pr3GP As Long 'precio segun la cantidad de videos 3GP
    Dim PrThemes As Long 'precio segun la cantidad de temas
    Dim PrIsoDVD As Long 'precio segun la cantidad de temas 'mp01
    
    PrAu = GetPricesAudio(mTotalAudio)
    PrVi = GetPricesVideo(mTotalVideo)
    PrRin = GetPricesRingtones(mTotalRingtones)
    PrWal = GetPricesWallpapers(mTotalWallpapers)
    PrJa = GetPricesJava(mTotalJava)
    PrIso = GetPricesIso(mTotalIso) 'mm91
    Pr3GP = GetPrices3GP(mTotal3GP) 'mm91
    PrThemes = GetPricesThemes(mTotalThemes) 'mm91
    PrIsoDVD = GetPricesIsoDVD(mTotalIsoDVD) 'mp01
    
    CalculateTotalPrice = PrAu + PrVi + PrRin + PrWal + PrJa + PrIso + Pr3GP + PrThemes + PrIsoDVD 'esta expresado en creditos Mp01
End Function

Public Function CleanFile(I As Long) As Long
    'limpia un archivo de la lista full y achica esta matriz para no dejar lugares vacios
    If (I > UBound(mListaMMFull)) Or (I = 0) Then
        Exit Function
    End If
    
    Dim H As Long
    mListaMMFull(I).cPath = ""
    'si es el ultimo es una cosa
    If I = UBound(mListaMMFull) Then
        ReDim Preserve mListaMMFull(I - 1)
    Else 'si no corro todo para abajo
        For H = I To UBound(mListaMMFull) - 1
            mListaMMFull(H) = mListaMMFull(H + 1)
        Next H
        
        ReDim Preserve mListaMMFull(UBound(mListaMMFull) - 1)
    End If
    
    CleanFile = 0
End Function

Public Function CleanEmptyFolder()
    'ver integridad de carpetas elegidas con canciones
    'ver si las carpetas que hayan sido elegidas tengan al menos una cancion
    'esto porque puede ser que se borre el unico archivo de una carpeta
    
    Dim H As Long
    
    'como se cambia el fin del for tengo un do-loop
    Do
INIAGAIN:
        For H = 1 To UBound(mListaMM)
            T = mListaMM(H).cPath
            Dim Dep As Long
            Dim Z As Long
            If mListaMM(H).cFolder Then
            'If Right(T, 1) = "\" Then
                Dep = GetDependent(H)
                If Dep = 0 Then 'carpeta sin nada!
                    mListaMM(H).cPath = ""  'marco para borrar
                    'BORRAR Y RETRAE LA MATRIZ EN 1
                    'si es el ultimo es una cosa
                    If H < UBound(mListaMM) Then
                        'corro todo para abajo
                        For Z = H To UBound(mListaMM) - 1
                            mListaMM(Z) = mListaMM(Z + 1)
                        Next Z
                    End If
                    ReDim Preserve mListaMM(UBound(mListaMM) - 1)
                    GoTo INIAGAIN
                End If
            Else
                'si es un mp3 y se borro su correspondiente en listaFull lo saco tambien!
                Dim Esta As Boolean
                Esta = False
                For Z = 1 To UBound(mListaMMFull)
                    If T = mListaMMFull(Z).cPath Then
                        Esta = True
                        Exit For
                    End If
                Next Z
                
                If Esta = False Then 'lo saco de la lista basica tambien!!
                    If H < UBound(mListaMM) Then
                        'corro todo para abajo
                        For Z = H To UBound(mListaMM) - 1
                            mListaMM(Z) = mListaMM(Z + 1)
                        Next Z
                    End If
                    ReDim Preserve mListaMM(UBound(mListaMM) - 1)
                    GoTo INIAGAIN
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
            If mListaMM(H).cPath = "" Then
                CleanSel H
                limpio = True
                Exit For
            End If
       Next H
    'Loop
    
    CarritoChange = True
    
End Function

Public Function CleanFileSoloMarca(I As Long) As Long
    'marcar para eliminar
    If (I > UBound(mListaMMFull)) Or (I = 0) Then
        Exit Function
    End If
    
    CarritoChange = True
    
    mListaMMFull(I).cPath = ""
    
End Function

Public Function CleanMarcados() As Long
    'una vez que se marcaron varios para eliminar este los quita prolijamente de la matriz
    Dim H As Long
    Dim limpio As Boolean
    Do
        limpio = False
        For H = 1 To UBound(mListaMMFull)
             If mListaMMFull(H).cPath = "" Then
                 CleanFile H
                 limpio = True
                 Exit For
             End If
        Next H
        If limpio = False Then Exit Do
    Loop
    CarritoChange = True
End Function

Public Function CleanSel(I As Long) As Long
    'borrar de lista de selecciones
    'opcionalmente eliminar en cascada
    
    If (I > UBound(mListaMM)) Or (I = 0) Then
        Exit Function
    End If
    
    Dim T As String, H As Long
    T = mListaMM(I).cPath
    If mListaMM(I).cFolder Then
    'If Right(T, 1) = "\" Then
        'eliminar sus dependencias
        Dim EP As String
        For H = 1 To UBound(mListaMMFull)
            EP = fso.GetParentFolderName(GetElementFull(H)) + "\"
            If EP = mListaMM(I).cPath Then
                'si es parte de esa carpeta lo marco para eliminar
                mListaMMFull(H).cPath = ""
            End If
        Next H
        'aqui se eliminan todos juntos losmarcados
        CleanMarcados
    Else
        'eliminar normal
        'ver que numero tiene en la matriz full para usar CleanFile
        Dim Y As Long
        For Y = 1 To UBound(mListaMMFull)
            If mListaMMFull(Y).cPath = mListaMM(I).cPath Then
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

Public Function GetDependent(I As Long) As Long
    'saber cuantos archivos estan enganchados a una carpeta
    If (I > UBound(mListaMM)) Or (I = 0) Then
        Exit Function
    End If
    
    Dim D As Long
    D = 0 'contador temporal
    'ver si es una cancion o una carpeta
    Dim T As String
    T = mListaMM(I).cPath
    If mListaMM(I).cFolder Then
    'If Right(T, 1) = "\" Then
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

Public Function GetTotalRingtones() As Long
    GetTotalRingtones = mTotalRingtones
End Function

Public Function GetTotalWallpapers() As Long
    GetTotalWallpapers = mTotalWallpapers
End Function

Public Function GetTotalJava() As Long
    GetTotalJava = mTotalJava
End Function
'mm91
Public Function GetTotalISO() As Long
    GetTotalISO = mTotalIso
End Function
'mp01
Public Function GetTotalISODVD() As Long
    GetTotalISODVD = mTotalIsoDVD
End Function

'mm91
Public Function GetTotal3GP() As Long
    GetTotal3GP = mTotal3GP
End Function

'mm91
Public Function GetTotalThemes() As Long
    GetTotalThemes = mTotalThemes
End Function

Public Function AddFile(nFile As String, Optional fromFolder As Boolean = False, Optional perfil = 1) As Long
    Dim H As Long
    
    'ver si ya existe!!!
    If ExisteFileInCart(nFile) = False Then
        'estaba hecho para venir desde afuera pero en AddFolder se copiaba el codigo pero si
        'agregar a mListaMM. Para no repetir agrege un parametro opcional
        
        If fromFolder = False Then
            H = UBound(mListaMM) + 1
            ReDim Preserve mListaMM(H)
            mListaMM(H).cPath = nFile
            mListaMM(H).cFolder = 0 'es un archivo comun
            mListaMM(H).cPerfil = perfil
        End If
        
        H = UBound(mListaMMFull) + 1
        ReDim Preserve mListaMMFull(H)
        mListaMMFull(H).cPath = nFile
        mListaMMFull(H).cPerfil = perfil
        
        'agrega al contador de tipos de archivos
        miniUpTotalTypeFile nFile, H
        
        CarritoChange = True
    Else
        AddFile = 1 'YA ESTABA!
        Exit Function
    End If
    
    AddFile = 0
End Function

Public Function AddFolder(nFolder As String, Optional PerfilPreferido As Long = 1) As Long
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
    'OM- se agrega una carpeta al carrito _
        (se deberia agregar lo mismo que el tipo ya esta viendo segun perfil definido antes)
    'quizas un parametro en addfolder seria solucion
    T = ObtenerArchMM(nFolder, True, PerfilPreferido)
    Dim J As Long
    
    Dim CantSel As Long
    CantSel = 0 'me fijo cuantos realmente no estaban para ver si se cargo todo o ninguno
    Dim nFile As String
    For H = 1 To UBound(T)
        nFile = txtInLista(T(H), 0, "#")
        'nombreTemas = txtInLista(MATRIZ_TEMAS(c), 1, "#")
        'ME ASEGURO QUE NO ESTE ELGIDO YA
        If AddFile(nFile, True, PerfilPreferido) = 0 Then 'devuelve cero si no estaba ya agregado y 1 si es un repetido
            CantSel = CantSel + 1
        End If
    Next H
    
    AddFolder = 3
    If CantSel = 0 Then
        'ya estaban TODAS de antes, no debo agregar nada, me voy .... chauuuuuuuuuuuu
        Exit Function
    End If
    
    If CantSel < UBound(T) Then
        'deberia avisar que al menos una cancion ya estaba
    End If
    
    H = UBound(mListaMM) + 1
    ReDim Preserve mListaMM(H)
    mListaMM(H).cPath = nFolder  'como termina en "\" sabre despues que es una carpeta
    mListaMM(H).cFolder = 1 'indico que es una carpeta
    mListaMM(H).cPerfil = PerfilPreferido
    
    AddFolder = 0
    CarritoChange = True
    Exit Function
ER:
     
End Function

Public Function GetElement(I As Long) As String
    If (I > UBound(mListaMM)) Or (I = 0) Then
        GetElement = ""
        Exit Function
    End If
    
    GetElement = mListaMM(I).cPath
End Function

Public Function GetPricesAudioBase(I As Long) As Long
    If (I > UBound(mPricesAudio)) Or (I = 0) Then
        GetPricesAudioBase = 0
        Exit Function
    End If
    
    GetPricesAudioBase = mPricesAudio(I)
End Function

Public Function KillPricesAudioBase(I As Long) As Long
    If (I > UBound(mPricesAudio)) Or (I = 0) Then
        KillPricesAudioBase = -1
        Exit Function
    End If
    
    mPricesAudio(I) = 0
    KillPricesAudioBase = 0
    
    CarritoChange = True
End Function

Public Function KillPricesVideoBase(I As Long) As Long
    If (I > UBound(mPricesVideo)) Or (I = 0) Then
        KillPricesVideoBase = -1
        Exit Function
    End If
    
    mPricesVideo(I) = 0
    KillPricesVideoBase = 0
    
    CarritoChange = True
End Function

Public Function GetPricesVideoBase(I As Long) As Long
    If (I > UBound(mPricesVideo)) Or (I = 0) Then
        GetPricesVideoBase = 0
        Exit Function
    End If
    
    GetPricesVideoBase = mPricesVideo(I)
End Function

Public Function GetPricesRingtonesBase(I As Long) As Long
    If (I > UBound(mPricesRingtones)) Or (I = 0) Then
        GetPricesRingtonesBase = 0
        Exit Function
    End If
    
    GetPricesRingtonesBase = mPricesRingtones(I)
End Function

Public Function KillPricesRingtonesBase(I As Long) As Long
    If (I > UBound(mPricesRingtones)) Or (I = 0) Then
        KillPricesRingtonesBase = -1
        Exit Function
    End If
    
    mPricesRingtones(I) = 0
    KillPricesRingtonesBase = 0
    
    CarritoChange = True
End Function

Public Function GetPricesWallpapersBase(I As Long) As Long
    If (I > UBound(mPricesWallpapers)) Or (I = 0) Then
        GetPricesWallpapersBase = 0
        Exit Function
    End If
    
    GetPricesWallpapersBase = mPricesWallpapers(I)
End Function

Public Function KillPricesWallpapersBase(I As Long) As Long
    If (I > UBound(mPricesWallpapers)) Or (I = 0) Then
        KillPricesWallpapersBase = -1
        Exit Function
    End If
    
    mPricesWallpapers(I) = 0
    KillPricesWallpapersBase = 0
    
    CarritoChange = True
End Function

Public Function GetPricesJAVABase(I As Long) As Long
    If (I > UBound(mPricesJava)) Or (I = 0) Then
        GetPricesJAVABase = 0
        Exit Function
    End If
    
    GetPricesJAVABase = mPricesJava(I)
End Function

'mm91
Public Function GetPricesISOBase(I As Long) As Long
    If (I > UBound(mPricesISO)) Or (I = 0) Then
        GetPricesISOBase = 0
        Exit Function
    End If
    
    GetPricesISOBase = mPricesISO(I)
End Function

Public Function KillPricesISOBase(I As Long) As Long
    If (I > UBound(mPricesISO)) Or (I = 0) Then
        KillPricesISOBase = -1
        Exit Function
    End If
    
    mPricesISO(I) = 0
    KillPricesISOBase = 0
    
    CarritoChange = True
End Function

'mp01
Public Function GetPricesISODVDBase(I As Long) As Long
    If (I > UBound(mPricesIsoDVD)) Or (I = 0) Then
        GetPricesISODVDBase = 0
        Exit Function
    End If
    
    GetPricesISODVDBase = mPricesIsoDVD(I)
End Function
'mp01
Public Function KillPricesISODVDBase(I As Long) As Long
    If (I > UBound(mPricesIsoDVD)) Or (I = 0) Then
        KillPricesISODVDBase = -1
        Exit Function
    End If
    
    mPricesIsoDVD(I) = 0
    KillPricesISODVDBase = 0
    
    CarritoChange = True
End Function

'mm91
Public Function GetPrices3GPBase(I As Long) As Long
    If (I > UBound(mPrices3GP)) Or (I = 0) Then
        GetPrices3GPBase = 0
        Exit Function
    End If
    
    GetPrices3GPBase = mPrices3GP(I)
End Function

'mm91
Public Function GetPricesThemesBase(I As Long) As Long
    If (I > UBound(mPricesThemes)) Or (I = 0) Then
        GetPricesThemesBase = 0
        Exit Function
    End If
    
    GetPricesThemesBase = mPricesThemes(I)
End Function

Public Function KillPricesJAVABase(I As Long) As Long
    If (I > UBound(mPricesJava)) Or (I = 0) Then
        KillPricesJAVABase = -1
        Exit Function
    End If
    
    mPricesJava(I) = 0
    KillPricesJAVABase = 0
    
    CarritoChange = True
End Function

'mm91
Public Function KillPrices3GPBase(I As Long) As Long
    If (I > UBound(mPrices3GP)) Or (I = 0) Then
        KillPrices3GPBase = -1
        Exit Function
    End If
    
    mPrices3GP(I) = 0
    KillPrices3GPBase = 0
    
    CarritoChange = True
End Function

'mm91
Public Function KillPricesThemesBase(I As Long) As Long
    If (I > UBound(mPricesThemes)) Or (I = 0) Then
        KillPricesThemesBase = -1
        Exit Function
    End If
    
    mPricesThemes(I) = 0
    KillPricesThemesBase = 0
    
    CarritoChange = True
End Function

Public Function GetTotalPricesAudio() As Long
    GetTotalPricesAudio = UBound(mPricesAudio)
End Function

Public Function GetTotalPricesVideo() As Long
    GetTotalPricesVideo = UBound(mPricesVideo)
End Function

Public Function GetTotalPricesRingtones() As Long
    GetTotalPricesRingtones = UBound(mPricesRingtones)
End Function

Public Function GetTotalPricesWallpapers() As Long
    GetTotalPricesWallpapers = UBound(mPricesWallpapers)
End Function

Public Function GetTotalPricesJAVA() As Long
    GetTotalPricesJAVA = UBound(mPricesJava)
End Function

Public Function GetTotalPricesISO() As Long 'mm91
    GetTotalPricesISO = UBound(mPricesISO)
End Function

Public Function GetTotalPricesISODVD() As Long 'mp01
    GetTotalPricesISODVD = UBound(mPricesIsoDVD)
End Function

Public Function GetTotalPrices3GP() As Long 'mm91
    GetTotalPrices3GP = UBound(mPrices3GP)
End Function

Public Function GetTotalPricesThemes() As Long 'mm91
    GetTotalPricesThemes = UBound(mPricesThemes)
End Function

Public Function GetElementFull(I As Long) As String
    If (I > UBound(mListaMMFull)) Or (I = 0) Then
        GetElementFull = ""
        Exit Function
    End If
    
    GetElementFull = mListaMMFull(I).cPath
End Function

Public Function GetElementPath(I As Long) As String
    If (I > UBound(mListaMM)) Or (I = 0) Then
        GetElementPath = ""
        Exit Function
    End If
    
    'ver si es una cancion o una carpeta
    Dim T As String
    If mListaMM(I).cFolder Then
    'If Right(mListaMM(I).cPath, 1) = "\" Then
        T = mListaMM(I).cPath
    Else
        T = fso.GetParentFolderName(mListaMM(I).cPath)
    End If
    
    If Right(T, 1) <> "\" Then T = T + "\"
    
    GetElementPath = T
End Function

Public Function GetElementFullListPath(I As Long) As String
    If (I > UBound(mListaMMFull)) Or (I = 0) Then
        GetElementFullListPath = ""
        Exit Function
    End If
    
    'es una cancion si o si
    T = fso.GetParentFolderName(mListaMMFull(I).cPath)
    
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
            If fso.FileExists(mListaMMFull(H).cPath) Then
                TMP = TMP + ((fso.GetFile(mListaMMFull(H).cPath).Size / 1024) / 1024)
            End If
        Next H
        mGetTotalMB = Round(TMP, 2)
        GetTotalMB = mGetTotalMB
    Else
        'ver valor anterior
        GetTotalMB = mGetTotalMB
    End If
End Function

'SOLO LOS mp3s!!! , por que les puedo leer la info mas o menos facil
Public Function GetTotalMinutos() As Single
    'calcula el peso de todos los archivos
    
    Dim TMP As Single: TMP = 0
    
    'PODRIA VER SI NO HUBO CAMBIOS Y NO LEER TODO DE NUEVO
    If CarritoChange = True Then
        
        Dim MP3tmp As New MP3Info, Secs As Long
        Dim H As Long, tExt As String
        
        For H = 1 To UBound(mListaMMFull)
            tExt = LCase(fso.GetExtensionName(mListaMMFull(H).cPath))
            If tExt = "mp3" Then
                MP3tmp.FileName = mListaMMFull(H).cPath
                Secs = MP3tmp.Seconds
                TMP = TMP + CLng(Secs)
            End If
        Next H
        mGetTotalMinutos = Round(TMP / 60, 2)
        GetTotalMinutos = mGetTotalMinutos
    Else
        'ver valor anterior
        GetTotalMinutos = mGetTotalMinutos
    End If
    
    
End Function

Public Function GetElementName(I As Long) As String
    If (I > UBound(mListaMM)) Or (I = 0) Then
        GetElementName = ""
        Exit Function
    End If
    
    Dim T As String, T2 As String
    
    'ver si es una cancion o una carpeta
    If mListaMM(I).cFolder Then
    'If Right(mListaMM(I).cPath, 1) = "\" Then
        T = TR.Trad("Completo:%99%") + " " + fso.GetBaseName(mListaMM(I).cPath)
        GetElementName = T
    Else
        T = fso.GetBaseName(fso.GetParentFolderName(mListaMM(I).cPath))
        T2 = fso.GetBaseName(mListaMM(I).cPath)
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
    mTotalRingtones = 0
    mTotalWallpapers = 0
    mTotalJava = 0
    mTotalIso = 0 'mm91
    mTotal3GP = 0 'mm91
    mTotalThemes = 0
    mTotalIsoDVD = 0 'mp01
    
    mArchivo = "c:\cart.txt" 'y si no hay "c:\" YA ME PASO!
    'mModoCompra = 0
    CleanPrices
    mGetTotalMinutos = 0
    mGetTotalMB = 0
    
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
    Dim SP2() As String
    For H = 0 To UBound(SP)
        SP2 = Split(SP(H), Chr(5))
    
        ReDim Preserve mListaMM(H + 1) 'la matriz inicia en 1 siempre
        mListaMM(H + 1).cPath = SP2(0)
        
        'debo escribir y leer el perfil para que no se mezclen los ringtones con los MP3s
        mListaMM(H + 1).cPerfil = SP2(1)
        
        'ver que tipo es (para cFolder)
        If Right(mListaMM(H + 1).cPath, 1) = "\" Then
            mListaMM(H + 1).cFolder = 1
        Else
            mListaMM(H + 1).cFolder = 0
        End If
    Next H
    
    LoadCartFromDisk = 0
    
    CarritoChange = True
    
    Exit Function
ER:

    tERR.AppendLog "lcfdsk99", tERR.ErrToTXT(Err)
    ClearCart 'lo limpio por las dudas
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
            TE.Write mListaMM(H).cPath + Chr(5) + CStr(mListaMM(H).cPerfil)
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
    mTotalRingtones = 0
    mTotalWallpapers = 0
    mTotalJava = 0
    mTotalIso = 0 'mm91
    mTotalIsoDVD = 0 'mp01
    mTotal3GP = 0 'mm91
    mTotalThemes = 0
    mGetTotalMinutos = 0
    
    CarritoChange = True
    
    mGetTotalMB = 0
End Function

Public Function ExisteFileInCart(nFile As String) As Boolean
    'ver si hay algo!
    Dim H As Long, B As Boolean
    B = False
    
    For H = 1 To UBound(mListaMMFull)
        If LCase(mListaMMFull(H).cPath) = LCase(nFile) Then
            B = True
            Exit For
        End If
    Next H
    
    ExisteFileInCart = B
End Function

Public Function SavePrices(sFile As String)

    If fso.FileExists(sFile) Then fso.DeleteFile sFile
    Dim TePC As TextStream, H As Long
    
    'A|Audio                 'mp01
    'V|Video
    'R|Ringtones
    'W|Wallpaper
    'J|Juego java
    'I|Iso cd
    'D|Iso dvd
    'G|video 3gp
    'T|Tema cel
    
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
        
        For H = 1 To GetTotalPricesRingtones
            If GetPricesRingtonesBase(H) > 0 Then
                TePC.Write "R|" + CStr(H) + "|" + CStr(GetPricesRingtonesBase(H)) + Chr(6)
            End If
        Next H
        
        For H = 1 To GetTotalPricesWallpapers
            If GetPricesWallpapersBase(H) > 0 Then
                TePC.Write "W|" + CStr(H) + "|" + CStr(GetPricesWallpapersBase(H)) + Chr(6)
            End If
        Next H
        
        For H = 1 To GetTotalPricesJAVA
            If GetPricesJAVABase(H) > 0 Then
                TePC.Write "J|" + CStr(H) + "|" + CStr(GetPricesJAVABase(H)) + Chr(6)
            End If
        Next H
        
        For H = 1 To GetTotalPricesISO
            If GetPricesISOBase(H) > 0 Then
                TePC.Write "I|" + CStr(H) + "|" + CStr(GetPricesISOBase(H)) + Chr(6)
            End If
        Next H
        
        For H = 1 To GetTotalPricesISODVD 'mp01
            If GetPricesISODVDBase(H) > 0 Then
                TePC.Write "D|" + CStr(H) + "|" + CStr(GetPricesISODVDBase(H)) + Chr(6)
            End If
        Next H
        
        For H = 1 To GetTotalPrices3GP
            If GetPrices3GPBase(H) > 0 Then
                TePC.Write "G|" + CStr(H) + "|" + CStr(GetPrices3GPBase(H)) + Chr(6)
            End If
        Next H
        
        For H = 1 To GetTotalPricesThemes
            If GetPricesThemesBase(H) > 0 Then
                TePC.Write "T|" + CStr(H) + "|" + CStr(GetPricesThemesBase(H)) + Chr(6)
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
            TePC.Write "R|1|1" + Chr(6) 'ringtones
            TePC.Write "W|1|2" + Chr(6) 'wallpapers
            TePC.Write "J|1|3" + Chr(6) 'java
            TePC.Write "I|1|12" + Chr(6) 'iso cd
            TePC.Write "D|1|16" + Chr(6) 'iso dvd             'mp01
            TePC.Write "G|1|3" + Chr(6) '3gp
            TePC.Write "T|1|2" + Chr(6) 'temas
            
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
                tERR.Anotar "dacc2", SP9(0), SP9(1), SP9(2)
                Select Case UCase(SP9(0))
                    'musica
                    Case "A": SetPricesAudio CLng(SP9(1)), CLng(SP9(2))
                    'videillo
                    Case "V": SetPricesVideo CLng(SP9(1)), CLng(SP9(2))
                    'ringtonsillo
                    Case "R": SetPricesRingtones CLng(SP9(1)), CLng(SP9(2))
                    'wallpapersillo
                    Case "W": SetPricesWallpapers CLng(SP9(1)), CLng(SP9(2))
                    'javillo
                    Case "J": SetPricesJava CLng(SP9(1)), CLng(SP9(2))
                    'imagenes de disco ISO 'mm91
                    Case "I": SetPricesIso CLng(SP9(1)), CLng(SP9(2))
                    'videos para movil 'mm91
                    Case "G": SetPrices3GP CLng(SP9(1)), CLng(SP9(2))
                    'videos para movil 'mm91
                    Case "T": SetPricesThemes CLng(SP9(1)), CLng(SP9(2))
                    'mp01
                    Case "D": SetPricesIsoDVD CLng(SP9(1)), CLng(SP9(2))
                    
                    Case Else: tERR.AppendLog "FuckCart:" + SP9(0)
                End Select
            End If
sig:
        Next H
        'si no salio un puto precio ok entonces va de nuevo
        'es casi imposible despues de tantas verificaciones
    End If

    'revisar que al menos haya un precio de cada uno!!!
    'por ejemplo alguien que tenga 3pm 7.20.x puede tener precios de audio y video pero no de los otros
    'el archivo ya existe por lo que no lo inicializa desde cero...
    'por las dudas hago una revision.... PROBADA OK
    If UBound(mPricesAudio) = 0 Then SetPricesAudio 1, 2
    If UBound(mPricesVideo) = 0 Then SetPricesVideo 1, 4
    If UBound(mPricesRingtones) = 0 Then SetPricesRingtones 1, 1
    If UBound(mPricesWallpapers) = 0 Then SetPricesWallpapers 1, 2
    If UBound(mPricesJava) = 0 Then SetPricesJava 1, 3
    If UBound(mPricesISO) = 0 Then SetPricesIso 1, 12
    If UBound(mPrices3GP) = 0 Then SetPrices3GP 1, 4
    If UBound(mPricesThemes) = 0 Then SetPricesThemes 1, 4
    If UBound(mPricesIsoDVD) = 0 Then SetPricesIsoDVD 1, 16 'mp01

    CarritoChange = True
End Function

Public Function Update_Tipos_Cantidad() As Single
    'cuenta cuantos archivos de audio y video hay (cuando se borran no se descuentan)
    Dim H As Long
    mTotalAudio = 0
    mTotalVideo = 0
    mTotalRingtones = 0
    mTotalWallpapers = 0
    mTotalJava = 0
    mTotalIso = 0 'mm91
    mTotal3GP = 0 'mm91
    mTotalThemes = 0
    mTotalIsoDVD = 0 'mp01
    
    Dim Extranos As Long
    
    For H = 1 To UBound(mListaMMFull)
        Extranos = Extranos + miniUpTotalTypeFile(mListaMMFull(H).cPath, H)
    Next H
    
    Update_Tipos_Cantidad = Extranos 'deberia ser cero
End Function

Private Function miniUpTotalTypeFile(ptFile As String, indexMtxFull As Long) As Long
    'llena todo pero devuleve la cantida de objetos extraños incluidos
    'pido el indice a partir de ahora (ago08) para saber su perfil y separar ringtones de mp3s
    
    Dim Extranos As Long
    Extranos = 0
    Select Case LCase(fso.GetExtensionName(ptFile))
        Case "mp3" 'condicional a que pueda ser ringtone o cancion comun MP3
            'XXXX el raking deberia ser uno por cada tipo de contenido! por ahora solo musica!!
            If mListaMMFull(indexMtxFull).cPerfil = -1 Then mTotalAudio = mTotalAudio + 1
            
            If mListaMMFull(indexMtxFull).cPerfil = 0 Then mTotalAudio = mTotalAudio + 1
            If mListaMMFull(indexMtxFull).cPerfil = 1 Then mTotalAudio = mTotalAudio + 1
            If mListaMMFull(indexMtxFull).cPerfil = 2 Then mTotalRingtones = mTotalRingtones + 1
        Case "wma"
            mTotalAudio = mTotalAudio + 1
        Case "mpeg", "mpg", "avi", "wmv", "vob", "dat"
            mTotalVideo = mTotalVideo + 1
        Case "mn0", "mn1"
            Extranos = 1
        Case "jpg", "jpeg", "bmp", "gif"
            mTotalWallpapers = mTotalWallpapers + 1
        Case "jar" ', "jad"
            mTotalJava = mTotalJava + 1
        'mp01
        Case "iso", "nrg", "nr3", "nra", "nrb", "nrc", "nrd", "nre", "nrh", "nri", "nrm", "nru", "nrv", "nrw" 'mm91
            'DIFERENCIAR LAS IMÄGENES DE CD DE LAS DE DVD!!!
            If fso.GetFile(ptFile).Size > 734003200 Then '700 x 1024 x 1024
                mTotalIsoDVD = mTotalIsoDVD + 1
            Else
                mTotalIso = mTotalIso + 1
            End If
        Case "3gp" 'mm91 videos para celulares
            mTotal3GP = mTotal3GP + 1
        Case "nth", "thm" 'nokia theme + sony theme
            mTotalThemes = mTotalThemes + 1
        Case Else
            Extranos = 1
    End Select
End Function

'dice si es 1 carrito con solo 1 archivo de imagen de disco
Public Function isISO() As Boolean
    Dim isISO2 As Boolean
    If GetFileCantFull = 1 Then 'mp01 'DECIA CARRITO !!!
        Select Case LCase(fso.GetExtensionName(GetElementFull(1)))
            'formatos de imagenes de nero
            'NR3: cd de mp3s    /    'NRA: cd de audio    /  'NRB: cd-rom de arranque
            'NRC: nero usf/iso  /    'NRD: nero DVD       /  'NRE: cd extra
            'NRG: imagen        /    'NRH: cd-rom hibrido /  'NRI: cd-rom iso
            'NRM: cd mixto      /    'NRU: cd-rom udf     /  'NRV: cd supervideo
            'NRW: cd rom wma    /    'CDC: cd cover no tiene nada que ver con imagenes parece
            Case "iso", "nrg", "nr3", "nra", "nrb", "nrc", "nrd", "nre", "nrh", "nri", "nrm", "nru", "nrv", "nrw"
                isISO2 = True
            Case Else
                isISO2 = False
        End Select
    Else
        isISO2 = False 'se supone que lo deje llegar aqui con varios archivos si lo que queria es grabar una imagen iso
    End If
    
    isISO = isISO2
End Function

' se fija en que origen de discos esta cada uno de los elementos y hace una matriz con ellos
' lo uso por ahora para ver si un contenido es 100% de karaokes grabados y por lo tanto
' no exigir licencia de carrito
Public Function GetOrigenesContenido() As String()
    Dim H As Long, Oris() As String
    ReDim Oris(0) 'empiezo en base 1
    For H = 1 To UBound(mListaMMFull)
        'ver el origen actual
        Dim OriEste As String, tTmp() As String
        tTmp = Split(mListaMMFull(H).cPath, "\")
        OriEste = LCase(tTmp(UBound(tTmp) - 2)) 'ttmp(UBound(ttmp) es el archivo -- ttmp(UBound(ttmp)-1) es el disco -- ttmp(UBound(ttmp)-2) es el origen
            
        'ver si ya hay para comparar y no agregar repetidos
        If UBound(Oris) = 0 Then
            ReDim Oris(1)
            Oris(1) = OriEste
        Else
            Dim H2 As Long, repe As Boolean
            repe = False
            For H2 = 1 To UBound(Oris)
                If Oris(H2) = OriEste Then repe = True
            Next H2
            If repe = False Then
                H2 = UBound(Oris) + 1
                ReDim Preserve Oris(H2)
                Oris(H2) = OriEste
            End If
        End If
    Next H

    GetOrigenesContenido = Oris

End Function

'aplicacion de la funciona anterior para el caso especifico de karaokes grabados
Public Function solo1Origen(verif As String) As Boolean
    Dim vf() As String
    vf = GetOrigenesContenido
    
    solo1Origen = (UBound(vf) = 1) And (LCase(vf(1)) = LCase(verif))

End Function
