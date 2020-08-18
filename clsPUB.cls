VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPUB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'modulo para las publicidades
'la onda es que cada vez que se cargue un tema se cargue la publicidad correspondiente
'es importante que en la lista de temas pendientes no se muestren!!!!

Private mHabilitarPublicidades As Boolean 'hay o no pub mpe o video?
Private mHabilitarPublicidadesIMG As Boolean 'hay o no pub (imagenes rotativas)?

Private mSonarPublicidadesCada As Long 'cantidad de temas que deben pasar para
'que suene una publicidad
Private mSonarPublicidadesIMGCada As Long 'cantidad de segundos que deben pasar para que cambie
' la publicidad mostrada en index

Private mContadorTemas As Long 'contador de temas reproducidos. Se pone en cero cuando se pasa

'una publicidad en MP3 o AVI !!
Private mPubs() As String 'matriz con la lista de publicidades
Private mPubsSORT() As String 'matriz con la lista de publicidades ORDENADAS!!!
'una publicidad en JPG o BMP o GIF
Private mPubsIMG() As String 'matriz con la lista de publicidades
Private mPubsSORTIMG() As String 'matriz con la lista de publicidades ORDENADAS!!!

Private mUltimaReproducida As Long 'ultima publicidad reproducida para servir de guia
Private mUltimaReproducidaIMG As Long 'ultima publicidad reproducida para servir de guia

Private mPubsEnLista As Long 'publicidades en la lista de espera. Para saber cuanto
'restar del los temas pendientes reales. Como quedan en lña lsita de reproduccion el indicador
' TEMAS PENDIENTES (X) debe saber cuanto restar al toal que tiene la lista

Public Property Get HabilitarPublicidades() As Boolean
    HabilitarPublicidades = mHabilitarPublicidades
End Property

Public Property Let HabilitarPublicidades(Publicidades As Boolean)
    mHabilitarPublicidades = Publicidades
End Property

Public Property Get HabilitarPublicidadesIMG() As Boolean
    HabilitarPublicidadesIMG = mHabilitarPublicidadesIMG
End Property

Public Property Let HabilitarPublicidadesIMG(Publicidades As Boolean)
    mHabilitarPublicidadesIMG = Publicidades
End Property

Public Property Get SonarPublicidadesCada() As Long
    SonarPublicidadesCada = mSonarPublicidadesCada
End Property

Public Property Let SonarPublicidadesCada(SonarCada As Long)
    mSonarPublicidadesCada = SonarCada
End Property

Public Property Get SonarPublicidadesIMGCada() As Long
    SonarPublicidadesIMGCada = mSonarPublicidadesIMGCada
End Property

Public Property Let SonarPublicidadesIMGCada(SonarCada As Long)
    mSonarPublicidadesIMGCada = SonarCada
End Property

Public Property Get PubsEnLista() As Long
    PubsEnLista = mPubsEnLista
End Property

Public Property Let PubsEnLista(PEL As Long)
    mPubsEnLista = PEL
End Property

Public Property Get ContadorTemas() As Long
    've cuantos temas sin publicidad van. Al pasar una publicidad se pone en cero
    ContadorTemas = mContadorTemas
End Property

Public Property Let ContadorTemas(cTemas As Long)
    've cuantos temas sin publicidad van. Al pasar una publicidad se pone en cero
    mContadorTemas = cTemas
End Property

Public Function ArchsPubs(uIndex As Long) As String
    'matriz de las publicidades que hay
    'todas se reproducen alfabeticamente para que el usuario pueda tener algun control como
    'ponerle números para saber que va a salir o duplicar alguna si le pinta
    If uIndex > UBound(mPubs) Then
        ArchsPubs = ""
        'MsgBox "Boludo, me estas pidiendo una publicidad que no existe!!!"
    Else
        ArchsPubs = mPubsSORT(uIndex)
    End If
End Function

Public Function ArchsPubsIMG(uIndex As Long) As String
    'matriz de las publicidades que hay
    'todas se reproducen alfabeticamente para que el usuario pueda tener algun control como
    'ponerle números para saber que va a salir o duplicar alguna si le pinta
    If uIndex > UBound(mPubsIMG) Then
        ArchsPubsIMG = ""
        'MsgBox "Boludo, me estas pidiendo una publicidad que no existe!!!"
    Else
        ArchsPubsIMG = mPubsSORTIMG(uIndex)
    End If
End Function

Public Sub CargarPUBs()
    'cargar la lista de publicidades disponibles
    'debe estar en AP+"pub"
    ReDim mPubs(0)
    ReDim mPubsIMG(0)
    Dim CONT As Long
    CONT = 0
    Dim CONTimg As Long
    CONTimg = 0
    Dim Arch As String
    Arch = Dir(AP + "pub\*.*", vbArchive)
    Dim IsMM As Boolean
    Dim IsIMG As Boolean
    Do While Arch <> ""
        IsIMG = False
        IsMM = False
        'ver que tenga los formatos que busco
        If LCase(Right(Arch, 3)) = "mp3" Then IsMM = True
        If LCase(Right(Arch, 3)) = "avi" Then IsMM = True
        If LCase(Right(Arch, 3)) = "mpg" Then IsMM = True
        If LCase(Right(Arch, 4)) = "mpeg" Then IsMM = True
        'ver si es IMG
        If LCase(Right(Arch, 3)) = "jpg" Then IsIMG = True
        If LCase(Right(Arch, 3)) = "gif" Then IsIMG = True
        If LCase(Right(Arch, 3)) = "bmp" Then IsIMG = True
        If LCase(Right(Arch, 4)) = "jpeg" Then IsIMG = True
        
        If IsMM Then
            CONT = CONT + 1
            ReDim Preserve mPubs(CONT)
            mPubs(CONT) = UCase(AP + "pub\" + Arch)
        End If
        
        If IsIMG Then
            CONTimg = CONTimg + 1
            ReDim Preserve mPubsIMG(CONTimg)
            mPubsIMG(CONTimg) = UCase(AP + "pub\" + Arch)
        End If
        
        Arch = Dir
    Loop
    
    'ordenarlo!!!!!!!!!!
    'primero ordeno los MP3+Video
    '--------------------------------
    Dim MinSTR As String 'comparacion de cadenas. Empiezo con el máximo
    Dim ubicMIN As Long 'indice en la matriz del menor encontrado cada vuelta
    MinSTR = "zzzzzzzzzzzzzzzz"
    Dim c As Long, mtx As Long, ValComp As String
    c = 0 'cantidad de minimos encontrados
    Dim Ordenados() As Long 'matriz con los indices ordenados
    Do
        For mtx = 0 To UBound(mPubs)
            ValComp = mPubs(mtx)
            If ValComp < MinSTR Then
                MinSTR = ValComp
                ubicMIN = mtx
            End If
        Next
        'lo mando al ultimo para que no vuelva a salir como primero
        mPubs(ubicMIN) = "ZZZ" + mPubs(ubicMIN)
        ReDim Preserve Ordenados(c)
        Ordenados(c) = ubicMIN
        c = c + 1
        If c > UBound(mPubs) Then Exit Do
        MinSTR = "zzzzzzzzzz"
    Loop
    'cargar todos y sacar las primera 3 zetas
    'limpiar la matriz ordenada
    ReDim mPubsSORT(0)
    For mtx = 0 To UBound(mPubs)
        ReDim Preserve mPubsSORT(mtx)
        mPubsSORT(mtx) = Right(mPubs(Ordenados(mtx)), Len(mPubs(Ordenados(mtx))) - 3)
    Next
    '--------------------------------
    'ahora ordeno las imágenes
    'Dim MinSTR As String 'comparacion de cadenas. Empiezo con el máximo
    'Dim ubicMIN As Long 'indice en la matriz del menor encontrado cada vuelta
    MinSTR = "zzzzzzzzzzzzzzzz"
    'Dim c As Long, mtx As Long, ValComp As String
    c = 0 'cantidad de minimos encontrados
    ReDim Preserve Ordenados(0) 'Dim Ordenados() As Long 'matriz con los indices ordenados
    Do
        For mtx = 0 To UBound(mPubsIMG)
            ValComp = mPubsIMG(mtx)
            If ValComp < MinSTR Then
                MinSTR = ValComp
                ubicMIN = mtx
            End If
        Next
        'lo mando al ultimo para que no vuelva a salir como primero
        mPubsIMG(ubicMIN) = "ZZZ" + mPubsIMG(ubicMIN)
        ReDim Preserve Ordenados(c)
        Ordenados(c) = ubicMIN
        c = c + 1
        If c > UBound(mPubsIMG) Then Exit Do
        MinSTR = "zzzzzzzzzz"
    Loop
    'cargar todos y sacar las primera 3 zetas
    'limpiar la matriz ordenada
    ReDim mPubsSORTIMG(0)
    For mtx = 0 To UBound(mPubsIMG)
        ReDim Preserve mPubsSORTIMG(mtx)
        mPubsSORTIMG(mtx) = Right(mPubsIMG(Ordenados(mtx)), Len(mPubsIMG(Ordenados(mtx))) - 3)
    Next
    
End Sub

Public Property Get TotalPUBs() As Long
    'saber el total para no pasarme de largo
    'empieza en cero
    TotalPUBs = UBound(mPubs)
End Property

Public Property Get TotalPUBsIMG() As Long
    'saber el total para no pasarme de largo
    'empieza en cero
    TotalPUBsIMG = UBound(mPubsIMG)
End Property

Public Property Get UltimaReproducida() As Long
    'indica cual se uso para ayudar al sistema aguiarse
    'en realidad es la ultima que se mando a la lista. Nunca sabe cual es la ultima que se reprodujo
    UltimaReproducida = mUltimaReproducida
End Property

Public Property Get UltimaReproducidaIMG() As Long
    'indica cual se uso para ayudar al sistema aguiarse
    'en realidad es la ultima que se mando a la lista. Nunca sabe cual es la ultima que se reprodujo
    UltimaReproducidaIMG = mUltimaReproducidaIMG
End Property

Public Property Let UltimaReproducida(Ultima As Long)
    mUltimaReproducida = Ultima
End Property

Public Property Let UltimaReproducidaIMG(Ultima As Long)
    mUltimaReproducidaIMG = Ultima
End Property

Private Sub Class_Initialize()
    mUltimaReproducida = 0
    mUltimaReproducidaIMG = 0
End Sub
