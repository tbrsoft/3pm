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

Private mHabilitarPublicidadesMp3Vid As Boolean 'hay o no pub mpe o video?
Private mHabilitarPublicidadesIMG As Boolean 'hay o no pub (imagenes rotativas)?
Private mHabilitarPublicidadesTXT As Boolean 'de texto en en PassTXT
Private mHabilitarPublicidadesVMute As Boolean 'de texto en en PassTXT

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

Private mPubsBigIMG() As String 'matriz con la lista de publicidades grandes para el formulario de video
Private mPubsSORTBigIMG() As String 'matriz con la lista de publicidades ORDENADAS!!!

Private mPubsTXT() As String 'matriz con la lista de publicidades
Private mPubsSortTXT() As String 'matriz con la lista de publicidades ORDENADAS!!!

Private mPubsMute() As String 'matriz con la lista de videos en AP/pubMute
Private mPubsSortMute() As String 'matriz con la lista ORDENADA!!!

Private mUltimaReproducida As Long 'ultima publicidad reproducida para servir de guia
Private mUltimaReproducidaIMG As Long 'ultima publicidad reproducida para servir de guia
Private mUltimaReproducidaBigIMG As Long 'ultima publicidad reproducida (GRANDES) para servir de guia
Private mUltimaReproducidaTXT As Long 'ultima publicidad reproducida (GRANDES) para servir de guia
Private mUltimaReproducidaVMute As Long 'ultima publicidad reproducida (GRANDES) para servir de guia

Private mPubsEnLista As Long 'SOLO PARA PUB MP3/VIDEO.
'Publicidades en la lista de espera. Para saber cuanto
'restar del los temas pendientes reales. Como quedan en la lsita de reproduccion el indicador
' TEMAS PENDIENTES (X) debe saber cuanto restar al total que tiene la lista

Public Property Get HabilitarPublicidadesVMute() As Boolean
    HabilitarPublicidadesVMute = mHabilitarPublicidadesVMute
End Property

Public Property Let HabilitarPublicidadesVMute(Publicidades As Boolean)
    mHabilitarPublicidadesVMute = Publicidades
End Property

Public Property Get HabilitarPublicidadesTXT() As Boolean
    HabilitarPublicidadesTXT = mHabilitarPublicidadesTXT
End Property

Public Property Let HabilitarPublicidadesTXT(Publicidades As Boolean)
    mHabilitarPublicidadesTXT = Publicidades
End Property

Public Property Get HabilitarPublicidadesMp3Vid() As Boolean
    HabilitarPublicidadesMp3Vid = mHabilitarPublicidadesMp3Vid
End Property

Public Property Let HabilitarPublicidadesMp3Vid(Publicidades As Boolean)
    mHabilitarPublicidadesMp3Vid = Publicidades
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
    've cuantos temas sin publicidad van. Al pasar una publicidad _
        se pone en cero
    ContadorTemas = mContadorTemas
End Property

Public Property Let ContadorTemas(cTemas As Long)
    've cuantos temas sin publicidad van. Al pasar una publicidad se _
        pone en cero
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

Public Function ArchsVMute(uIndex As Long) As String
    'matriz de las publicidades que hay
    'todas se reproducen alfabeticamente para que el usuario pueda tener algun control como
    'ponerle números para saber que va a salir o duplicar alguna si le pinta
    If uIndex > UBound(mPubsMute) Then
        ArchsVMute = ""
        'MsgBox "Boludo, me estas pidiendo una publicidad que no existe!!!"
    Else
        ArchsVMute = mPubsSortMute(uIndex)
    End If
End Function

Public Function ArchsPubsTXT(uIndex As Long) As String
    'matriz de las publicidades que hay
    'todas se reproducen alfabeticamente para que el usuario pueda tener algun control como
    'ponerle números para saber que va a salir o duplicar alguna si le pinta
    If uIndex > UBound(mPubsTXT) Then
        ArchsPubsTXT = ""
        'MsgBox "Boludo, me estas pidiendo una publicidad que no existe!!!"
    Else
        ArchsPubsTXT = mPubsSortTXT(uIndex)
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

Public Function ArchsPubsBigIMG(uIndex As Long) As String
    'matriz de las publicidades que hay
    'todas se reproducen alfabeticamente para que el usuario pueda tener algun control como
    'ponerle números para saber que va a salir o duplicar alguna si le pinta
    If uIndex > UBound(mPubsBigIMG) Then
        ArchsPubsBigIMG = ""
        'MsgBox "Boludo, me estas pidiendo una publicidad que no existe!!!"
    Else
        ArchsPubsBigIMG = mPubsSORTBigIMG(uIndex)
    End If
End Function

Public Sub CargarPUBs()
    'cargar la lista de publicidades disponibles
    'debe estar en AP+"pub"
    ReDim mPubs(0)
    ReDim mPubsTXT(0)
    ReDim mPubsIMG(0)
    ReDim mPubsBigIMG(0)
    ReDim mPubsMute(0)
    Dim CONT As Long: CONT = 0
    Dim ContImg As Long: ContImg = 0
    Dim ContBigImg As Long: ContBigImg = 0
    Dim ContTXT As Long: ContTXT = 0
    
    Dim Arch As String
    Arch = Dir(AP + "pub\*.*", vbArchive)
    Dim IsMM As Boolean
    Dim IsIMG As Boolean
    Dim IsTXT As Boolean
    Dim IsBigIMG As Boolean 'imagen de mas de 50 KB para la salida de video
    Do While Arch <> ""
        IsIMG = False: IsMM = False: IsBigIMG = False: IsTXT = False
        'ver que tenga los formatos que busco
        If LCase(Right(Arch, 3)) = "txt" Then IsTXT = True
        If LCase(Right(Arch, 3)) = "mp3" Then IsMM = True
        '''If LCase(Right(Arch, 3)) = "mp4" Then IsMM = True
        If LCase(Right(Arch, 3)) = "wma" Then IsMM = True
        If LCase(Right(Arch, 3)) = "avi" Then IsMM = True
        If LCase(Right(Arch, 3)) = "vob" Then IsMM = True
        If LCase(Right(Arch, 3)) = "mpg" Then IsMM = True
        If LCase(Right(Arch, 3)) = "dat" Then IsMM = True
        If LCase(Right(Arch, 4)) = "mpeg" Then IsMM = True
        'ver si es IMG
        If LCase(Right(Arch, 3)) = "jpg" Then IsIMG = True
        If LCase(Right(Arch, 3)) = "gif" Then IsIMG = True
        If LCase(Right(Arch, 3)) = "bmp" Then IsIMG = True
        If LCase(Right(Arch, 4)) = "jpeg" Then IsIMG = True
        'si es de mas de 50 entonces es BigIMG
        If IsIMG Then
            Dim TamPub As Long
            TamPub = FileLen(AP + "pub\" + Arch)
            If TamPub > 50000 Then
                'es grande no poner en la lista de publicidades para la index
                IsIMG = False
                IsBigIMG = True
            End If
        End If
        If IsMM Then
            CONT = CONT + 1
            ReDim Preserve mPubs(CONT)
            mPubs(CONT) = UCase(AP + "pub\" + Arch)
        End If
        If IsTXT Then
            ContTXT = ContTXT + 1
            ReDim Preserve mPubsTXT(ContTXT)
            mPubsTXT(ContTXT) = UCase(AP + "pub\" + Arch)
        End If
        If IsIMG Then
            ContImg = ContImg + 1
            ReDim Preserve mPubsIMG(ContImg)
            mPubsIMG(ContImg) = UCase(AP + "pub\" + Arch)
        End If
        
        If IsBigIMG Then
            ContBigImg = ContBigImg + 1
            ReDim Preserve mPubsBigIMG(ContBigImg)
            mPubsBigIMG(ContBigImg) = UCase(AP + "pub\" + Arch)
        End If
        
        Arch = Dir
    Loop
    
    'ver los videos mudos
    Arch = Dir(AP + "pubmute\*.*", vbArchive)
    Dim ContMute As Long: ContMute = 0
    Do While Arch <> ""
        IsMM = False
        'ver que tenga los formatos que busco
        If LCase(Right(Arch, 3)) = "avi" Then IsMM = True
        If LCase(Right(Arch, 3)) = "vob" Then IsMM = True
        If LCase(Right(Arch, 3)) = "mpg" Then IsMM = True
        If LCase(Right(Arch, 3)) = "dat" Then IsMM = True
        If LCase(Right(Arch, 4)) = "mpeg" Then IsMM = True
        
        If IsMM Then
            ContMute = ContMute + 1
            ReDim Preserve mPubsMute(ContMute)
            mPubsMute(ContMute) = UCase(AP + "pubmute\" + Arch)
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
    '--------------------------------
    'ahora ordeno las imágenes
    'Dim MinSTR As String 'comparacion de cadenas. Empiezo con el máximo
    'Dim ubicMIN As Long 'indice en la matriz del menor encontrado cada vuelta
    MinSTR = "zzzzzzzzzzzzzzzz"
    'Dim c As Long, mtx As Long, ValComp As String
    c = 0 'cantidad de minimos encontrados
    ReDim Preserve Ordenados(0) 'Dim Ordenados() As Long 'matriz con los indices ordenados
    Do
        For mtx = 0 To UBound(mPubsMute)
            ValComp = mPubsMute(mtx)
            If ValComp < MinSTR Then
                MinSTR = ValComp
                ubicMIN = mtx
            End If
        Next
        'lo mando al ultimo para que no vuelva a salir como primero
        mPubsMute(ubicMIN) = "ZZZ" + mPubsMute(ubicMIN)
        ReDim Preserve Ordenados(c)
        Ordenados(c) = ubicMIN
        c = c + 1
        If c > UBound(mPubsMute) Then Exit Do
        MinSTR = "zzzzzzzzzz"
    Loop
    'cargar todos y sacar las primera 3 zetas
    'limpiar la matriz ordenada
    ReDim mPubsSortMute(0)
    For mtx = 0 To UBound(mPubsMute)
        ReDim Preserve mPubsSortMute(mtx)
        mPubsSortMute(mtx) = Right(mPubsMute(Ordenados(mtx)), _
            Len(mPubsMute(Ordenados(mtx))) - 3)
    Next
    '--------------------------------
    'ahora ordeno los txt
    'Dim MinSTR As String 'comparacion de cadenas. Empiezo con el máximo
    'Dim ubicMIN As Long 'indice en la matriz del menor encontrado cada vuelta
    MinSTR = "zzzzzzzzzzzzzzzz"
    'Dim c As Long, mtx As Long, ValComp As String
    c = 0 'cantidad de minimos encontrados
    ReDim Preserve Ordenados(0) 'Dim Ordenados() As Long 'matriz con los indices ordenados
    Do
        For mtx = 0 To UBound(mPubsTXT)
            ValComp = mPubsTXT(mtx)
            If ValComp < MinSTR Then
                MinSTR = ValComp
                ubicMIN = mtx
            End If
        Next
        'lo mando al ultimo para que no vuelva a salir como primero
        mPubsTXT(ubicMIN) = "ZZZ" + mPubsTXT(ubicMIN)
        ReDim Preserve Ordenados(c)
        Ordenados(c) = ubicMIN
        c = c + 1
        If c > UBound(mPubsTXT) Then Exit Do
        MinSTR = "zzzzzzzzzz"
    Loop
    'cargar todos y sacar las primera 3 zetas
    'limpiar la matriz ordenada
    ReDim mPubsSortTXT(0)
    For mtx = 0 To UBound(mPubsTXT)
        ReDim Preserve mPubsSortTXT(mtx)
        mPubsSortTXT(mtx) = Right(mPubsTXT(Ordenados(mtx)), Len(mPubsTXT(Ordenados(mtx))) - 3)
    Next
    '--------------------------------
    'ahora ordeno las imágenes GRANDES de la salida de video
    'Dim MinSTR As String 'comparacion de cadenas. Empiezo con el máximo
    'Dim ubicMIN As Long 'indice en la matriz del menor encontrado cada vuelta
    MinSTR = "zzzzzzzzzzzzzzzz"
    'Dim c As Long, mtx As Long, ValComp As String
    c = 0 'cantidad de minimos encontrados
    ReDim Preserve Ordenados(0) 'Dim Ordenados() As Long 'matriz con los indices ordenados
    Do
        For mtx = 0 To UBound(mPubsBigIMG)
            ValComp = mPubsBigIMG(mtx)
            If ValComp < MinSTR Then
                MinSTR = ValComp
                ubicMIN = mtx
            End If
        Next
        'lo mando al ultimo para que no vuelva a salir como primero
        mPubsBigIMG(ubicMIN) = "ZZZ" + mPubsBigIMG(ubicMIN)
        ReDim Preserve Ordenados(c)
        Ordenados(c) = ubicMIN
        c = c + 1
        If c > UBound(mPubsBigIMG) Then Exit Do
        MinSTR = "zzzzzzzzzz"
    Loop
    'cargar todos y sacar las primera 3 zetas
    'limpiar la matriz ordenada
    ReDim mPubsSORTBigIMG(0)
    For mtx = 0 To UBound(mPubsBigIMG)
        ReDim Preserve mPubsSORTBigIMG(mtx)
        mPubsSORTBigIMG(mtx) = Right(mPubsBigIMG(Ordenados(mtx)), Len(mPubsBigIMG(Ordenados(mtx))) - 3)
    Next
    
End Sub

Public Property Get TotalPUBs() As Long
    'saber el total para no pasarme de largo
    'empieza en cero
    TotalPUBs = UBound(mPubs)
End Property

Public Property Get TotalPUBsMUTE() As Long
    'saber el total para no pasarme de largo
    'empieza en cero
    TotalPUBsMUTE = UBound(mPubsMute)
End Property

Public Property Get TotalPUBsTXT() As Long
    'saber el total para no pasarme de largo
    'empieza en cero
    TotalPUBsTXT = UBound(mPubsTXT)
End Property

Public Property Get TotalPUBsIMG() As Long
    'saber el total para no pasarme de largo
    'empieza en cero
    TotalPUBsIMG = UBound(mPubsIMG)
End Property

Public Property Get TotalPUBsBigIMG() As Long
    'saber el total para no pasarme de largo
    'empieza en cero
    TotalPUBsBigIMG = UBound(mPubsBigIMG)
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

Public Property Get UltimaReproducidaTXT() As Long
    'indica cual se uso para ayudar al sistema aguiarse
    'en realidad es la ultima que se mando a la lista. Nunca sabe cual es la ultima que se reprodujo
    UltimaReproducidaTXT = mUltimaReproducidaTXT
End Property

Public Property Get UltimaReproducidaBigIMG() As Long
    'indica cual se uso para ayudar al sistema a guiarse
    'en realidad es la ultima que se mando a la lista. Nunca sabe cual es la ultima que se reprodujo
    UltimaReproducidaBigIMG = mUltimaReproducidaBigIMG
End Property

Public Property Get UltimaReproducidaVMute() As Long
    'indica cual se uso para ayudar al sistema a guiarse
    'en realidad es la ultima que se mando a la lista. Nunca sabe cual es la ultima que se reprodujo
    UltimaReproducidaVMute = mUltimaReproducidaVMute
End Property

Public Property Let UltimaReproducidaVMute(Ultima As Long)
    mUltimaReproducidaVMute = Ultima
End Property

Public Property Let UltimaReproducida(Ultima As Long)
    mUltimaReproducida = Ultima
End Property

Public Property Let UltimaReproducidaIMG(Ultima As Long)
    mUltimaReproducidaIMG = Ultima
End Property

Public Property Let UltimaReproducidaTXT(Ultima As Long)
    mUltimaReproducidaTXT = Ultima
End Property

Public Property Let UltimaReproducidaBigIMG(Ultima As Long)
    mUltimaReproducidaBigIMG = Ultima
End Property

Private Sub Class_Initialize()
    mUltimaReproducida = 0
    mUltimaReproducidaIMG = 0
    mUltimaReproducidaBigIMG = 0
    mUltimaReproducidaTXT = 0
    mUltimaReproducidaVMute = 0
End Sub
