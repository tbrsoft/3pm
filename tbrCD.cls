VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'modulo de clase para grabar cds de audio y mp3s

'dll del andres paliza
Dim devices As tbrBurner.clsDevices
Dim WithEvents device As tbrBurner.clsDevice
Attribute device.VB_VarHelpID = -1

'///////////////////////////////////////////////////////////
'objetos que pueden ser grabados
Private mImage As tbrBurner.clsImage 'imagen NRG o ISO
Private mMp3 As tbrBurner.clsMp3  'CD cd datos o MP3s
Private mAudio As tbrBurner.clsTracks 'disco de audio viejo
'///////////////////////////////////////////////////////////

'cada una de las unidades para grabar cds detectadas
Private Type tbrDriveCd
    dNombre As String 'alguna descroipcion que encuentre
    dUnidad As String 'letra de unidad
End Type

Private Unidades() As tbrDriveCd 'paquete de unidades de cd disponibles
'el cero no va, datos validos a partir del 1

Public Enum TipoCD
    CDAudio = 1
    CDMP3 = 2
    CDDatos = 3
    ImagenNGR = 4
    ImagenISO = 5
    DVDData = 6
End Enum
Private tipeThis As TipoCD 'tipo de la actual compilacion


'evento de porcentaje
Public Event PorcGraba(porc As Long)
Public Event MensajeLog(Mensaje As String)

Private Status As Long
'porcentaje de grabacion
'-1 si no esta iniciado
'0-99 si esta grabando (99 es que termino algun proceso interno llamado fase)
'100 si finalizo ok la grabacion
'100+x con el error x

Private MsgSinLeer As String 'acumula mensajes de log sin leer para cuando no use los eventos
Private MsgFull As String 'acumula todos los mensajes
Private LastMsg As String 'ultima de las notificaciones!

'solo una vez al inicio de 3Pm o la primera vez que se inicie el carrito
Public Function DetectarUnidades() As Long
    'devuelve la cantidad encontrada, un cero es una alerta de que no hay
    tERR.Anotar "bgaa"
    ReDim Unidades(0)
    
    Dim DRs As Drives, DS As Drive
    Set DRs = fso.Drives

    For Each DS In DRs
        tERR.Anotar "gbaa", DS.DriveType
        If DS.DriveType = 4 Then '4 es CDROM 'supongo que son grbadoras, no debería ser tan muerto
            If DS.IsReady Then
                AgregarUnidad DS.VolumeName, DS.DriveLetter
            Else
                AgregarUnidad "SIN DISCO", DS.DriveLetter
            End If
        End If
    Next
    
    DetectarUnidades = Cantidad
End Function

Public Function Cantidad() As Long
    Cantidad = UBound(Unidades)
End Function

Private Function AgregarUnidad(pNombre As String, pUnidad As String) As tbrDriveCd
    Dim H As Long
    H = UBound(Unidades) + 1
    ReDim Preserve Unidades(H)
    Unidades(H).dNombre = pNombre
    Unidades(H).dUnidad = pUnidad
    tERR.Anotar "bgab", pNombre + "//" + pUnidad
    AgregarUnidad = Unidades(H)
End Function

Public Function GetMsgSinLeer() As String
    GetMsgSinLeer = MsgSinLeer
    MsgSinLeer = ""
End Function

Public Function GetMsgFull() As String
    GetMsgFull = MsgFull
End Function

Public Function GetLastMsg() As String
    GetLastMsg = LastMsg
End Function

Public Function CleanMsgFull()
    MsgFull = ""
End Function

Private Sub Class_Initialize()
    'si no tiene el nero o tbrBurner instalado y registyrado la linea que genera el clsMP3 dara error
    On Local Error GoTo errCD
    
    ReDim Unidades(0)
    Status = -1 'indica que no esta grabando ni haciendo nada
    tERR.Anotar "bgae"
    MsgSinLeer = ""
    MsgFull = ""
    LastMsg = ""
    
    'mm91
    tERR.Anotar "bgae41"
    Set mMp3 = New tbrBurner.clsMp3
    tERR.Anotar "bgae42"
    Set mImage = New tbrBurner.clsImage 'imagen NRG o ISO
    tERR.Anotar "bgae43"
    Set mAudio = New tbrBurner.clsTracks 'disco de audio viejo
    
    Exit Sub
    
errCD:
    tERR.AppendLog "bgae44", tERR.ErrToTXT(Err)
    Resume Next
End Sub

Public Function Iniciar() As Long
    On Local Error GoTo NoIniCD
    'ver si hay grabadoras
    Dim H As Long
    H = DetectarUnidades
    tERR.Anotar "bgaf", H
    
    If H >= 1 Then
        
        Set devices = New tbrBurner.clsDevices
        'Solo se hace la primera vez despues solo se pide DefaultDevice
        devices.SetDefaultDevice devices.GetItem(0) 'meto la priomera grabadora que haya
        
        ''' Se deberia recorrer la coleccion de devices para saber cual elegir
        Set device = devices.GetDefaultDevice()
        device.Initialize
        
        Iniciar = 0 'todo ok
        Status = 0
    Else
        tERR.Anotar "bgag"
        If H = 0 Then Iniciar = 1 'significa que no hay grabadoras
        Status = -2 'se intento inicializar pero hay error
    End If
    
    Exit Function

NoIniCD:
    Select Case Err.Number
        Case 429 'ActiveX component can't create object
            tERR.AppendLog "NoIniCD02", tERR.ErrToTXT(Err)
            Iniciar = 3
        Case Else
            tERR.AppendLog "NoIniCD01", tERR.ErrToTXT(Err)
    End Select
    
    
End Function

'cuando termina de grabarse un CD y me diçoy por enterado debo avisar aqui!!!
Public Sub SetStatus(newStatus As Long)
    Status = newStatus
End Sub

Public Function GetStatus() As Long
    GetStatus = Status
End Function

Public Function SetCdType(tipo As TipoCD)
    If tipo = CDAudio Then device.MediaType = NERO_MEDIA_TYPE_NERO_MEDIA_CD
    If tipo = CDDatos Then device.MediaType = NERO_MEDIA_TYPE_NERO_MEDIA_CD 'no estoy seguro!
    If tipo = CDMP3 Then device.MediaType = NERO_MEDIA_TYPE_NERO_MEDIA_CD 'no estoy seguro!
    If tipo = ImagenNGR Then device.MediaType = NERO_MEDIA_TYPE_NERO_MEDIA_DVD_ANY
    If tipo = ImagenISO Then device.MediaType = NERO_MEDIA_TYPE_NERO_MEDIA_DVD_ANY
    If tipo = DVDData Then device.MediaType = NERO_MEDIA_TYPE_NERO_MEDIA_DVD_ANY
    
    tipeThis = tipo
End Function

Public Function AddFile(sPath As String) As Long
    If fso.FileExists(sPath) = False Then
        AddFile = -1
        Exit Function
    End If
    
    mMp3.addFileToIsoTrack sPath
End Function

'agregar una carpeta a un cd para grabarse
Public Function AddFolder(sPathFolder As String, Optional RecursiveInside As Boolean = True) As Long
    If fso.FolderExists(sPathFolder) = False Then
        AddFolder = -1
        Exit Function
    End If
    
    If RecursiveInside Then
        mMp3.AddFolderRecursivelyToISOTrack sPathFolder
    Else
        mMp3.AddFolderToISOTrack sPathFolder
    End If
    
    AddFolder = 0
End Function

Public Function SetImageToSave(sPath As String) As Long
    If fso.FileExists(sPath) = False Then
        SetImageToSave = -1
        Exit Function
    End If
    
    'hacer un select case con el formato de la imagen
    'xxxx
 
    mImage.ImagePath = sPath
    SetImageToSave = 0
End Function

Public Function AddTrackAudio(sPath As String) As Long
    
    If fso.FileExists(sPath) = False Then
        AddTrackAudio = -1
        Exit Function
    End If
    
    Dim T As New tbrBurner.clsTrack
    T.FileName = sPath
    
    Select Case LCase(fso.GetExtensionName(sPath))
        Case "mp3": T.TrackTipe = NERO_AUDIO_TRACK_TYPE_NERO_AUDIO_TRACK_FILE_MP3
        Case "wma": T.TrackTipe = NERO_AUDIO_TRACK_TYPE_NERO_AUDIO_TRACK_FILE_WMA
        Case "wav": T.TrackTipe = NERO_AUDIO_TRACK_TYPE_NERO_AUDIO_TRACK_FILE_WAV
        'case"NOSE":t.TrackTipe = NERO_AUDIO_TRACK_TYPE_NERO_AUDIO_TRACK_THROUGH_EVENT
        Case Else
            AddTrackAudio = -2
            Exit Function
    End Select
        
    mAudio.Add T

End Function

'empieza a grabar
Public Sub StartSave()
    
    If tipeThis = CDAudio Then device.BurnAudio mAudio
    
    If tipeThis = CDDatos Then device.BurnMp3 mMp3
    If tipeThis = CDMP3 Then device.BurnMp3 mMp3
        
    If tipeThis = ImagenNGR Then device.BurnImage mImage
    If tipeThis = ImagenISO Then device.BurnImage mImage
    
    '???? sera asi un dvd de datos ????
    If tipeThis = DVDData Then device.BurnMp3 mMp3
    
End Sub

'dice si se puede grabar, tiene en cuenta el espacio
Public Function CanSaveDataMode(CantMB As Long) As Boolean
    If CantMB > 695 Then
        CanSaveDataMode = False
    Else
        CanSaveDataMode = True
    End If
End Function

'mm90
'dice si se puede grabar, tiene en cuenta el espacio
Public Function CanSaveDVDMode(CantMB As Long) As Boolean
    If CantMB > 4482 Then 'esto es 4,7 GB mentirosos (en decimal o sea 4.700.000.000 bytes = 4,38 GB reales)
        CanSaveDVDMode = False
    Else
        CanSaveDVDMode = True
    End If
End Function

'dice si se puede grabar, tiene en cuenta el espacio
Public Function CanSaveAudioMode(CantMinutos As Long) As Boolean
    If CantMinutos > 78 Then
        CanSaveAudioMode = False
    Else
        CanSaveAudioMode = True
    End If
End Function

Private Sub device_onAddLogLine(ByVal pMsg As String)
    tERR.Anotar "bgab-2", pMsg
    mLog pMsg
    'tERR.AppendSinHist CStr(time) + ": " + pMsg
    RaiseEvent MensajeLog(pMsg)
End Sub

Private Sub mLog(mMsg As String)
    MsgSinLeer = MsgSinLeer + vbCrLf + CStr(time) + ": " + mMsg + " (" + CStr(Status) + ")"
    MsgFull = MsgFull + vbCrLf + CStr(time) + ": " + mMsg + " (" + CStr(Status) + ")"
    LastMsg = mMsg
    tERR.AppendSinHist "papalo " + CStr(time) + ": " + mMsg + " (" + CStr(Status) + ")"
End Sub

Private Sub device_onDoneBurn(ByVal pError As tbrBurner.NERO_BURN_ERROR)
    Select Case pError
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_OK
            'la grabacion finalizo ok
            LastMsg = "Grafacion finalizo completamente ok"
            tERR.Anotar "bgac"
            Status = 100
            
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_BAD_MESSAGE_FILE
            LastMsg = "NERO_BURN_ERROR_NERO_BURN_BAD_MESSAGE_FILE"
            tERR.AppendLog "QUEMOCD-001", LastMsg
            Status = 100 + pError
            
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_DRIVE_NOT_ALLOWED
            LastMsg = "NERO_BURN_ERROR_NERO_BURN_DRIVE_NOT_ALLOWED"
            tERR.AppendLog "QUEMOCD-002", LastMsg
            Status = 100 + pError
            
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_FAILED
            LastMsg = "NERO_BURN_ERROR_NERO_BURN_FAILED"
            tERR.AppendLog "QUEMOCD-003", LastMsg
            Status = 100 + pError
            
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_FUNCTION_NOT_ALLOWED
            LastMsg = "NERO_BURN_ERROR_NERO_BURN_FUNCTION_NOT_ALLOWED"
            tERR.AppendLog "QUEMOCD-004", LastMsg
            Status = 100 + pError
            
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_INVALID_DRIVE
            LastMsg = "NERO_BURN_ERROR_NERO_BURN_INVALID_DRIVE"
            tERR.AppendLog "QUEMOCD-005", LastMsg
            Status = 100 + pError
            
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_UNKNOWN_CD_FORMAT
            LastMsg = "NERO_BURN_ERROR_NERO_BURN_UNKNOWN_CD_FORMAT"
            tERR.AppendLog "QUEMOCD-006", LastMsg
            Status = 100 + pError
            
        Case Else 'xxxx ver manejo mas detallado
            LastMsg = "Error desconocido"
            tERR.Anotar "bgad", CLng(pError)
            tERR.AppendLog "QUEMOCD-000"
            Status = 100 + pError
    End Select

End Sub

Private Sub device_OnMessage(ByVal eType As tbrBurner.eMsgType, ByVal pMsg As String)
    Dim MyMsg As String 'traduccion para mi o para mostrarle al cliente
    
    Select Case eType
        Case eMsgType_NON_EMPTY_CD
            'detecto que hay un cd pero no esta vacio!!
            MyMsg = "El disco insertado no esta vacío !"
        Case eMsgType_SET_PHASE
            'cambio de fase, cada fase es una parte del proceso global y cada una de ellas tiene un texto
            MyMsg = pMsg
        Case eMsgType_WAIT_CD
            'esta esperando algun cd ... (puede pasar si requiero un dvd y meten un cd)
            MyMsg = "Esperando disco virgen ..."
        Case eMsgType_WAIT_CD_DONE
            'la espera de un cd virgen joia salio ok
            MyMsg = "Disco insertado - verificado OK"
    End Select
    
    mLog pMsg
    
    RaiseEvent MensajeLog(pMsg)
End Sub

Private Sub device_OnProgress(ByVal ProgressInPercent As Long)
    Status = ProgressInPercent
    'a veces termina otros procesos secundarios y yo entiendo mal por fuera
    If Status >= 100 Then Status = 99
    tERR.Anotar "papale", CStr(time) + ": " + pMsg + " (" + CStr(Status) + ")"
    RaiseEvent PorcGraba(Status)
End Sub

Public Sub CancelarSave()
    mLog "Cacelando grabación ..."
End Sub
