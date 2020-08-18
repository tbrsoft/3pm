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
    CDDATOS = 3
    ImagenNRG = 4
    ImagenISO = 5
    DVDData = 6
    CDISO = 7
End Enum
Private tipoThis As TipoCD 'tipo de la actual compilacion


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
Private LastMsgNumber As Long  'para olvidarme de los textos le doy un id a cada evento que pasa

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

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

'mm94
Public Function GetLastMsgNumber() As String
    GetLastMsgNumber = LastMsgNumber
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
    LastMsgNumber = -1 'mm94
    'mm91
    tERR.Anotar "bgae41"
    Set mMp3 = New tbrBurner.clsMp3
    tERR.Anotar "bgae42"
    Set mImage = New tbrBurner.clsImage 'imagen NRG o ISO
    tERR.Anotar "bgae43"
    Set mAudio = New tbrBurner.clsTracks  'disco de audio viejo
    
    Exit Sub
    
errCD:
    tERR.AppendLog "bgae44", tERR.ErrToTXT(Err)
    Resume Next
End Sub

Public Function Iniciar() As Long
    On Local Error GoTo NoIniCD
    'ver si hay grabadoras, si no ni abro lo del paliza
    Dim H As Long
    H = DetectarUnidades
    tERR.Anotar "bgaf", H
    
    If H >= 1 Then
            
        tERR.Anotar "bgaf-3", H
        Set devices = New tbrBurner.clsDevices
        'Solo se hace la primera vez despues solo se pide DefaultDevice
        
        'ver que sea un discpositivo que sirva para algo antes de asignarlo
        'tERR.Anotar "bgaf-3b", devices.GetItem(0).Name 'SEGUIRAQUI SE VA TODO A LA MIERDA SI PIDO ESTO
        
        'tERR.Anotar "bgaf-3c", devices.GetItem(0).PathImageHardDrive
        tERR.Anotar "bgaf-3d", devices.GetItem(0).ToString
        
        tERR.Anotar "bgaf-4"
        devices.SetDefaultDevice devices.GetItem(0)  'meto la primera grabadora que haya
        
        tERR.Anotar "bgaf-5", TPr
        '' Se deberia recorrer la coleccion de devices para saber cual elegir
        
        
        'si funciona devices = clsdevices
        'pero falla por error 430 device=clsdevice
        
        Set device = devices.GetDefaultDevice
        
        tERR.Anotar "bgaf-6"
        device.Initialize
        
        tERR.Anotar "bgaf-7"
        Iniciar = 0 'todo ok
        Status = 0
    Else
        tERR.Anotar "bgag"
        If H = 0 Then Iniciar = 1 'significa que no hay grabadoras
        Status = -2 'se intento inicializar pero hay error
    End If
    
    tERR.Anotar "bgaf-8"
    Exit Function

NoIniCD:
    Select Case Err.Number
        Case 429 'ActiveX component can't create object
            tERR.AppendLog "NoIniCD02", tERR.ErrToTXT(Err)
            Iniciar = 3
        Case -2147024894 'err de automatizacion
            Iniciar = Err.Number
        Case Else
            tERR.AppendLog "NoIniCD01", tERR.ErrToTXT(Err)
            Iniciar = Err.Number
    End Select
    
End Function

'cuando termina de grabarse un CD y me doy por enterado debo avisar aqui!!!
Public Sub SetStatus(newStatus As Long)
    Status = newStatus
End Sub

Public Function GetStatus() As Long
    GetStatus = Status
End Function

Public Function SetCdType(tipo As TipoCD)
    If IsNull(device) Then
        tERR.AppendLog "NoIniDevice09112"
    Else
        If tipo = CDAudio Then device.MediaType = NERO_MEDIA_TYPE_NERO_MEDIA_CD
        If tipo = CDDATOS Then device.MediaType = NERO_MEDIA_TYPE_NERO_MEDIA_CD 'no estoy seguro!
        If tipo = CDMP3 Then device.MediaType = NERO_MEDIA_TYPE_NERO_MEDIA_CD 'no estoy seguro!
        If tipo = CDISO Then device.MediaType = NERO_MEDIA_TYPE_NERO_MEDIA_CD 'no estoy seguro!
        If tipo = ImagenNRG Then device.MediaType = NERO_MEDIA_TYPE_NERO_MEDIA_DVD_ANY
        If tipo = ImagenISO Then device.MediaType = NERO_MEDIA_TYPE_NERO_MEDIA_DVD_ANY
        If tipo = DVDData Then device.MediaType = NERO_MEDIA_TYPE_NERO_MEDIA_DVD_ANY
        
        tipoThis = tipo
    End If
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
        'case "NOSE":t.TrackTipe = NERO_AUDIO_TRACK_TYPE_NERO_AUDIO_TRACK_THROUGH_EVENT
        Case Else
            'AVISAR AL IDIOTA POR EJEPLO SI SON MPG!!
            tERR.AppendLog "bagk-094", LCase(fso.GetExtensionName(sPath))
            AddTrackAudio = -2
            Exit Function
    End Select
        
    mLog "Agregando track " + fso.GetBaseName(sPath), 1029
    mAudio.Add T

End Function

'empieza a grabar
Public Sub StartSave()
    
    On Local Error GoTo errSAVE
    
    CDDoor "open" 'hasta que el nero lo pida ...
    
    'ver si hay algun proceso mal terminado!!
    'NO SE COMO !
    
    'se acaba de abrir la puerta del CD, darle tiempo antes de mandar la grabación
    'ya que si la ndo y llega el evento que espero un cd todo se aborta
    Status = 1 'si no no muestra nada !!!
    Dim n As Single
    Dim S As Long, LastS As Long 'es para mi propio segudero
    n = Timer + 10
    Do While Timer < n
        DoEvents
        S = CLng(n - Timer)
        If LastS <> S Then
            mLog "Esperando disco vacio" + vbCrLf + CStr(S), 1030
            LastS = S
        End If
    Loop
    mLog "Aguarde ...", 1031

    
    If tipoThis = CDAudio Then device.BurnAudio mAudio
    
    'mm94
    'mMp3.ISOTrack.Name = "" 'n9ombre de la compilacion (nombre del cd)
    
    If tipoThis = CDDATOS Then device.BurnMp3 mMp3
    If tipoThis = CDMP3 Then device.BurnMp3 mMp3
        
    If tipoThis = CDISO Then device.BurnImage mImage
    If tipoThis = ImagenNRG Then device.BurnImage mImage
    If tipoThis = ImagenISO Then device.BurnImage mImage
    
    If tipoThis = DVDData Then device.BurnMp3 mMp3
    
    Exit Sub
    
errSAVE:
    tERR.AppendLog "gasc", tERR.ErrToTXT(Err)
    Select Case Err.Number
        Case -2147220984 'ya hay un proceso de grabacion del nero!!!
            Abort
    End Select
    
    
End Sub

Public Sub Abort()
    'la compilacion de nero se debe borrar
    Select Case tipoThis
        Case CDAudio:            Set mAudio = New tbrBurner.clsTracks
        Case CDDATOS:            Set mMp3 = New tbrBurner.clsMp3
        Case CDISO:              Set mImage = New tbrBurner.clsImage
        Case CDMP3:              Set mMp3 = New tbrBurner.clsMp3
        Case DVDData:            Set mMp3 = New tbrBurner.clsMp3
        Case ImagenISO:          Set mImage = New tbrBurner.clsImage
        Case ImagenNRG:          Set mImage = New tbrBurner.clsImage
    End Select
    
    device.Abort = True
    mLog "Cancelando grabación ...", 1023
End Sub

'dice si se puede grabar, tiene en cuenta el espacio
Public Function CanSaveDataMode(CantMB As Long, paraISO As Boolean) As Boolean
    If paraISO Then
        If CantMB > 720 Then
            CanSaveDataMode = False
        Else
            CanSaveDataMode = True
        End If
    Else
        If CantMB > 690 Then
            CanSaveDataMode = False
        Else
            CanSaveDataMode = True
        End If
    End If
End Function

'mm90
'dice si se puede grabar, tiene en cuenta el espacio
Public Function CanSaveDVDMode(CantMB As Long, paraISO As Boolean) As Boolean
    If paraISO Then
        If CantMB > 4600 Then 'esto es 4,7 GB mentirosos (en decimal o sea 4.700.000.000 bytes = 4,38 GB reales)
            CanSaveDVDMode = False
        Else
            CanSaveDVDMode = True
        End If
    Else
        If CantMB > 4400 Then 'esto es 4,7 GB mentirosos (en decimal o sea 4.700.000.000 bytes = 4,38 GB reales)
            CanSaveDVDMode = False
        Else
            CanSaveDVDMode = True
        End If
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
    'puedo recibir algun texzto y no tener porcentaje o que andresP lo haya sacada. Busco que todo se vea en el 3PM
    If Status = 0 Then Status = 1
    tERR.Anotar "bgab-2", pMsg
    mLog pMsg, 1027 'mm94
    'tERR.AppendSinHist CStr(time) + ": " + pMsg
    RaiseEvent MensajeLog(pMsg)
End Sub

Private Sub mLog(mMsg As String, msgNumber As Long)
    MsgSinLeer = MsgSinLeer + vbCrLf + CStr(time) + ": " + mMsg + " (" + CStr(Status) + ")"
    MsgFull = MsgFull + vbCrLf + CStr(time) + ": " + mMsg + " (" + CStr(Status) + ")"
    LastMsg = mMsg
    LastMsgNumber = msgNumber 'mm94
    tERR.AppendSinHist "papalo " + CStr(time) + ": " + mMsg + " (" + CStr(Status) + ")"
End Sub

Private Sub device_onDoneBurn(ByVal pError As tbrBurner.NERO_BURN_ERROR)
    Select Case pError
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_OK
            'la grabacion finalizo ok
            'mm94 enb esta y en cada case
            mLog "Grabacion finalizo completamente ok", 1009
            tERR.Anotar "bgac"
            Status = 100
            
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_BAD_MESSAGE_FILE
            mLog "Error: Error del archivo de mensaje", 1010
            tERR.AppendLog "QUEMOCD-001", LastMsg
            Status = 100 + pError
            
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_DRIVE_NOT_ALLOWED
            mLog "Error: Unidad de discos no habilitada", 1011
            tERR.AppendLog "QUEMOCD-002", LastMsg
            Status = 100 + pError
            
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_FAILED
            mLog "Fallo la grabación", 1012
            tERR.AppendLog "QUEMOCD-003", LastMsg
            Status = 100 + pError
            
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_FUNCTION_NOT_ALLOWED
            mLog "Error: función no permitida", 1013
            tERR.AppendLog "QUEMOCD-004", LastMsg
            Status = 100 + pError
            
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_INVALID_DRIVE
            mLog "Error: Unidad de disco inválida", 1014
            tERR.AppendLog "QUEMOCD-005", LastMsg
            Status = 100 + pError
            
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_UNKNOWN_CD_FORMAT
            mLog "Error: Formato de disco erróneo", 1015
            tERR.AppendLog "QUEMOCD-006", LastMsg
            Status = 100 + pError
        
        'mm94
        Case tbrBurner.NERO_BURN_ERROR.NERO_BURN_ERROR_NERO_BURN_USER_ABORT
            mLog "proceso cancelado", 1028
            tERR.AppendLog "QUEMOCD-007", LastMsg
            Status = 100 + pError
            
        Case Else 'xxxx ver manejo mas detallado
            mLog "Error desconocido", 1016
            tERR.Anotar "bgad", CLng(pError)
            tERR.AppendLog "QUEMOCD-000"
            Status = 100 + pError
    End Select

End Sub

Private Sub device_OnMessage(ByVal eType As tbrBurner.eMsgType, ByVal pMsg As String)
    'puedo recibir algun texto y no tener porcentaje o que andresP lo haya sacada. Busco que todo se vea en el 3PM
    If Status = 0 Then Status = 1
    
    Select Case eType
        Case eMsgType_NON_EMPTY_CD
            'detecto que hay un cd pero no esta vacio!!
            mLog "El disco insertado no esta vacío !", 1017
        Case eMsgType_SET_PHASE
            'cambio de fase, cada fase es una parte del proceso global y cada una de ellas tiene un texto
            mLog pMsg, 1018
        Case eMsgType_WAIT_CD
            
            'si esta esperando el cd quiere decir que no lo inserto
            'yo por mi cuenta abro la bandeja de cd y le doy x segundos para colocar uno
            'si cuando pongo startSave llego aqui quiere decir que no inserto!!!!
            
            mLog "Esperando disco vacio ...", 1019
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            'parece que no recibo otro evento hasta que no responda y salga de este evento!!!!!
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            'si se va a cancelar debe ser aqui necesariamente!!!
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            'paradoja loca nunca va a llegar el evento de que entro un cd ok hasta que no salga de aca
            'y no pudo cancelar si salgo de aca
            '
            'por todo esto NO USO ESTE EVNETO MAS que para confirmar que no inserto el CD el tiempo que le di
            'por lo tanto ingresar aqui significa cancelar!!!
            Abort
            
        Case eMsgType_WAIT_CD_DONE
            'la espera de un cd virgen joia salio ok
            mLog "Disco insertado - verificado OK", 1021
        Case eMsgType_ABORTED
            'la espera de un cd virgen joia salio ok
            mLog "Grabación cancelada", 1022
            Status = 199
        Case Else
            MsgBox "!!!"
    End Select
    
    RaiseEvent MensajeLog(pMsg)
End Sub

Private Sub device_OnProgress(ByVal ProgressInPercent As Long)
    Status = ProgressInPercent
    'a veces termina otros procesos secundarios y yo entiendo mal por fuera
    If Status >= 100 Then Status = 99
    tERR.Anotar "papale", CStr(time) + ": " + pMsg + " (" + CStr(Status) + ")"
    RaiseEvent PorcGraba(Status)
End Sub

'mm96
Public Function CDDoor(strAction As String) As Long
    On Local Error GoTo e
    'abre "open" o cierra "closed" la puerta de la unidad de CD
    ' y si hay mas de una !!!!
    Dim sRET As String, RET As Long
    sRET = Space(127)
    RET = mciSendString("set CDAudio door " + strAction, sRET, Len(sRET), 0)
    tERR.Anotar "gasa", RET, sRET
    CDDoor = RET
    
    Exit Function
    
e:
    tERR.AppendLog "gasa2", tERR.ErrToTXT(Err)
End Function

