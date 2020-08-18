VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim Ret As String * 128 'lo uso en todos lados

Private WithEvents Reloj0 As tbrTimer.clsTimer
Attribute Reloj0.VB_VarHelpID = -1
Private WithEvents Reloj1 As tbrTimer.clsTimer
Attribute Reloj1.VB_VarHelpID = -1
Private WithEvents Reloj2 As tbrTimer.clsTimer
Attribute Reloj2.VB_VarHelpID = -1
Private WithEvents Reloj3 As tbrTimer.clsTimer
Attribute Reloj3.VB_VarHelpID = -1
Private WithEvents RelojKar As tbrTimer.clsTimer
Attribute RelojKar.VB_VarHelpID = -1

Private Terr As New tbrErrores.clsTbrERR

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
    (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetDeviceID Lib "winmm.dll" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long
Private Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long

Private Type RECT
        left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const WS_CHILD = &H40000000

'Property Variables:
Dim m_FileName(4) As String 'ahora son 4 (para kundera 6.8.200)
Dim m_Volumen(4) As Long 'ahora son 4 (para kundera 6.8.200)
Dim dwReturn As Long

Private Alias(4) As String
'lista de alias para usar en enganches o similares

Public Event Played(SecondsPlayed As Long, iAlias As Long)
Public Event BeginPlay(iAlias As Long)
Public Event EndPlay(iAlias As Long)
Public Event mmError(txtMasHist As String)
Public Event FaltaNextEvKAR(dMiliSec As Double)

Private TMPs As String
Private mTotalSec(3) As Single
Private mTotalFrames(3) As Single
Private mFramePerSecond(3) As Single
Private mSecActual(3) As Single
Private mHastaTema(3) As Long 'al abrir a DLL el total tema(3) que decia _
    cuanto duraba (y lo cortaba si se cancelaba una cancion u otro) o sea _
    en general es la duracion real en segundos pero puede ser menos para _
    cortarlo antes

Private mEsGratis(3) As Boolean

Private FSO As New Scripting.FileSystemObject

Private KAR As New tbrPlayMN0
Private TEM As New tbrSoftEncrMan.clstbrENC


Private MilliSecKarIni As Double
Private MilliSecKar As Double

Public Sub doTem(i As Boolean, cl As String, ai As String, ao As String) 'encripta o desencripta un archivo
    TEM.Encriptar i, cl, ai, ao
End Sub

'no se puede usar matriz en WithEvents !!!
Private Sub Reloj0_Timer(): Reloj_Timer 0: End Sub
Private Sub Reloj1_Timer(): Reloj_Timer 1: End Sub
Private Sub Reloj2_Timer(): Reloj_Timer 2: End Sub
Private Sub Reloj3_Timer(): Reloj_Timer 3: End Sub

Public Function GetCurrentMultimediaPos(iAlias As Long) As Long
    'se refiere a frame actual
    
    Dim dwReturn As Long
    Dim pos As String * 128
    Terr.Anotar "002-0001t", iAlias, Alias(iAlias)
    dwReturn = mciSendString("status " & Alias(iAlias) & " position", pos, 128, 0&)
    Terr.Anotar "002-0001s", dwReturn
    If Not dwReturn = 0 Then  'not success
        GetCurrentMultimediaPos = -1
        Exit Function
    End If
    
    'Success
    GetCurrentMultimediaPos = CLng(pos)
End Function

Public Function GetTotalframes(iAlias As Long) As Long
    
    Dim dwReturn As Long
    Dim Total As String * 128
    
    dwReturn = mciSendString("status " & Alias(iAlias) & " length", Total, 128, 0&)
    
    If Not dwReturn = 0 Then  'not success
        mTotalFrames(iAlias) = -1
        GetTotalframes = -1
        Exit Function
    End If
    
    'Success
    mTotalFrames(iAlias) = CLng(Total)
    GetTotalframes = mTotalFrames(iAlias)
End Function

Public Property Get Volumen(iAlias As Long) As Long
    Terr.Anotar "002-0011", iAlias, m_Volumen(iAlias)
    Volumen = m_Volumen(iAlias) / 10
End Property

Public Property Let Volumen(iAlias As Long, ByVal New_Volumen As Long)
    On Error GoTo ERmp3
    'en mi máquina anda del 0 al 1000 (en todas)
    Terr.Anotar "002-0013"
    m_Volumen(iAlias) = New_Volumen * 10
    TMPs = "SetAudio " + Alias(iAlias) + " Volume To " + CStr(m_Volumen(iAlias))
    Terr.Anotar "002-0014", iAlias, TMPs
    Ret = mciSendString(TMPs, 0&, 0&, 0&)
    Terr.Anotar "002-0015", Ret
    If Ret <> 0 Then
        LogErrorMCI Ret
        'no se pudo modificar el volumen
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoVolumen" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    Terr.Anotar "002-0018"
    Exit Property
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Property

Public Property Get FileName(iAlias As Long) As String
    FileName = m_FileName(iAlias)
End Property

Public Property Let FileName(iAlias As Long, ByVal New_FileName As String)
    m_FileName(iAlias) = New_FileName
End Property

Public Function isPlayingAny() As Boolean
    Dim TmpB As Boolean
    TmpB = False
    Dim A44 As Long
    For A44 = 0 To 3
        If IsPlaying(A44) Then
            TmpB = True
            Exit For
        End If
    Next A44
    isPlayingAny = TmpB
End Function

Public Function isPlayingClock(iAlias As Long) As Boolean
    isPlayingClock = (RELOJ(iAlias).Interval > 0)
End Function

Public Function IsPlaying(iAlias As Long) As Boolean
    Dim S As String * 128
    On Error GoTo ERmp3
    Terr.Anotar "002-0027"
    If m_FileName(iAlias) = "" Then
        Terr.Anotar "002-0028", iAlias
        IsPlaying = False
    Else
        Terr.Anotar "002-0030"
        Ret = mciSendString("status " + Alias(iAlias) + " mode", S, 128, 0&)
        Terr.Anotar "002-0031", Ret, iAlias
        If Ret = 263 Then '263 es cuando no ha abierto nada
            IsPlaying = False
            Exit Function
        End If
        If Ret <> 0 Then
            LogErrorMCI Ret
            'no se pudo modificar el volumen
            RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "IsplayStat" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
            'WriteLog "No se pudo definir el estado de ejecucion." + ". Tema: " + m_FileName + " Function IsPlaying", False
        End If
        'EN ESTE CASO ES NULO O ALGO ASI
        'YA QUE MCI NO TIENE LA CAPACIDAD DE STATUS!!!
        If Ret = 274 Then
            Terr.Anotar "002-0034b", iAlias
            IsPlaying = True
        Else
            Terr.Anotar "002-0034", S
            IsPlaying = (Mid(S, 1, 7) = "playing")
        End If
        
        IsPlaying = (Mid(S, 1, 7) = "playing")
    End If
    
    Exit Function
ERmp3:
    Terr.Anotar "002-0034b"
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function DoOpen(iAlias As Long) As Long
    
    'devuelve
        '0: ok
        '1: no existe el archivo
        '2: WINDOWS NO REPRODUCE MP3 !
        '3 al mandar el Mci Open fallo !
    
    On Error GoTo ERmp3
    MiniCerrar iAlias
    'ver si esta el archivo
    Terr.Anotar "002-0041"
    If FSO.FileExists(m_FileName(iAlias)) = False Then
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoFile" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        DoOpen = 1
        Exit Function
    End If
    
    Terr.Anotar "002-0040"
    Dim lenShort As Long, TMP As String * 255
    lenShort = GetShortPathName(m_FileName(iAlias), TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacioes el reproductor no anda. JOYA JOYA JOYA
    Terr.Anotar "002-0045", m_FileName(iAlias)
    
    Dim FileNameSHORT As String
    FileNameSHORT = left$(TMP, lenShort)
    
    Terr.Anotar "002-0046", FileNameSHORT
    
    Dim cmdToDo As String * 255
    
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias " + _
        Alias(iAlias) + " style " & WS_CHILD
    
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    Terr.Anotar "002-0049", dwReturn
    
    If dwReturn <> 0 Then
        If dwReturn = 263 Then
            'si da el error 263 es probable que la máquina no tenga MCI, lo que le paso a Mauro con W98 PE y a efren con ME
            RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "WINDOWS NO REPRODUCE MP3!!!" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
            DoOpen = 2
            Exit Function
        End If
        
        LogErrorMCI dwReturn
        'no se puedo abrir!!!
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoAbre" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        DoOpen = 3
        Exit Function
    End If
    
    Terr.Anotar "002-0054"
    DoOpen = 0
    MiniInicVarMci iAlias
    
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function DoOpenVideo(Style As String, HWind As Long, _
    X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, iAlias As Long)
    
    On Error GoTo ERmp3
    
    Terr.Anotar "002-0058"
    
    MiniCerrar iAlias
    
    Dim cmdToDo As String * 255
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    If Dir(m_FileName(iAlias)) = "" Then
        Terr.Anotar "002-0065b"
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoFile" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        Exit Function
    End If
    Terr.Anotar "002-0068"
    lenShort = GetShortPathName(m_FileName(iAlias), TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacioes
    'el reproductor no anda. JOYA JOYA JOYA
    Terr.Anotar "002-0069", m_FileName(iAlias)
    'volu = mciGetDeviceID(lenShort)
    FileNameSHORT = left$(TMP, lenShort)
    Terr.Anotar "002-0070", FileNameSHORT
    
    Terr.Anotar "002-0071", HWind
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias " + _
        Alias(iAlias) + " parent " + CStr(HWind) + " style " & WS_CHILD 'Style
        
    Terr.Anotar "002-0072", cmdToDo '                         xxxxxx
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    Terr.Anotar "002-0073", dwReturn, Style, HWind
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoAbreVid" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    Terr.Anotar "002-0076", X1, X2, Y1, Y2
    
    'por si mando ancho o alto en cero!!
    If X2 = 0 Or Y2 = 0 Then
        'Get Window Size
        Dim rec As RECT
        Call GetWindowRect(HWind, rec)
        X2 = rec.Right - rec.left
        Y2 = rec.Bottom - rec.top
    End If

    cmdToDo = "put " + Alias(iAlias) + " window at " + CStr(X1) + " " + CStr(Y1) + _
        " " + CStr(X2) + " " + CStr(Y2)
    
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    Terr.Anotar "002-0077", Style, HWind, dwReturn
    
    If dwReturn <> 0 Then
        '********************
        'si es 346 es que no tiene ventana de presentacion (base+90=MCIERR_NO_WINDOW)
        '¿¿¿¿¿????????
        'probe con style popup y overlapped y no sirven ni solucionan
        If dwReturn = 346 Then
            RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "MCIERR_NO_WINDOW=346" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
            'pasa con videos con codecs nuevos que al cargarse si funciona en
            'WMP pero no en 3PM!!!!!!!!!!!!!!!!!!!
        '********************
        Else
            LogErrorMCI dwReturn
            
            RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoAbreVid2" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        End If
    End If
    
    MiniInicVarMci iAlias
    
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Private Function GetFramesPerSecond(iAlias As Long) As Long
    mFramePerSecond(iAlias) = mTotalFrames(iAlias) / (mTotalSec(iAlias))
    GetFramesPerSecond = mFramePerSecond(iAlias)
End Function

Public Function DoPlay(iAlias As Long, Optional FullScreen As Boolean = False) As Long
    '0 ok
    '1 error mci
    
    On Error GoTo ERmp3
    Terr.Anotar "002-0082", CStr(FullScreen), iAlias
    If FullScreen Then
        'dwReturn = mciSendString("play " + Alias(iAlias) + " fullscreen from 0 to " + CStr(mTotalFrames(iAlias)), 0&, 0&, 0&)
        dwReturn = mciSendString("play " + Alias(iAlias) + " fullscreen from 0", 0&, 0&, 0&)
    Else
        'dwReturn = mciSendString("play " + Alias(iAlias) + " from 0 to " + CStr(mTotalFrames(iAlias)), 0&, 0&, 0&)
        dwReturn = mciSendString("play " + Alias(iAlias) + " from 0", 0&, 0&, 0&)
    End If
    If dwReturn <> 0 Then
        DoPlay = 1
        LogErrorMCI dwReturn
        
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoPlay" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        Exit Function
    End If
    Terr.Anotar "002-0086", dwReturn
    RELOJ(iAlias).Interval = 400 'puede variar
    Terr.Anotar "002-0087"
    DoPlay = 0
    RaiseEvent BeginPlay(iAlias)
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function DoPause(iAlias As Long)
    On Error GoTo ERmp3
    Terr.Anotar "002-0088"
    dwReturn = mciSendString("pause " + Alias(iAlias), 0&, 0&, 0&)
    Terr.Anotar "002-0089"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoPause" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    Terr.Anotar "002-0092"
    RELOJ(iAlias).Interval = 0
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function DoStop(iAlias As Long) As String
    On Error GoTo ERmp3
    Terr.Anotar "002-0093", iAlias
    dwReturn = mciSendString("stop " + Alias(iAlias), 0&, 0&, 0&)
    If dwReturn <> 0 And dwReturn <> 263 Then
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoStop" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    Terr.Anotar "002-0097"
    RELOJ(iAlias).Interval = 0
    Terr.Anotar "002-0098"
    RaiseEvent EndPlay(iAlias)
    
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function DoClose(iAlias As Long) As String
    'iAlias es el que hay que cerrar o:
        '99 cierra todos
        
    On Error GoTo ERmp3
    Terr.Anotar "002-0099", iAlias
    If iAlias = 99 Then 'cierra todos
        Dim F11 As Long
        For F11 = 0 To 3
            dwReturn = mciSendString("stop " + Alias(F11), 0&, 0&, 0&)
            RELOJ(F11).Interval = 0
        Next F11
        
        'limpiar el kar !
        RelojKar.Interval = 0: RelojKar.Enabled = False: KAR.BorraTemp
        dwReturn = mciSendString("close all", 0&, 0&, 0&)
        
        If dwReturn <> 0 And dwReturn <> 263 Then '263 ES CUANDO NO HAY NADA ABIERTO
            LogErrorMCI dwReturn
            RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoClose" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        End If
        
    Else 'o solo el elegido
        dwReturn = mciSendString("close " + Alias(iAlias), 0&, 0&, 0&)
        If dwReturn <> 0 And dwReturn <> 263 Then '263 ES CUANDO NO HAY NADA ABIERTO
            LogErrorMCI dwReturn
            RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoClose2" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        End If
        Terr.Anotar "002-0103"
        RELOJ(iAlias).Interval = 0
        
        'reloj especial del karaoke!
        If iAlias = 2 Then
            RelojKar.Interval = 0: RelojKar.Enabled = False
            KAR.BorraTemp
        End If
            
    End If
    'SI SIGUE EL RELOJ SE MARCAN 1000 errores!!!!!!!!!!
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function PercentPlay(iAlias As Long)
    On Error GoTo ERmp3
    Terr.Anotar "002-0104"
    PercentPlay = PositionInSec(iAlias) / mTotalSec(iAlias) * 100
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function PositionInSec(iAlias As Long) As Long
    PositionInSec = mSecActual(iAlias)
End Function

Public Function FaltaInSec(iAlias As Long)
    On Error GoTo ERmp3
    Terr.Anotar "002-0130"
    FaltaInSec = mTotalSec(iAlias) - mSecActual(iAlias)
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function Position(iAlias As Long) As String
    On Error GoTo ERmp3
    Dim Mins As Long, Sec As Long
    Terr.Anotar "002-0131"
    Sec = mSecActual(iAlias)
    Terr.Anotar "002-0132r", Sec, iAlias
    If Sec < 60 Then Position = "00:" & Format(Sec, "00")
    If Sec > 59 Then
        Terr.Anotar "002-0134r"
        Mins = Int(Sec / 60)
        Terr.Anotar "002-0135r", Mins
        Sec = Sec - (Mins * 60)
        Terr.Anotar "002-0136r", Sec
        Position = Format(Mins, "00") & ":" & Format(Sec, "00")
    End If
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function Falta(iAlias As Long) As String
    On Error GoTo ERmp3
    Dim Mins As Long, Sec As Long
    Terr.Anotar "002-0131"
    Sec = FaltaInSec(iAlias)
    Terr.Anotar "002-0132", Sec, iAlias
    If Sec < 60 Then Falta = "00:" & Format(Sec, "00")
    If Sec > 59 Then
        Terr.Anotar "002-0134"
        Mins = Int(Sec / 60)
        Terr.Anotar "002-0135", Mins
        Sec = Sec - (Mins * 60)
        Terr.Anotar "002-0136", Sec
        Falta = Format(Mins, "00") & ":" & Format(Sec, "00")
    End If
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function LengthInSec(iAlias As Long) As Long
    Terr.Anotar "002-0137"
    LengthInSec = CLng(mTotalSec(iAlias))
End Function

Public Function Length(iAlias As Long) As String
    Dim Sec As Long, Mins As Long
    On Error GoTo ERmp3
    Terr.Anotar "002-0148"
    Sec = mTotalSec(iAlias)
    Terr.Anotar "002-0149"
    If Sec < 60 Then Length = "0:" & Format(Sec, "00")
    Terr.Anotar "002-0150"
    If Sec > 59 Then
        Terr.Anotar "002-0151"
        Mins = Int(Sec / 60)
        Terr.Anotar "002-0152"
        Sec = Sec - (Mins * 60)
        Terr.Anotar "002-0153"
        Length = Format(Mins, "00") & ":" & Format(Sec, "00")
    End If
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function SeekTo(Second, iAlias As Long)
    On Error GoTo ERmp3
    Terr.Anotar "002-0154", iAlias
    
    'transformar segundos a frames
    Dim STF As Long
    STF = (Second / 1000) * mFramePerSecond(iAlias)
    If IsPlaying(iAlias) Then
        Terr.Anotar "002-0155"
        dwReturn = mciSendString("play " + Alias(iAlias) + " from " & STF, 0&, 0&, 0&)
        Terr.Anotar "002-0156"
        If dwReturn <> 0 And dwReturn <> 282 Then '282 es que pide un lugar de tiempo que no existe!
            LogErrorMCI dwReturn
            Terr.Anotar "002-0158"
            RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoSeek" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        End If
    Else
        Terr.Anotar "002-0159"
        dwReturn = mciSendString("seek " + Alias(iAlias) + " to " & STF, 0, 0, 0)
        If dwReturn <> 0 And dwReturn <> 263 And dwReturn <> 282 Then
            LogErrorMCI dwReturn
            RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoSeek2" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        End If
    End If
    
    If iAlias = 2 Then ResincroKar
    
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Function Record() 'no se como hara con los alias, estimo que graba todo ¿¿??
    On Error GoTo ERmp3
    Terr.Anotar "002-0162"
    dwReturn = mciSendString("Close MP3rec", 0&, 0&, 0&)
    Terr.Anotar "002-0163"
    If dwReturn <> 0 And dwReturn <> 263 Then '263 es cuando no hay nada abierto
        Terr.Anotar "002-0164"
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        Terr.Anotar "002-0165"
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoRecord" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    Terr.Anotar "002-0166"
    Dim cmdToDo As String * 255
    Terr.Anotar "002-0167"
    'abrir nuevo
    Terr.Anotar "002-0168"
    cmdToDo = "open new type WaveAudio Alias MP3rec"
    Terr.Anotar "002-0169"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoRecord2" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        Exit Function
    End If
    'iniciar grabacion
    Terr.Anotar "002-0174"
    cmdToDo = "record MP3rec"
    Terr.Anotar "002-0175"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoRec3" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Function StopRecord()
    On Error GoTo ERmp3
    Dim cmdToDo As String * 255
    Terr.Anotar "002-0178"
    'parar nuevo
    cmdToDo = "stop MP3rec"
    Terr.Anotar "002-0179"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    Terr.Anotar "002-0180"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoRec4" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    'grabar grabacion
    Terr.Anotar "002-0182"
    cmdToDo = "save MP3rec c:\3pm.wav"
    Terr.Anotar "002-0183"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    Terr.Anotar "002-0184"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NosREC" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    
    'cerrra grabacion
    Terr.Anotar "002-0185"
    cmdToDo = "Close MP3rec "
    Terr.Anotar "002-0186"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    Terr.Anotar "002-0187"
    If dwReturn <> 0 And dwReturn <> 263 Then '263 es cuando no hay nada abierto
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NosRec2" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Sub LogErrorMCI(CodeErrMCI)
    On Error GoTo ERmp3
    
    Dim Buffer As String, Largo As Integer
    Buffer = Space$(512)
    
    Largo = mciGetErrorString(CodeErrMCI, Buffer, Len(Buffer))
    
    Dim ErrTEXT As String
    
    ErrTEXT = Trim(left(Buffer, Len(Buffer)))
    'en este writelog pongo la fecha y hora
    
    RaiseEvent mmError("****ERRmmPlayer -M-C-I- ****" + vbCrLf + _
                        "N:" + Trim(CStr(CodeErrMCI)) + vbCrLf + _
                        ErrTEXT + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)

    Exit Sub
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Sub

Public Function QuickLargoDeTema(TemaQuick As String) As String
    On Local Error GoTo ErrFunc
    Terr.Anotar "002-0192"
    
    QuickLargoDeTema = "N/S"
    
    FileName(2) = TemaQuick
    DoOpen 2
    MiniInicVarMci 2
    
    If dwReturn = 264 Then 'no hay memoria sufuciente!!!
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoMEM" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        'En este caso queda tildado y no puede volver a mostrar hasta que se
        'cierre el MCI original (el que reproduce)
        Exit Function
    End If
    
    '------------ver el largo--------------
    Dim Sec As Long, Mins As Long
    Sec = mTotalFrames(2) / mFramePerSecond(2)
    Terr.Anotar "002-0210", Sec
    If Sec < 60 Then
        QuickLargoDeTema = "00:" & Format(Sec, "00")
    Else
        Terr.Anotar "002-0212"
        Mins = Int(Sec / 60)
        Terr.Anotar "002-0213"
        Sec = Sec - (Mins * 60)
        Terr.Anotar "002-0214"
        QuickLargoDeTema = Format(Mins, "00") & ":" & Format(Sec, "00")
        Terr.Anotar "002-0215"
    End If
    
    MiniCerrar 2
    
    Exit Function
    
ErrFunc:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
 
End Function

Private Function SoloNumeros(TXT As String) As String
    Dim Largo As Long
    Largo = Len(TXT)
    Dim TmpNumber As String
    TmpNumber = ""
    Dim Letra As String, A As Long
    For A = 1 To Largo
        Letra = Mid(TXT, A, 1)
        If IsNumeric(Letra) Then
            TmpNumber = TmpNumber + Letra
        End If
    Next
    If TmpNumber = "" Then TmpNumber = "0"
    SoloNumeros = TmpNumber
End Function


Public Sub SetDefaultDevice(typeDevice As String, drvDefaultDevice As String)
    'this sub is very important to set the default MCI device
    'maybe xing mpeg installed in your computer and it not support
    'all multimedia files
    'because of this you can rest the default device of MCI to
    'drivers microsft
    'which came with windows or you when install Microsft media player
    'ok any way the default device Following:
    'Device Type        Driver
    'MPEGVideo          mciqtz.drv          this is the most important
    'sequencer          mciseq.drv
    'avivideo           mciavi.drv
    'waveaudio          mciwave.drv
    'videodisc          mcipionr.drv
    'cdaudio            mcicda.drv
    
    'the following for ATI all in Wonder 128 VGA card
    'DvdVideo           MciCinem.drv DVD
    'ATIMPEGVIDEO       mciatim1.drv
    
    'e.g. :
    'SetDefaultDevice "MPEGVideo", "mciqtz.drv" ' this the most
    'improtant device and it will receives calls mci
    'Some programs change this device like xing mpeg
    'and if this occur you can not play all mutimedia files
    'and will occur unexpected errors
    'because of this write this line when your program loaded
    'SetDefaultDevice "MPEGVideo", "mciqtz.drv"
    'to set the strongest default device
    
    'Note: Windows 2000 not use system.ini to set drivers.it use registry.
    
    Dim Res As String
    Dim TMP As String * 255
    Dim Windir As String
    Res = GetWindowsDirectory(TMP, 255)
    Windir = left$(TMP, Res)
    Res = WritePrivateProfileString("MCI", typeDevice, drvDefaultDevice, Windir & "\" & "system.ini")
End Sub

Public Function GetDefaultDevice(typeDevice As String) As String
    'this Function help you if you want to know the default device
    'the parameter must be the device type like:
    'MPEGVideo
    'sequencer
    'avivideo
    'waveaudio
    'videodisc
    'cdaudio
    'and the returned value is a string for the default device
    'Please read the description of sub SetDefaultDevice
    
    Dim TMP As String * 255
    Dim Res As String
    Dim Windir As String
    Res = GetWindowsDirectory(TMP, 255)
    Windir = left$(TMP, Res)
    Res = GetPrivateProfileString("MCI", typeDevice, "None", TMP, 255, Windir & "\" & "system.ini")
    GetDefaultDevice = left$(TMP, Res)
End Function

Private Function MiniCerrar(iAlias As Long) As Long
    
    Terr.Anotar "002-0036", iAlias, Alias(iAlias)
    Ret = mciSendString("Close " + Alias(iAlias), 0&, 0&, 0&)
    Terr.Anotar "002-0037"
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoCierra" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
End Function

Public Sub MiniInicVarMci(iAlias As Long)
    'usar formato de frames que es mmucho mas rapido
    Terr.Anotar "iniMCI001", iAlias
    
    Dim Total As String * 128
    dwReturn = mciSendString("set " & Alias(iAlias) & " time format frames", Total, 128, 0&)
    
    'VER QUE DA!!!
    'xxxx
    
    'llenar las variables basicas para no usar de nuevo mciSend
    GetTotalframes iAlias 'aqui se carga mTotalFrames
    GetTotalTimeByMS iAlias 'aqui se carga mtotalSec
    GetFramesPerSecond iAlias 'aquis e dividen los dos anteriores y se obtiene mFrames per sec
    
    'simepre asi, puede ser cambiado si quiero que termine antes desde 3PM
    mHastaTema(iAlias) = mTotalSec(iAlias)
End Sub

Public Function GetTotalTimeByMS(iAlias As Long) As Long
    
    Dim dwReturn As Long
    Dim TotalTime As String * 128

    dwReturn = mciSendString("set " & Alias(iAlias) & " time format ms", TotalTime, 128, 0&)
    dwReturn = mciSendString("status " & Alias(iAlias) & " length", TotalTime, 128, 0&)
    
    mciSendString "set " & Alias(iAlias) & " time format frames", 0&, 0&, 0& ' return focus to frames not to time
    
    If Not dwReturn = 0 Then  'not success
        mTotalSec(iAlias) = -1
        GetTotalTimeByMS = -1
        Exit Function
    End If
    
    'Success
    mTotalSec(iAlias) = Val(TotalTime) / 1000
    GetTotalTimeByMS = Val(TotalTime)
End Function


Private Sub Class_Initialize()
    Set Reloj0 = New tbrTimer.clsTimer
    Set Reloj1 = New tbrTimer.clsTimer
    Set Reloj2 = New tbrTimer.clsTimer
    Set Reloj3 = New tbrTimer.clsTimer
    Set RelojKar = New tbrTimer.clsTimer
    RelojKar.Enabled = False
    
    Reloj0.Interval = 0
    Reloj1.Interval = 0
    Reloj2.Interval = 0
    Reloj3.Interval = 0
    Reloj0.Enabled = True
    Reloj1.Enabled = True
    Reloj2.Enabled = True
    Reloj3.Enabled = True
    
    Terr.FileLog = App.Path + "\regMM.log"  'en realidad nunca uso el append ya que le paso el detalle al programa usuario de la DLL para que haga lo que quiera!
    Terr.LargoAcumula = 300
    
    Terr.Anotar "MP3001"
    
    Alias(0) = "MP3Play0"
    Alias(1) = "Mp3Play1"
    Alias(2) = "Mp3Play2"
    Alias(3) = "Mp3Play3"
    'asegurarse que se reproduzca de manera predeterminada con lo que corresponde
    
    'this Function help you if you want to know the default device
    'the parameter must be the device type like:
    'MPEGVideo
    'sequencer
    'avivideo
    'waveaudio
    'videodisc
    If Not GetDefaultDevice("MPEGVideo") = "mciqtz.drv" Then
        'if Driver"mciqtz.drv" not the default device for type
        '"MpegVideo" then set mciqtz.drv as a default device
        
        SetDefaultDevice "MPEGVideo", "mciqtz.drv"
        'this mciqtz.drv most improtant driver and it will receives calls mci for MPEG types
        'Some programs change this device like xing mpeg
        'and if this occur you can not play all mutimedia files
        'and will occur unexpected errors
    End If
    
    If Not GetDefaultDevice("sequencer") = "mciseq.drv" Then
        'if Driver"mciseq.drv" not the default device for type
        '"sequencer" then set mciqtz.drv as a default device
        SetDefaultDevice "sequencer", "mciseq.drv"
    End If
    
    If Not GetDefaultDevice("avivideo") = "mciavi.drv" Then
        'if Driver"mciavi.drv" not the default device for type
        '"avivideo" then set avivideo as a default device
        SetDefaultDevice "avivideo", "mciavi.drv"
    End If

End Sub

Private Function RELOJ(i) As tbrTimer.clsTimer
    If i = 0 Then Set RELOJ = Reloj0
    If i = 1 Then Set RELOJ = Reloj1
    If i = 2 Then Set RELOJ = Reloj2
    If i = 3 Then Set RELOJ = Reloj3
End Function

Public Property Let HastaTema(i As Long, lHasta As Long)
    mHastaTema(i) = lHasta
End Property

Public Property Get HastaTema(i As Long) As Long
    HastaTema = mHastaTema(i)
End Property

Public Property Get EsGratis(i As Long) As Boolean
    EsGratis = mEsGratis(i)
End Property

Public Property Let EsGratis(i As Long, NewEsGratis As Boolean)
    mEsGratis(i) = NewEsGratis
End Property

Public Function DoOpenKar(ArchMN0 As String, ToPicture As Object, ToPelota As Object) As Long
    
    On Error GoTo ERmp3
    DoOpenKar = 0
    If FSO.FileExists(ArchMN0) = False Then
        RaiseEvent mmError("****ERRmmPlayerK****" + vbCrLf + _
                        "NoFile" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        DoOpenKar = 1
        Exit Function
    End If
    
    MiniCerrar 2
    ToPicture.AutoRedraw = True
    Terr.Anotar "kar001"
    'limpia el picture y esconde la pelota
    KAR.CleanPic
    Terr.Anotar "Kar002"
    'definir donde se muestra
    KAR.DefinePictureBox ToPicture
    KAR.DefinePelota ToPelota
    Terr.Anotar "Kar003"
    KAR.DoOpenKar ArchMN0 'lo separa y fabrica el mp3 en getPathMP3
    Terr.Anotar "Kar004"
    m_FileName(2) = KAR.GetPathMP3
    DoOpen 2
    
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function DoPlayKar()
    'arrancar la musica
    DoPlay 2
    
    'descifrar en que milisegundo empezo la reproduccion!
    '              = ahora - tiempo que lleva reproducido
    
    ResincroKar
    
    RelojKar.Interval = 40
    RelojKar.Enabled = True
End Function

Public Sub DoStopKar()
    RelojKar.Interval = 0: RelojKar.Enabled = False
    KAR.CleanPic
    DoStop 2
    KAR.BorraTemp
End Sub

Private Sub RelojKar_Timer()
    MilliSecKar = (CDbl(Timer) * 1000) - MilliSecKarIni
    RaiseEvent FaltaNextEvKAR(KAR.EjecutarEvento(MilliSecKar))
End Sub

Public Sub ResincroKar()
    Dim EmpezoHace As Double
    EmpezoHace = GetCurrentMultimediaPos(2) / mFramePerSecond(2)
    MilliSecKarIni = (CDbl(Timer) - EmpezoHace) * 1000
End Sub

Private Sub Reloj_Timer(Index As Integer)
    On Error GoTo ERmp3
    Terr.Anotar "002-0001", Index, mFramePerSecond(CLng(Index))
    'primero ver si termina el tema
    
    Dim CurrPos As Single, lFPS As Long
    CurrPos = GetCurrentMultimediaPos(CLng(Index))
    
    lFPS = mFramePerSecond(CLng(Index))
    
    'no dividir por cero!
    If lFPS = 0 Then lFPS = 1
    mSecActual(CLng(Index)) = CurrPos / lFPS
    Terr.Anotar "002-0001b", CurrPos, mSecActual(Index)
    
    
    Dim Termino As Boolean: Termino = False
    
    If CurrPos = -1 Or mTotalFrames(CLng(Index)) = -1 Then Termino = True: GoTo EndSong
    If CurrPos >= (mTotalFrames(CLng(Index)) - 1) Then Termino = True: GoTo EndSong
    'If IsPlaying(CLng(Index)) = False Then Termino = True: GoTo EndSong
    
    If mSecActual(CLng(Index)) >= mHastaTema(Index) Then Termino = True: GoTo EndSong
    
    'Terr.AppendSinHist "I:" + CStr(Index) + _
        " CurrPos:" + CStr(CurrPos) + _
        " Sec:" + CStr(mSecActual(CLng(Index))) + _
        " TotalF:" + CStr(mTotalFrames(CLng(Index))) + _
        " Termino=" + CStr(Termino)
    
EndSong:
    If Termino Then
        RELOJ(Index).Interval = 0
        DoStop CLng(Index)
        'si termino un karaoke!!!
        If Index = 2 Then
            RelojKar.Interval = 0: RelojKar.Enabled = False
            KAR.CleanPic
        End If
        'RaiseEvent EndPlay(CLng(Index))
        Exit Sub
    Else
        'saber que segundo es ...
        Dim sActual As Long
        RaiseEvent Played(CLng(mSecActual(CLng(Index))), CLng(Index))
    End If
    
    Exit Sub
    
'    If IsPlaying(CLng(Index)) = False Then
'        Reloj(Index).Interval = 0
'        RaiseEvent EndPlay(CLng(Index))
'        Exit Sub 'ESTO NO ESTABA!!!!!!!, seguia mandando el evento!!!!!!!!!!
'    End If
'    'y SOLO si no termino largar el evento. Antes estaba alreves!!!!!!!!!
'    tERR.Anotar "002-0005"
'    RaiseEvent Played(PositionInSec(CLng(Index)), CLng(Index))
'    Exit Sub
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    
    Resume Next
End Sub
