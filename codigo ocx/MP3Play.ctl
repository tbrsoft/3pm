VERSION 5.00
Begin VB.UserControl MP3Play 
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   Picture         =   "MP3Play.ctx":0000
   PropertyPages   =   "MP3Play.ctx":2AFB
   ScaleHeight     =   1620
   ScaleWidth      =   1500
   ToolboxBitmap   =   "MP3Play.ctx":2B0C
   Begin VB.Timer Reloj 
      Left            =   570
      Top             =   570
   End
End
Attribute VB_Name = "MP3Play"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetDeviceID Lib "winmm.dll" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long
Private Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
'Default Property Values:
Const m_def_FileName = ""
'Property Variables:
Dim m_FileName As String
Dim m_Volumen As Long
Dim dwReturn As Long

Event Played(SecondsPlayed As Long)
Event BeginPlay()
Event EndPlay()

Private Sub Reloj_Timer()
    'primero ver si ermina el tema
    If IsPlaying = False Then
        Reloj.Interval = 0
        RaiseEvent EndPlay
        Exit Sub 'ESTO NO ESTABA!!!!!!!, seguia mandando el evento!!!!!!!!!!
    End If
    'y SOLO si no termino largar el evento. Antes estaba alreves!!!!!!!!!
    RaiseEvent Played(PositionInSec)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 1620
    UserControl.Width = 1500
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Volumen() As Long
    m_Volumen = m_Volumen / 10
    Volumen = m_Volumen
End Property

Public Property Let Volumen(ByVal New_Volumen As Long)
    'en mi máquina anda del 0 al 1000
    m_Volumen = New_Volumen * 10 ' * 30 - 3000
    Ret = mciSendString("SetAudio MP3Play Volume To " + CStr(m_Volumen), 0, 0, 0)
    If Ret <> 0 Then
        LogErrorMCI Ret
        'no se pudo modificar el volumen
        WriteLog "No se pudo poner el volumen en " + CStr(New_Volumen) + ". Tema: " + m_FileName + ". Property Let Volume", False
    End If
    PropertyChanged "Volumen"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    m_FileName = New_FileName
    PropertyChanged "FileName"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_FileName = m_def_FileName
    'Visible = True
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
End Sub

Public Function IsPlaying() As Boolean
    If m_FileName = "" Then
        IsPlaying = False
    Else
        Static s As String * 30
        Ret = mciSendString("status MP3Play mode", s, Len(s), 0)
        If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
            LogErrorMCI Ret
            'no se pudo modificar el volumen
            WriteLog "No se pudo definir el estado de ejecucion." + ". Tema: " + m_FileName + " Function IsPlaying", False
        End If
        IsPlaying = (Mid$(s, 1, 7) = "playing")
    End If
End Function

Public Function DoOpen()
    'DoStop    'DoClose
    'si uso esos dos mando el evento endPlay y se arma un kilombo
    Dim Ret As String * 128
    Ret = mciSendString("Close MP3Play", 0, 0, 0)
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        WriteLog "No se pudo cerrar MCI para reabrir tema." + ". Tema: " + m_FileName + " Function DoOpen", False
    End If
    
    Dim cmdToDo As String * 255
    
    
    
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    
    If FSO.FileExists(m_FileName) = False Then        '
        WriteLog "No existe el archivo mp3 que se intenta abrir." + m_FileName + " Function DoOpen", True
        Exit Function
    End If
    lenShort = GetShortPathName(m_FileName, TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacioes
    'el reproductor no anda. JOYA JOYA JOYA
    
    FileNameSHORT = Left$(TMP, lenShort)
    glo_hWnd = hWnd
    
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias MP3Play style child"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        If dwReturn = 263 Then
            'si da el error 263 es probable que la máquina no tenga MCI, lo que le paso a Mauro con W98 PE y a efren con ME
            WriteLog "WINDOWS NO REPRODUCE MP3!!!. INSTALE EL REPRODUCTOR CORESPONDIENTE " + _
                "A SU VERSION DE WINDOWS", True
            MsgBox "No se ha podido abrir el fichero debido a un problema existente en Windows. " + vbCrLf + _
                "Revise que el reproductor multimedia de Windows este instalado y funcione correctamente." + _
                "Notifique a tbrSoft de esto para más detalles"
                
        End If
        
        LogErrorMCI dwReturn
        'no se puedo abrir!!!
        WriteLog "No se pudo abrir un fichero mp3." + ". Tema: " + m_FileName + " Function DoOpen", False
    End If
    
    'uso todo en milisegundos
    dwReturn = mciSendString("set MP3Play time format milliseconds", 0, 0, 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo establecer el formato a milisegundos." + ". Tema: " + m_FileName + " Function DoOpen", False
    End If
    
End Function

Public Function DoOpenVideo(Style As String, HWind As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    'DoStop
    'DoClose
    'si uso esos dos mando el evento endPlay y se arma un kilombo
    Dim Ret As String * 128
    Ret = mciSendString("Close MP3Play", 0, 0, 0)
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        'no se pudo modificar el volumen
        WriteLog "No se pudo cerrar MCI (video) para reabrir tema." + ". Tema: " + m_FileName + "Function DoOpenVideo", False
    End If
    
    Dim cmdToDo As String * 255
    
    
    
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    
    If Dir(m_FileName) = "" Then
        WriteLog "No existe el archivo de video que se intenta abrir." + m_FileName + " Function DoOpenVideo", True
        Exit Function
    End If
    lenShort = GetShortPathName(m_FileName, TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacioes
    'el reproductor no anda. JOYA JOYA JOYA
    
    'volu = mciGetDeviceID(lenShort)
    FileNameSHORT = Left$(TMP, lenShort)
    glo_hWnd = hWnd
    
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias MP3Play style " + Style + " parent " + CStr(HWind)
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo abrir un fichero de video." + ". Tema: " + m_FileName + " Function DoOpenVideo", False
    End If
    
    cmdToDo = "put MP3Play window at " + CStr(X1) + " " + CStr(Y1) + " " + CStr(X2) + " " + CStr(Y2) + " "
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo poner el video en la posicion y el tamaño especificado." + ". Tema: " + m_FileName + " Function DoOpenVideo (put MP3Play window at)", False
    End If
    
    'Dim s As String
    's = Space(30)
    'cmdToDo = "capability MP3Play can stretch"
    'dwReturn = mciSendString(cmdToDo, s, Len(s), 0)
    'If dwReturn <> 0 Then MsgBox "Error n°: " + Str(dwReturn)
   '
   ' s = Trim(Left(s, 4))
   ' If UCase(s) = "TRUE" Then
   '     cmdToDo = "window MP3Play stretch"
   '     dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
   '     If dwReturn <> 0 Then MsgBox "Error n°: " + Str(dwReturn)
   ' End If
        
    'uso todo en milisegundos
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    
End Function

Public Function DoPlay()
    
    dwReturn = mciSendString("play MP3Play", 0, 0, 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo ejecutar un fichero." + m_FileName + " Function DoPlay", False
    End If
    Reloj.Interval = 1000
    
    RaiseEvent BeginPlay
End Function

Public Function DoPause()
    dwReturn = mciSendString("pause MP3Play", 0, 0, 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo poner en pausa un fichero." + m_FileName + " Function DoPause", False
    End If
    Reloj.Interval = 0
End Function

Public Function DoStop() As String
    dwReturn = mciSendString("stop MP3Play", 0, 0, 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo parar un fichero." + m_FileName + " Function DoStop", False
    End If
    Reloj.Interval = 0
    RaiseEvent EndPlay
End Function

Public Function DoClose() As String
    dwReturn = mciSendString("close MP3Play", 0, 0, 0)
    If dwReturn <> 0 And dwReturn <> 263 Then '263 ES CUANDO NO HAY NADA ABIERTO
        LogErrorMCI dwReturn
        WriteLog "No se pudo cerrar MCI." + ". Tema: " + m_FileName + " Function DoClose", False
    End If
    'SI SIGUE EL RELOJ SE MARCAN 1000 errores!!!!!!!!!!
    Reloj.Interval = 0
End Function

Public Function PercentPlay()
    PercentPlay = PositionInSec / LengthInSec * 100
End Function

Public Function PositionInSec()
    On Local Error GoTo ErrFunc
    Static s As String * 30
    dwReturn = mciSendString("status MP3Play position", s, Len(s), 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo establecer la posicion." + m_FileName + " Function PositionInSec", False
    End If
    PositionInSec = Int(Mid$(s, 1, Len(s)) / 1000)
    Exit Function
ErrFunc:
    On Local Error GoTo ErrFunc2
    WriteLog "El Valor devuelto por MCI para PositionInSec no es válido." + ". Tema: " + m_FileName + " Valor= " + s, True
    'el error puede ser que el primer caracter de S no sea valido
    PositionInSec = Int(Mid$(s, 2, Len(s)) / 1000)
    Exit Function
ErrFunc2:
    WriteLog "El Valor devuelto (preuba desde el segundo caracter) por MCI para PositionInSec no es válido. Valor 2° prueba = " + Mid$(s, 2, Len(s)), False
End Function

Public Function Position()
    On Local Error GoTo ErrFunc
    Static s As String * 30
    dwReturn = mciSendString("status MP3Play position", s, Len(s), 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo establecer la posicion." + m_FileName + " Function Position", False
    End If
    
    sec = Int(Mid$(s, 1, Len(s)) / 1000)
    If sec < 60 Then Position = "0:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        Position = Format(mins, "00") & ":" & Format(sec, "00")
    End If
    
    Exit Function
ErrFunc:
    On Local Error GoTo ErrFunc2
    WriteLog "El Valor devuelto por MCI para PositionInSec no es válido." + ". Tema: " + m_FileName + " Valor= " + s, True
    'el error puede ser que el primer caracter de S no sea valido
    Position = Int(Mid$(s, 2, Len(s)) / 1000)
    Resume 'volver a ver que pasa
    Exit Function
ErrFunc2:
    WriteLog "El Valor devuelto (preuba desde el segundo caracter) por MCI para Position no es válido. Valor 2° prueba = " + Mid$(s, 2, Len(s)), False
    
End Function

Public Function FaltaInSec()
    Static s As String * 30
    FaltaInSec = LengthInSec - PositionInSec 'llamo a la funcion para que se manejen los errores desde ahi
End Function

Public Function Falta()
    sec = FaltaInSec
    If sec < 60 Then Falta = "0:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        Falta = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End Function

Public Function LengthInSec()
    On Local Error GoTo ErrFunc
    Static s As String * 30
    dwReturn = mciSendString("status MP3Play length", s, Len(s), 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo establecer la duracion." + m_FileName + " Function LengthInSec", False
    End If
    LengthInSec = Int(Trim(Mid$(s, 1, Len(s)) / 1000))
    Exit Function
ErrFunc:
    On Local Error GoTo ErrFunc2
    WriteLog "El Valor devuelto por MCI para Length no es válido." + ". Tema: " + m_FileName + " Valor= " + s, True
    'el error puede ser que el primer caracter de S no sea valido
    LengthInSec = Int(Trim(Mid$(s, 2, Len(s)) / 1000))
    Resume 'volver a ver que pasa
    Exit Function
ErrFunc2:
    WriteLog "El Valor devuelto (preuba desde el segundo caracter) por MCI para Length no es válido. Valor 2° prueba = " + Mid$(s, 2, Len(s)), False
End Function

Public Function Length()
    sec = LengthInSec 'pateo posibles errores a LengthInSec
    If sec < 60 Then Length = "0:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        Length = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End Function

Public Function SeekTo(Second)
    If IsPlaying = True Then
        dwReturn = mciSendString("play MP3Play from " & Second, 0, 0, 0)
        If dwReturn <> 0 Then
            LogErrorMCI dwReturn
            'no se pudo modificar el volumen
            WriteLog "No se pudo seek mientras se ejecutaba." + m_FileName + " Function SeekTo", False
        End If
    Else
        dwReturn = mciSendString("seek MP3Play to " & Second, 0, 0, 0)
        If dwReturn <> 0 Then
            LogErrorMCI dwReturn
            'no se pudo modificar el volumen
            WriteLog "No se pudo seek mientras estaba detenida la ejecucion." + m_FileName + " Function SeekTo", False
        End If
    End If
End Function

Function Record()
    
    dwReturn = mciSendString("Close MP3rec", 0, 0, 0)
    If dwReturn <> 0 And dwReturn <> 263 Then '263 es cuando no hay nada abierto
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo cerrar para comenzar a grabar." + m_FileName + " Function Record", False
    End If
    
    Dim cmdToDo As String * 255
    
    'abrir nuevo
    cmdToDo = "open new type WaveAudio Alias MP3rec"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo abrir un nuevo WaveAudio para comenzar a grabar." + m_FileName + " Function Record", False
        Exit Function
    End If
    'iniciar grabacion
    cmdToDo = "record MP3rec"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo comenzar a grabar." + m_FileName + " Function Record", False
    End If
End Function

Function StopRecord()
    Dim cmdToDo As String * 255
    
    'parar nuevo
    cmdToDo = "stop MP3rec"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo detener la grabacion." + m_FileName + " Function StopRecord", False
    End If
    'grabar grabacion
    cmdToDo = "save MP3rec 3pm.wav"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo grabar en archivo al detener la grabacion." + m_FileName + " Function StopRecord", False
    End If
    
    'cerrra grabacion
    cmdToDo = "Close MP3rec "
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 And dwReturn <> 263 Then '263 es cuando no hay nada abierto
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo cerrar MCI de grabacion ." + m_FileName + " Function StopRecord", False
    End If
            
End Function

Public Sub LogErrorMCI(CodeErrMCI)
    Dim Buffer As String, Largo As Integer
    Buffer = Space$(512)
    Largo = mciGetErrorString(CodeErrMCI, Buffer, Len(Buffer))
    Dim ErrTEXT As String
    ErrTEXT = Left(Buffer, Len(Buffer))
    'en este writelog pongo la fecha y hora
    WriteLog "Error MCI nº " + Trim(Str(CodeErrMCI)) + ": " + ErrTEXT, True
End Sub

Public Sub WriteLog(TXT As String, PonerFecha As Boolean)
    
    If FSO.FileExists(AP + "log.txt") = False Then
        Set TE = FSO.CreateTextFile(AP + "log.txt", False)
        TE.Close
    End If
    
    'ver si no es demasiado grande
    If FileLen(AP + "log.txt") > 100000 Then 'hasta 100 KB aguanto
        'pasarlo a otro archivo y volver a vrearlo
        If FSO.FileExists(AP + "OLDlog.txt") Then FSO.DeleteFile AP + "OLDlog.txt", True
        FSO.MoveFile AP + "log.txt", AP + "OLDlog.txt"
        Set TE = FSO.CreateTextFile(AP + "log.txt", False)
        TE.Close
    End If
    'finalmente escribir
    
    
    Set TE = FSO.OpenTextFile(AP + "log.txt", ForAppending, False)
    If PonerFecha Then
        TE.WriteLine Trim(Str(Date)) + " / " + Trim(Str(Time)) + vbCrLf + TXT
    Else
        TE.WriteLine TXT
    End If
    TE.Close
End Sub

Public Function QuickLargoDeTema(TemaQuick As String) As String
    QuickLargoDeTema = "N/S"
    '------------cerrar si estaba abierto--------------
    Ret = mciSendString("Close MP3quick", 0, 0, 0)
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no hay abierto nada
        LogErrorMCI Ret
        WriteLog "No se pudo cerrar MCI para reabrir tema MP3 quick." + ". Tema: " + TemaQuick + " Function QuickLargoDeTema", False
    End If
    '------------abrir--------------
    Dim cmdToDo As String * 255
    
    
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    
    If Dir(TemaQuick) = "" Then       '
        WriteLog "No existe el archivo mp3 que se intenta abrir (QUICK)." + TemaQuick + " Function QuickLargoDeTema", True
        Exit Function
    End If
    lenShort = GetShortPathName(TemaQuick, TMP, 255)
    
    FileNameSHORT = Left$(TMP, lenShort)
    glo_hWnd = hWnd
    
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias MP3quick"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    
    If dwReturn = 264 Then 'no hay memoria sufuciente!!!
        LogErrorMCI dwReturn
        WriteLog "No se pudo abrir MCI (QUICK) ." + ". Tema: " + TemaQuick + " Function QuickLargoDeTema", False
        'En este caso queda tildado y no puede volver a mostrar hasta que se
        'cierre el MCI original (el que reproduce)
        Exit Function
    End If
    
    If dwReturn <> 0 And dwReturn <> 264 Then
        LogErrorMCI dwReturn
        WriteLog "No se pudo abrir un fichero mp3 (QUICK)." + ". Tema: " + TemaQuick + " Function QuickLargoDeTema", False
        Exit Function
    End If
    
    '------------poner en milisegundos--------------
    dwReturn = mciSendString("set MP3quick time format milliseconds", 0, 0, 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        WriteLog "No se pudo establecer el formato a milisegundos." + ". Tema: " + temacuick + " Function QuickLargoDeTema", False
    End If
    '------------ver el largo--------------
    On Local Error GoTo ErrFunc
    Static s As String * 30
    dwReturn = mciSendString("status MP3quick length", s, Len(s), 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo establecer la duracion." + TemaQuick + " Function QuickLargoDeTema", False
    End If
    
    sec = Int(Mid$(s, 1, Len(s)) / 1000)
    
    If sec < 60 Then QuickLargoDeTema = "00:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        QuickLargoDeTema = Format(mins, "00") & ":" & Format(sec, "00")
    End If
    Exit Function
ErrFunc:
    On Local Error GoTo ErrFunc2
    WriteLog "El Valor devuelto por MCI para Length (QUICK) no es válido." + ". Tema: " + TemaQuick + " Valor= " + s, True
    'el error puede ser que el primer caracter de S no sea valido
    LargoQuick = Int(Mid$(s, 2, Len(s)) / 1000)
    Resume 'volver a ver que pasa
    Exit Function
ErrFunc2:
    WriteLog "El Valor devuelto (preuba desde el segundo caracter) por MCI (QUICK) para Length no es válido. Valor 2° prueba = " + Mid$(s, 2, Len(s)), False
    
    'Dim Ret As String * 128
    Ret = mciSendString("Close MP3quick", 0, 0, 0)
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        WriteLog "No se pudo cerrar MCI para reabrir tema." + ". Tema: " + m_FileName + " Function DoOpen", False
    End If
    
End Function
