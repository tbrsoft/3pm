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
End
Attribute VB_Name = "MP3Play"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Dim vol, vol2
'Public Property Get parar() As Boolean
'parar = UserControl.parar
'
'End Property
'Public Property Let parar(ByVal new_parar As Boolean)
'UserControl.parar() = new_parar
'PropertyChanged "parar"
'End Property

Private Sub UserControl_Resize()
    UserControl.Height = 480
    UserControl.Width = 480
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
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
End Sub

Private Sub UserControl_Terminate()
    mmStop
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
End Sub

Public Function IsPlaying() As Boolean
Static s As String * 30
    mciSendString "status MP3Play mode", s, Len(s), 0
    IsPlaying = (Mid$(s, 1, 7) = "playing")
End Function

Public Function mmPlay()
Dim cmdToDo As String * 255
Dim dwReturn As Long
Dim ret As String * 128

Dim tmp As String * 255
Dim lenShort As Long
Dim ShortPathAndFie As String
    
    If Dir(FileName) = "" Then
       ' mmOpen = "Error with input file"
        mmOpen = "Error al ingresar el archivo"
        Exit Function
    End If
    lenShort = GetShortPathName(FileName, tmp, 255)
    'la funcion transforma todo a 8.3 por que con espacioes
    'el reproductor no anda. JOYA JOYA JOYA
    
'   volu = mciGetDeviceID(lenShort)
    ShortPathAndFie = Left$(tmp, lenShort)
    glo_hWnd = hWnd
    cmdToDo = "open " & ShortPathAndFie & " type MPEGVideo Alias MP3Play"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)


'    If dwReturn < 1 Then  'not success
'        mciGetErrorString dwReturn, ret, 128
'        mmOpen = ret
'       MsgBox "es"
'        MsgBox ret, vbCritical
'        Exit Function
'    End If
    
    mmOpen = "Success"
    mciSendString "play MP3Play", 0, 0, 0
'    vol2 = midiOutSetVolume(volu, frmMP3Play.voltata)
'frmMP3Play.Caption = vol2 & " " & frmMP3Play.voltata
End Function

Public Function mmPause()
    'Enabled = False
    mciSendString "pause MP3Play", 0, 0, 0
'mciSendString ,,,
End Function

Public Function mmStop() As String
    Enabled = False
    mciSendString "stop MP3Play", 0, 0, 0
    mciSendString "close MP3Play", 0, 0, 0
End Function

Public Function PositionInSec()
Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play position", s, Len(s), 0
    PositionInSec = Int(Mid$(s, 1, Len(s)) / 1000)

End Function

Public Function Position()
Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play position", s, Len(s), 0
    sec = Int(Mid$(s, 1, Len(s)) / 1000)
    If sec < 60 Then Position = "0:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        Position = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End Function

Public Function LengthInSec()
Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play length", s, Len(s), 0
    LengthInSec = Int(Val(Mid$(s, 1, Len(s))) / 1000)
               'Round(CInt(Mid$(s, 1, Len(s))) / 1000) (solo en VB6
End Function

Public Function Length()
Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play length", s, Len(s), 0
    sec = Int(Val(Mid$(s, 1, Len(s))) / 1000)
        'Round(CInt(Mid$(s, 1, Len(s))) / 1000) (solo en VB6)
    If sec < 60 Then Length = "0:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        Length = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End Function

Public Function About()
    frmCtlAbout.Show vbModal, Me
End Function

Public Function SeekTo(Second)
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    If IsPlaying = True Then mciSendString "play MP3Play from " & Second, 0, 0, 0
    If IsPlaying = False Then mciSendString "seek MP3Play to " & Second, 0, 0, 0
End Function
