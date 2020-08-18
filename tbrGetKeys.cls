VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrGetKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal vKey As Long) As Integer


Private ListaTeclas() As Byte 'lista de las teclas que deseo escuchar
Private ListaTeclasState() As Long 'lista del estado las teclas para medir mas de una pulsacion
Private ListaTeclasStateLast() As Long

Private LLegados As String


Dim WithEvents TKey As tbrTimer.clsTimer
Attribute TKey.VB_VarHelpID = -1
Public Event LlegoTecla(nTecla As Byte)

Public Sub Startlisen(ListaValores As String)
    'recibe lista de los valores deseados separados por espacio
    Dim SP() As String
    SP = Split(ListaValores)
    Dim A As Long
    For A = 0 To UBound(SP)
        ReDim Preserve ListaTeclas(A): ListaTeclas(A) = CByte(SP(A))
        ReDim Preserve ListaTeclasState(A): ListaTeclasState(A) = 0 'no pulsada
        ReDim Preserve ListaTeclasStateLast(A): ListaTeclasStateLast(A) = 0 'no pulsada
    Next A
    TKey.Interval = 20
    TKey.Enabled = True
End Sub

Public Sub StopLisen()
    TKey.Enabled = False
End Sub

Private Sub Class_Initialize()
    Set TKey = New tbrTimer.clsTimer
    TKey.Enabled = False
    TKey.Interval = 0
End Sub

Private Sub Class_Terminate()
    Set TKey = Nothing
End Sub

Private Sub TKey_Timer()
    Dim KeyResult As Long
    Dim A As Long
    For A = 0 To UBound(ListaTeclas)
        'el numero es el codigo ascii
        KeyResult = GetAsyncKeyState(ListaTeclas(A))
        KeyResult = Abs(KeyResult)
        'el bit mas significativo es 1 cuando el boton esta apretado
        'el menos significativo es uno si la tecla fue apretada desde que se llamo
        'por ultima vez a esta funcion
        
'o sea que:
'ABS 32768 = el boton esta apretado y NO hubo pulsaciones desde la ultima llamada
'ABS 32767 = el boton esta apretado y SI hubo pulsaciones desde la ultima llamada
'(quizas deba tomas esta ultima como dos pulsaciones?)
'ABS 0     = el boton NO esta apretado y NO hubo pulsaciones desde la ultima llamada
'ABS 1     = el NO boton esta apretado pero SI hubo pulsaciones desde la ultima llamada
        
        'o sea si es igual a 32767 deberia ser que la tecla esta apretada
        'pareciera que los valores negativos vienen cuando no tiene el foco
        'If CSng(KeyResult / 2) = CSng(KeyResult \ 2) Then GKs = True
        
        'los 4 valores me joden, me gustaria saber apretado y no apretado!
        If KeyResult >= 32767 Then 'Apretado con o sin señales anteriores
            ListaTeclasState(A) = 32767
        Else
            ListaTeclasState(A) = 1
        End If
        If ListaTeclasStateLast(A) <> ListaTeclasState(A) Then
            ListaTeclasStateLast(A) = ListaTeclasState(A)
            'LLegados = LLegados + vbCrLf + CStr(Timer) + ":" + CStr(ListaTeclasStateLast(A))
            If Abs(ListaTeclasState(A)) = 32767 Then
                RaiseEvent LlegoTecla(ListaTeclas(A))
            End If
        End If
        
'        If GetKey(ListaTeclas(A)) Then
'            If ListaTeclasState(A) = 0 Then  'no pulsada
'                ListaTeclasState(A) = 1 'marcar como pulsada
'                RaiseEvent LlegoTecla(ListaTeclas(A))
'            End If
'        Else
'            ListaTeclasState(A) = 0 'no pulsada
'        End If
    Next A
End Sub

Public Function IsLisen() As Boolean
    IsLisen = (TKey.Interval > 0) And (TKey.Enabled)
End Function
