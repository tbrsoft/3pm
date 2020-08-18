VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TbrBtManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function TbrBT_Initialize Lib "tbrBlueC.dll" () As Integer
Private Declare Function TbrBT_Inquiry Lib "tbrBlueC.dll" () As Integer
Private Declare Function TbrBT_Uninitialize Lib "tbrBlueC.dll" () As Integer
Private Declare Function TbrBT_IsStarted Lib "tbrBlueC.dll" () As Integer
Private Declare Sub TbrBT_SetLocalName Lib "tbrBlueC.dll" (ByVal name As String)
Public mcol As New Collection
Public Event inquiryFinish()
Public Event inquiryReport(BtDevice As TbrBtDevice)
Public Event connectionStatusReport(msg As String)

' cuando hago un push object este evento es el que me dice si se pudo concretar
'el envio en msg trae la informacion por la cual no se pudo enviar el objeto
Public Event pushReport(p As Boolean, msg As String)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETTEXT = &HC
Private evHwnd As Long

Private mPushStatus As Long
'estado de un envio
'0 = haciendo nada
'1 = Comenzo envio
'2 = Llegó ok
'3 = Llegó mal
'4 = usuario se canso de esperar o cancelo

Private mInquiereStatus As Long
'0 = haciendo nada
'1 = comezo busqueda
'2 = termino busqueda
Public Property Get PushStatus() As Long
    PushStatus = mPushStatus
End Property

Public Property Let PushStatus(NewPush As Long)
    'lo uso para volver a poner en cero cuando lei el resultado
    mPushStatus = NewPush
End Property

Public Property Get InquiereStatus() As Long
    InquiereStatus = mInquiereStatus
End Property

Public Function UseEventMSG(useHWND As Long)
    evHwnd = useHWND
End Function

'sme es para anuunciar al exterior, writeevent es para acumular privado
Private Function SME(T As String)  'sendMessage Event
    SME = SendMessage(evHwnd, WM_SETTEXT, 0, ByVal T)
    'los mensajes van por tipo separados por |!
    '0|
    '1|inquiereFinish
    '2|ConnectionServiceStatus + texto del status
    '3|se mando el archivo ok
    '4|se mando el archivo MAL
    '5|llego un dispositivo + nombre + address
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    ' para que funcione herramientas atributos del procedimiento
    ' avanzadas -> id del procedimiento -4 , parado sobre este procedimiento
    Set NewEnum = mcol.[_NewEnum]
End Property

Public Sub SetLocalBtName(pName As String)
    TbrBT_SetLocalName pName
End Sub

Public Function IsStarted() As Boolean
     Dim result As Boolean
     
      If TbrBT_IsStarted() = 1 Then
        result = True
      Else
        result = False
      End If
      IsStarted = result
End Function

Public Property Get Count() As Long
    'se usa al obtener el número de elementos de la
    'colección. Sintaxis: Debug.Print x.Count
    Count = mcol.Count
End Property
Public Function getItembyIndex(p As Integer) As TbrBtDevice
    Set getItembyIndex = mcol.Item(p)
End Function
Public Sub Remove(pId As String)
       mcol.Remove pId
End Sub

Public Sub closeBluetooth()
'Dim pM As ProcesosManager
'Dim p As tProceso
'Set pM = New ProcesosManager
'Set p = pM.itemByName("BlueSoleil.exe")
'If Not p Is Nothing Then
'    p.MatarProceso
'End If

End Sub

Public Sub add(p As TbrBtDevice)
    If itemByName(p.name) Is Nothing Then
    mcol.add p, Trim(p.uAddr)
    End If
End Sub

Public Property Get Item(pId As Variant) As TbrBtDevice
   On Error Resume Next
   Set Item = mcol.Item(pId)
End Property

Public Function itemByName(nombre As String) As TbrBtDevice
    Dim p As TbrBtDevice
    For Each p In mcol
        If p.name = nombre Then Set itemByName = mcol.Item(p.uAddr)
    Next
End Function

Public Function itemByAddress(address As String) As TbrBtDevice
    Dim p As TbrBtDevice
    For Each p In mcol
        If p.getAddress = address Then Set itemByAddress = mcol.Item(p.uAddr)
    Next
End Function

Public Sub inquiryDev()
    mInquiereStatus = 1
    ReiniciarColeccion
    TbrBT_Inquiry
End Sub

Public Sub ReiniciarColeccion()
    'cuando sale un usuario y entra otro hay otros dispositivos en el ambiente
    'es otro momento. Si esto este modulo se queda como que siempre esta lo que busco la ultima vez
    'y en realidad es mas dinámico
    Set mcol = Nothing
    Set mcol = New Collection
End Sub

'Esta funcion se utiliza para que pueda mandar mensajes desde C++ appName es el nombre del
'formulario principal
Public Sub RegisterCallBack(FormCaption As String)
   TbrBT_RegisterCallBack FormCaption
End Sub

Public Sub initialize()

    evHwnd = 0
    mPushStatus = 0
    mInquiereStatus = 0
    
    Dim res As Long
    On Error Resume Next

    Kill App.path + "\BtInfo.log"
    
    res = TbrBT_Initialize()
    If res <> 1 Then
        MsgBox "Ocurrio un error al inicializar la Libreria" + Str(res)
    End If
End Sub
Public Sub unInitialize()
    TbrBT_Uninitialize
End Sub
Friend Sub BtM_CBK_inqFinsish()
    mInquiereStatus = 2
    RaiseEvent inquiryFinish
    If evHwnd <> 0 Then SME "1"
End Sub

'Friend Sub BtM_CBK_ConnectionServiceStatus(ByRef p As ConnectionStructure)
'    RaiseEvent connectionStatusReport(CStr(p.vMsg))
'End Sub

Friend Sub BtM_CBK_PushReport(b As Boolean, msg As String)
    
    If b Then
        mPushStatus = 2
        If evHwnd <> 0 Then SME "3|" + msg
    Else
        mPushStatus = 3
        If evHwnd <> 0 Then SME "4|" + msg
    End If
    
    RaiseEvent pushReport(b, msg)
End Sub
Friend Sub BtM_CBK_inquiry(ByRef p As InquiryStructure)
    Dim T As New TbrBtDevice
    T.name = p.vName
    T.uAddr = p.vAddress
    T.PairStatus = p.lPairStatus
    add T
    RaiseEvent inquiryReport(T)
    If evHwnd <> 0 Then SME "5|" + T.name + "|" + T.getAddress
End Sub

