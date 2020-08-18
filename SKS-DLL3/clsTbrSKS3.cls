VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTbrSKS3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'posicion normal modo escucha del puerto desde la PC.
'posicion normal del puerto de status h37.
'
'luis:
'1- Escribe el puerto de datos.
'2- status en h77 significa leer el puerto
'3- yo escibo 50 + el dato recibido
'4- yo mando interrupcion (pasar el bit 0 de 1 a 0)
'5- si el dato lo entendi bien luis pone el status en h37 (posicion normal). Si
'
'luis tiene varios datos en cola pone el status directamente en h77 para empezar
'
'el proceso de nuevo
'6- Si luis no verifica OK la comunicacion vuelve a escribir el puerto de datos y
'el status en h7F. Se vuelve al 3) hasta una maximo de 2 veces.
'
'7- vuelve al modo normal h37 es status
'
'8- no nos entendimos y dejamos ese dato de lado
'valores reservados
'

'124: Habilitar teclado (hasta ese momento luis ignora a los botones)
'126: Le aviso a luis que el programa se cerro (luis ignora los botones)
'no usar el 126 para resetar, el 124 solo basta
'125: pido que se simule el supuesto boton 24
'123: escuche que luis mando el 63 (64) (que dice que no nos entendemos luego de 2 intentos)
'123 tambien significa que recibi ok un dato de licencia
'en realizada cambia el status a cero para seguir recibiendo interrupciones

'122: luis me interrumpio y el dato leido no es valido (1 al 49 son validos)

'NUEVOS VALORES
Dim nPLACA(4) As Long
Dim NumPLACA As Long 'numero final
'121: me da las cifras 9 y 10 del numero de placa                   nPlaca(4)
'120: me da las cifras 7 y 8 del numero de placa                    nPlaca(3)
'119: me da las cifras 5 y 6 del numero de placa                    nPlaca(2)
'118: me da las cifras 3 y 4 del numero de placa                    nPlaca(1)
'117: me da las cifras 1 y 2 del numero de placa (las menos signif) nPlaca(0)

'LICENCIA!!!
'116: aumentar contador 1 (cont1++) (valacum1=valacum1+mv1(cont1)+cont4)
'115: aumentar contador 2 (cont2++) (valacum1=valacum1+mv2(cont2)+cont3)
'114: aumentar contador 3 (cont3++) (valacum1=valacum1+mv3(cont3)+cont2)
'113: cont4++v devuelve un valor al azar entre 40 y 49 (inclusive ambos)

Dim Cont1 As Long
Dim Cont2 As Long
Dim Cont3 As Long
Dim Cont4 As Long

Dim ValAcum1 As Long

'cont1 va de 1 a 64 y al ser 65 se hace 1
'valacum1 puede ser desde cero a 63 no pasarse de alli y hacer un x mod 64 para que nunca se pase
'si valacum recibido es mayor que 63 algo esta mal

'al 116 la respuesta de luis es valacum1, al 115 valacum2, al 114 valacum3

'112 resetear a cero cont 1,2 y 3 y valacum 1,2 y 3
'tambien el luis se resetea con el 126 que le informo que me apago
'luis manda un 111 para certificar que puso en cero sus contadores

Private Declare Function inP Lib "inpout32.dll" Alias "Inp32" (ByVal PortAddress As Integer) As Integer
Private Declare Sub Out Lib "inpout32.dll" Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)

Dim LastDato As Byte 'para pasar lo ultimo leido entre varios procedimientos
Dim LastStatus As Byte 'ultima valor para saber cuando cambia el puerto de status

Dim LastDatoT2 As Byte 'solo para situaciones excepcionales que se prende el T2

Dim ptoConv As Long 'en que punto estoy de la conversacion (son los de arriba)
'ptoConv = 7  recibi boton ok y quedo dispnible
'ptoConv = 2  me dice que lea el puerto por primera vez
'ptoConv = 3  lei el dato de luis, le respondo con lo que escuche y me voy a esperar que me confirme
'ptoConv = 37 lei el dato de luis por segunda vez, le respondo con lo que escuche y me voy a esperar que me confirme

Dim EstoyEnElReloj As Long '0 no estoy, 1 si estoy
Dim Counters(23) As Long 'cantidad de veces de cada señal

'cada vez que reciba una señal se la mando a alguien!
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETTEXT = &HC
'HWND identificador del objeto que recibira los eventos!
Private mHwndMsg As Long

Private mInterval As Long

Dim WithEvents tbrTimer1 As tbrTimer.clsTimer 'para leer el status
Attribute tbrTimer1.VB_VarHelpID = -1
Dim WithEvents T2 As tbrTimer.clsTimer 'para leer los datos. NUNCA USAR. SOLO DEPURAR FALLAS
Attribute T2.VB_VarHelpID = -1

Private mV1, mV2, mV3 'vectores de los textos
Private SignalType As Byte 'para cuando hablemos de licencia no lleguen como botones
'0 es boton

Private totLIC As Long 'total pedidas
Private totLicOK As Long 'total ok

Private tmpLastDato As Long 'lo pongo en -1 cuando quiero y espero un resultado especifico
'en principio para no esperar tiempo de mas en las licencias

Private OutPort As Integer
Private InPort As Integer
Private CtrlPort As Integer

'sme es para anuunciar al exterior, writeevent es para acumular privado
Private Function SME(T As String)  'sendMessage Event
    SME = SendMessage(mHwndMsg, WM_SETTEXT, 0, ByVal T)
    'los mensajes van por tipo separados por |!
    '0|mbitdata     datos que cambian
    '1|mbitStatus   status que cambia
    '2|info         mensajes para acumular para el error
    '3|error        error exacto
End Function

Public Property Get HwndMsg() As Long
    HwndMsg = mHwndMsg
End Property

Public Property Let HwndMsg(NewHW As Long)
    'primero cargar la propiedad hwndisform para que se vea aca el nombre !!!
    'WriteEvent "66|" + CStr(NewHW) + "|" + mFormActual + "||", True, 0
    mHwndMsg = NewHW
End Property

Private Sub T2_Timer()

    'leo todo el tiempo el puerto de datos por la dudas que este escuchando mal el status
    Dim LS2 As Byte
    LS2 = GetDato(False)
    
    If LS2 <> LastDatoT2 Then
        SME "ChgT2: " + CStr(LS2)
    End If
    
    LastDatoT2 = LS2
End Sub

Private Sub tbrTimer1_Timer()
    If EstoyEnElReloj = 1 Then Exit Sub
    EstoyEnElReloj = 1
    'leo todo el tiempo el puerto de status
    Dim LS2 As Byte
    LS2 = inP(InPort)
    '*********************************************************************
    'agregado nuevo, solo me importan los bits 3 y el 6 (contand de 0 a 7)
    'transformo lo que lei solo para interpretar los valores derivados delos
    'bits que me interesan
    LS2 = ChangeBit(LS2, 0, 0)
    LS2 = ChangeBit(LS2, 1, 0)
    LS2 = ChangeBit(LS2, 2, 0)
    LS2 = ChangeBit(LS2, 4, 0)
    LS2 = ChangeBit(LS2, 5, 0)
    LS2 = ChangeBit(LS2, 7, 0)
    'solo quedaron como estaban los bits 3 y 6
    'entonces el:
    '0 *0* 00 *0* 000 =  0 dec es standBy o OK
    '0 *0* 00 *1* 000 =  8 dec comunicacion con fallas
    '0 *1* 00 *0* 000 = 64 dec es interrupcion para leer el puerto (botones comunes)
    '0 *1* 00 *1* 000 = 72 dec es interrupcion para leer el puerto (botones especiales)
    '*********************************************************************
    If LS2 <> LastStatus Then
        SME "ChgSTAT: " + CStr(LS2) + " * " + CStr(ptoConv) + " * " + CStr(SignalType)
    End If
    
    LastStatus = LS2
    
    Select Case LastStatus
        Case 0 'modo standBy en espera ...
            If ptoConv = 3 Then RecibiOK LastDato, CByte(55)  'tambien puede validar lo que le mande
            If ptoConv = 37 Then RecibiOK LastDato, CByte(56)  'tambien puede validar lo que le mande
            If ptoConv = 8 Then RecibiOK LastDato, CByte(54)  'tambien puede validar lo que le mande
            
            'vuelve o queda al valor ok
            ptoConv = 7
        Case 64 'me dice que lea el puerto
            
            If ptoConv = 3 Then GoTo FINrutina 'no entro 2 veces !!! (botones)
            If ptoConv = 8 Then GoTo FINrutina 'no entro 2 veces !!! (licencia)
            'paro un poco ...
            tbrTimer1.Enabled = False
            
            'como el luis no uso el mismo modelo de comunicaciones no hay validacion
            'en las comunicaciones por licencia
            If SignalType = 0 Then ptoConv = 2                'entre por primera vez
            
            'leo lo que dice
            LastDato = GetDato
            
            If SignalType = 0 Then
                'ver si lei al pedo y en ese caso ignorar
                If LastDato < 1 Or LastDato > 49 Then
                    SendDataLuis 122 'luis se va a ir a 55
                    ptoConv = 7
                Else
                    SendDataLuis LastDato + 50
                End If
                ptoConv = 3
            Else
                'que no entre 2 veces
                ptoConv = 8
                SendDataLuis 123
            End If
            'necesito esperar la confirmacion
            tbrTimer1.Enabled = True

        Case 72 'es el 7F
            
            If ptoConv = 37 Then GoTo FINrutina 'no entro 2 veces !!!
            
            'paro un poco ...
            tbrTimer1.Enabled = False
            
            ptoConv = 2 'entre por primera vez
            'leo lo que dice
            LastDato = GetDato
            
            'ver si lei al pedo y en ese caso ignorar
            If LastDato < 1 Or LastDato > 49 Then
                SendDataLuis 122 'luis se va a ir a 55
                ptoConv = 7
            Else
                SendDataLuis LastDato + 50
            End If
            
            ptoConv = 37 'paso 3 del ntr del 127
            
            'necesito esperar la confirmacion
            tbrTimer1.Enabled = True
        Case 8 'luis me dice que hubo dos comunicaciones seguidas con falla
            'solo una vez
            If ptoConv <> 8 Then SendDataLuis 123
            ptoConv = 8
    End Select
 
FINrutina:
    EstoyEnElReloj = 0
End Sub

Private Function BinToDec(sBin As String) As Long
    sBin = Trim(sBin)
    Dim Largo As Long
    Largo = Len(sBin)
    Dim POS As Long
    POS = 1
    Dim ValTotal As Long
    Dim Vals As Long 'temporales
    Dim Letra As String
    ValTotal = 0
    Dim A As Long
    For A = Largo To 1 Step -1
        Letra = Mid(sBin, A, 1)
        Vals = Val(Letra) * 2 ^ (Largo - A)
        ValTotal = ValTotal + Vals
    Next
    BinToDec = ValTotal
End Function

Private Function GetDato(Optional bShowError As Boolean = True) As Byte
    On Local Error GoTo NoD
    
    'me aseguro de escuchar
    Out CtrlPort, ChangeBit(inP(CtrlPort), 5, 1)
    Dim GD As Byte
    GD = inP(OutPort)
    'sacarle el fakin bit mas significativo
    'por que el luis no puede escribirlo
    GD = ChangeBit(GD, 7, 0)
    
    'los valores no validos supuestamente no los escribe luis en el puerto por lo _
        tanto los ignoro
        
'    If GD < 1 Or GD > 49 Then
'        If bShowError Then SME "GD MAL:" + CStr(GD)
'        'GD = 0 'bandera de dato leido al pedo
'    Else
'        If bShowError Then SME "GD  OK:" + CStr(GD)
'    End If
    
    If bShowError Then SME "GD:" + CStr(GD)
    GetDato = GD
    
    Exit Function
    
NoD:
    SME ("ErrGetDato")
    Resume Next
End Function

Private Function ChangeBit(ValToChange As Byte, nBit As Long, ValToSet As Long) As Byte
    On Local Error GoTo NoD
    
    'sme "VP-" + CStr(ValToChange) + "," + CStr(nBit) + "," + CStr(ValToSet)
    'veo el valor al que le voy a hacer un OR
    'si quiero setear un cero es mas facil
    Dim sB As String
    If ValToSet = 0 Then
        sB = String(7 - nBit, "1") 'los ceros mas significativos que el buscado
        sB = sB + "0"
        sB = sB + String(nBit, "1")
    Else
        sB = String(7 - nBit, "0") 'los ceros mas significativos que el buscado
        sB = sB + "1"
        sB = sB + String(nBit, "0")
    End If
    
    'sme "VP-sb=" + sB
    
    'lo transformo segun corresponda
    Dim ValPort As Byte
    If ValToSet = 0 Then
        ValPort = ValToChange And BinToDec(sB)
    Else
        ValPort = ValToChange Or BinToDec(sB)
    End If
    
    'sme "VP:" + CStr(ValPort)
    ChangeBit = ValPort
    
    Exit Function
    
NoD:
    SME "ErrChgBit"
    Resume Next
    
End Function

Public Sub Escribir(txtPort As String, txtValue As String)
    Out CInt(txtPort), CInt(txtValue) 'escribir algo
End Sub

Public Sub Prender()
    'compatibilidad con el ECP !!!!!!!!!
    Out &H77A, BinToDec("00100000")

    'es como un reseteo de lecturas, tomo como que no estoy esperando nada
    ptoConv = 7
    'prender poniendo el dato en 124
    SendDataLuis 124
    'con este boton virtual saco la demora que siempre hay al prender el chip de Luis
    SendDataLuis 125
    'luis manda un 24 confirmando que prendio OK. Con un delay
End Sub

Private Sub EsperarSec(Sec As Single)
    Dim T As Single
    T = Timer
    Do While Timer < T + Sec: DoEvents: Loop
End Sub

Public Function AddCont(nCont_0_to_3 As Long, Optional myTimeOut As Single = 1) As Long
    'lo reviso al hacer el ok
    SendDataLuis nCont_0_to_3 + 113, True
    'esperar la respuesta !!!
    SME "*** AddCont: " + CStr(nCont_0_to_3) + " * " + CStr(myTimeOut)
    tmpLastDato = -1 'apenas reciba un dato esto se cambia
    Dim J As Single
    J = Timer
    Do While J + myTimeOut > Timer
        
        DoEvents
        
        If tmpLastDato = -2 Then
            'llego un numero pero no era el esperado!
            AddCont = 1 'ok, recibi un dato como esperaba
            SME "*** AddCont=1!!"
            Exit Function
        End If
        
        If tmpLastDato >= 0 Then
            'NUMERO OK!
            AddCont = 0
            SME "*** AddCont=0!!"
            Exit Function
        End If
        
    Loop
    SME "*** AddCont=2!!"
    AddCont = 2 'si se queda con este es que llego al time out
    
    'luis me pidio que cuando haya un time out iniciar todo por las dudas
    'ya que generalmente falla a continuación
    
    ReIniContLuis
    
    esperar 1
    
End Function

Private Sub SendDataLuis(sData As Byte, Optional AsignarSType As Boolean = False)

    On Local Error GoTo NoD
    
    'para cuando pido cosas que no tiene que ver con botones
    If AsignarSType Then SignalType = sData
    'se va solo a cero al recibir OK!!!!
'    Else
'        SignalType = 0
'    End If
'
    SME "Send:" + CStr(sData) + ":" + CStr(SignalType)
    
    tbrTimer1.Enabled = False
    
    Dim iTMP As Integer
    iTMP = ChangeBit(inP(CtrlPort), 5, 0)
    'sme "Put 37A " + CStr(iTMP)
    'pasar al modo habla poniendo en cero el bit 5 del puerto de control
    Out CtrlPort, iTMP
    'escribir el valor deseado
    'SME "Put 378 " + CStr(sData)
    Out OutPort, CInt(sData)
    'le aviso a luis que esta el dato interrumpiendolo
    iTMP = ChangeBit(inP(CtrlPort), 0, 1)
    'sme "Put 37A " + CStr(iTMP)
    Out CtrlPort, iTMP
    
    iTMP = ChangeBit(inP(CtrlPort), 0, 0)
    'sme "Put 37A " + CStr(iTMP)
    Out CtrlPort, iTMP
    
    'volver al modo escucha
    iTMP = ChangeBit(inP(CtrlPort), 5, 1)
    'sme "Put 37A " + CStr(iTMP)
    Out CtrlPort, iTMP
    
    tbrTimer1.Enabled = True
    
    Exit Sub
    
NoD:
    SME "ErrSenData"
    Resume Next
End Sub

Private Sub RecibiOK(sVal As Byte, sModo As Byte)
        
    If SignalType = 0 Then
        If sVal <= 23 Then Counters(sVal) = Counters(sVal) + 1
        SME "sD:" + CStr(sVal) + ":" + CStr(sModo)
    Else
        'valores que espero a veces para saber el numero de placa
        If SignalType = 117 Then nPLACA(0) = sVal
        If SignalType = 118 Then nPLACA(1) = sVal
        If SignalType = 119 Then nPLACA(2) = sVal
        If SignalType = 120 Then nPLACA(3) = sVal
        If SignalType = 121 Then nPLACA(4) = sVal
        
        '**************************** LICENCIA!!!
        'pide contador
        '116: aumentar contador 1 (cont1++) (valacum1=valacum1+mv1(cont1)+cont4)
        '115: aumentar contador 2 (cont2++) (valacum1=valacum1+mv2(cont2)+cont3)
        '114: aumentar contador 3 (cont3++) (valacum1=valacum1+mv3(cont3)+cont2)
        '113: cont4++ devuelve un valor al azar entre 40 y 49 (inclusive ambos)
        
        If SignalType > 112 And SignalType < 117 Then
            
            If SignalType = 116 Then
                Cont1 = Cont1 + 1
                Mod64 Cont1
                ValAcum1 = ValAcum1 + CLng(mV1(Cont1)) + Cont4
            End If
            
            If SignalType = 115 Then
                Cont2 = Cont2 + 1
                Mod64 Cont2
                ValAcum1 = ValAcum1 + CLng(mV2(Cont2)) + Cont3
            End If
                
            If SignalType = 114 Then
                Cont3 = Cont3 + 1
                Mod64 Cont3
                ValAcum1 = ValAcum1 + CLng(mV3(Cont3)) + Cont2
            End If
            
            If SignalType = 113 Then
                Cont4 = Cont4 + 1
                Mod64 Cont4
            End If
            
            Mod64 ValAcum1
            
            SME "L" + CStr(SignalType) 'T
            SME "VAA:" + CStr(ValAcum1) 'T
            SME "mV1(c1):" + CStr(mV1(Cont1))
            
            SME "L1:" + CStr(sVal) + " L2:" + CStr(ValAcum1)
            totLIC = totLIC + 1
            
            'en todos los casos espera el numro de ValAcum1 pero en el 113 es un numero al azar entre 40 y 49
            
            If ((sVal = ValAcum1) And (SignalType <> 113)) Or _
                ((sVal > 39) And (sVal < 50) And (SignalType = 113)) Then
                
                totLicOK = totLicOK + 1
                SME "L18"
                tmpLastDato = sVal
            Else
                tmpLastDato = -2 'bandera de licencia mal
                SME "L17"
            End If
            SME "Cs " + CStr(Cont1) + CStr(Cont2) + CStr(Cont3) + CStr(Cont4)
            
        End If
        '**************************** FIN LICENCIA
        
        If SignalType = 112 Then
            If sVal = 111 Then
                SME "ReIniConts OK 111"
                ReIniCont
            Else
                SME "ReIniConts MAL " + CStr(sVal)
            End If
        End If
        
        SME "xD:" + CStr(sVal) + ":" + CStr(sModo) + ":" + CStr(SignalType)
        
        SignalType = 0 'puede seguir un boton
    End If
    
    'If sVal = 24 Then 'se prendio OK !!
    'sme "recibi OK (" + CStr(sModo) + "):" + CStr(sVal)
    
End Sub

Public Function GetResLicSTR() As String
    GetResLicSTR = "LIC:" + CStr(totLIC) + " - OK:" + CStr(totLicOK)
End Function

Public Function GetPorcLic() As Single
    Dim Ts As Single
    Ts = CSng(totLicOK / totLIC) * 100
    GetPorcLic = Ts
End Function

Private Function Mod64(ByRef L As Long)
    If L > 63 Then L = L Mod 64
End Function

Public Sub ReIniCounters()
    Dim I As Long
    For I = 0 To 23
        Counters(I) = 0
    Next I
End Sub

Public Function GetCounter(I As Long) As Long
    If I > 23 Then
        GetCounter = -1
    Else
        GetCounter = Counters(I)
    End If
End Function

Public Sub setPorts(outP As Integer, inP As Integer, ctrlP As Integer)
    OutPort = outP
    InPort = inP
    CtrlPort = ctrlP
End Sub

Private Sub Class_Initialize()

    'una PCI+LPT tenia estas direcciones
    'OutPort = &HB050
    'InPort = &HB051
    'CtrlPort = &HB052
    
    OutPort = &H378
    InPort = &H379
    CtrlPort = &H37A
    
    'esto es lpt2
    'OutPort = &H278
    'InPort = &H279
    'CtrlPort = &H27A


    'indicamos que la placa no esta cargada todavia
    nPLACA(0) = -1
    nPLACA(1) = -1
    nPLACA(2) = -1
    nPLACA(3) = -1
    nPLACA(4) = -1
    
    'INIT antes apenas se creaba la instancia arracaba
    'ahora esta aparte para poder setear los puertos antes de arranacar
End Sub

Public Sub INIT()

    ReIniCont
    
    Set tbrTimer1 = New tbrTimer.clsTimer
    tbrTimer1.Enabled = True
    tbrTimer1.Interval = 50
    
    Set T2 = New tbrTimer.clsTimer
    T2.Enabled = False
    
    SignalType = 0 'predeterminado son botones
    
    totLIC = 0
    totLicOK = 0
End Sub

Private Sub ReIniCont()
    Cont1 = 0
    Cont2 = 0
    Cont3 = 0
    Cont4 = 0
    ValAcum1 = 0
    totLIC = 0
    totLicOK = 0
End Sub

Public Sub ReIniContLuis()
    'este solo lo hace si luis me confirma que los puso el tambien en cero
    SendDataLuis 112, True
End Sub

Public Function GetnPlaca(fLista As String) As String

    '121: me da las cifras 9 y 10 del numero de placa                   nPlaca(4)
    '120: me da las cifras 7 y 8 del numero de placa                    nPlaca(3)
    '119: me da las cifras 5 y 6 del numero de placa                    nPlaca(2)
    '118: me da las cifras 3 y 4 del numero de placa                    nPlaca(1)
    '117: me da las cifras 1 y 2 del numero de placa (las mas signif) nPlaca(0)
    
    Dim H As Long, TT As Single
    For H = 0 To 4
        SME "gp" + CStr(H) + "vv"
        nPLACA(H) = -1 'la dejo vacia
        SendDataLuis 117 + H, True 'la pido
        TT = Timer
        Do While nPLACA(H) < 0
            DoEvents
            'no espero toda la visa
            If TT + 1 < Timer Then Exit Do
        Loop
    Next H
    
    For H = 0 To 4
        SME "gp" + CStr(H) + "vv" + CStr(nPLACA(H))
        'ver si alguno no se leyo y lo consulto de nuevo
        If nPLACA(H) < 0 Then
            SendDataLuis 117 + H, True 'la pido
            TT = Timer
            Do While nPLACA(H) < 0
                DoEvents
                'no espero toda la visa
                If TT + 1 < Timer Then Exit Do
            Loop
        End If
    Next H
    
    'lo que viene son botones
    SignalType = 0
    
    '//////////////////////////////////
    'extraño caso excepcional de rocchio
    'da: 127 - 93 - 54 - 96 - 9 donde deberia dar
    '      0 - 93 - 54 - 96 - 9!!!
    If nPLACA(0) > 99 Then nPLACA(0) = 0
    '//////////////////////////////////
    
    If nPLACA(0) < 0 Or nPLACA(1) < 0 Or nPLACA(2) < 0 Or nPLACA(3) < 0 Or nPLACA(4) < 0 Then
        NumPLACA = -1
        SME "NP:-1"
    Else
        NumPLACA = nPLACA(4) + _
               nPLACA(3) * 100 + _
               nPLACA(2) * 10000 + _
               nPLACA(1) * 1000000 + _
               nPLACA(0) * 100000000
        
        SME "gp231" + CStr(NumPLACA)
        
        GetVal CStr(NumPLACA), fLista, mV1, mV2, mV3
        
        SME "NP:" + CStr(NumPLACA) + " " + CStr(mV1(0)) + CStr(mV2(0)) + CStr(mV3(0))
        
    End If

    GetnPlaca = CStr(NumPLACA)

End Function

Public Function RvP(f As String, P As String) As String
    'me dice si un numero de placa anda ok
    GetVal P, f, mV1, mV2, mV3
    
    SME "NP:" + P + " " + CStr(mV1(0)) + "." + CStr(mV2(0)) + "." + CStr(mV3(0))
    RvP = "NP:" + P + " " + CStr(mV1(0)) + "." + CStr(mV2(0)) + "." + CStr(mV3(0))
    
End Function

Public Sub ToTimer2(Prender As Boolean)
    T2.Interval = 30
    T2.Enabled = Prender
End Sub

Public Sub SetInterval(NewIntervalMS As Long)
    mInterval = NewIntervalMS
    If NewIntervalMS < 5 Then mInterval = 5
    If NewIntervalMS > 500 Then mInterval = 500
    tbrTimer1.Interval = mInterval
End Sub

Public Function GetInterval() As Long
    GetInterval = mInterval
End Function

Private Sub Class_Terminate()
    SendDataLuis 126 'apagar botones
End Sub

Private Sub esperar(N As Single)
    N = Timer + N
    Do While Timer < N
        DoEvents
    Loop
End Sub

