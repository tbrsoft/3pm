VERSION 5.00
Object = "{165405B0-4A0E-4DCF-BE23-FAF75F9F9126}#1.0#0"; "tbrGrap_b.ocx"
Begin VB.Form F1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estadisticas 3PM"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrSoftGRAP_b.tbrGRAP G2 
      Height          =   3675
      Left            =   120
      TabIndex        =   3
      Top             =   4260
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   6482
   End
   Begin tbrSoftGRAP_b.tbrGRAP G1 
      Height          =   4005
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   7064
   End
   Begin VB.PictureBox picPorc 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   60
      ScaleHeight     =   375
      ScaleWidth      =   1245
      TabIndex        =   1
      Top             =   90
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tecla Izquierda SALIR - Tecla Derecha Eliminar todos los datos (reiniciar estadisticas)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   510
      TabIndex        =   0
      Top             =   8130
      Width           =   10365
   End
End
Attribute VB_Name = "F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New Scripting.FileSystemObject
Dim AP As String
Dim SF As String
Dim WF As String
Private Type tElement 'para ritmos, discos y canciones
    N1 As Long 'numeros sin sentido
    N2 As Long 'numeros sin sentido
    Nombre As String
    VentasU As Single 'plata por vender USB
    VentasB As Single 'plata por vender Bluetooth
    VentasC As Single 'venta en cds!
    Escuchados As Single 'plata por escuchar
    Fecha As Long 'para poder filtrar por fecha
End Type

'acumuladores de informacion. Un indice para cada elemnto nuevo
Dim Ritmo() As tElement
Dim DISCO() As tElement
Dim Cancion() As tElement
Dim Equipos() As tElement

Dim Platilla As tElement 'es contador completo de la suma de todo

'existe la posibilidad que vengan registros de equipos que tengan diferentes separadores decimales
'me aseguro aqui de contar como este equipo cuente
Dim SD As String 'separador decimal
Dim SDno As String 'el de los dos que no es separador1

Private Function AgregarTElement(tel() As tElement, Eq As String) As Long
    'se fija si esta el equipo, si mn o crea uno nuevo
    'devuelve el indice en la matriz equipos
    
    Dim J As Long, Esta As Long
    Esta = -1
    For J = 1 To UBound(tel)
        If Eq = tel(J).Nombre Then
            Esta = J
            Exit For
        End If
    Next J
    
    If Esta = -1 Then
        J = UBound(tel) + 1
        ReDim Preserve tel(J)
    Else
        J = Esta
    End If
    
    tel(J).Nombre = Eq
    AgregarTElement = J
    
End Function

Public Function STRceros(n As Variant, Cifras As Integer) As String
    'n es el numero y cifras es la cantidad final de cifras del str terminado
    'devuelve ej : para 232,6 = 000232 para 1902,12 = 000000001902
    'complaeta con ceroas adelante
    ' si n es mas lasgo que cifras devuelve el valor n sin ningun cero adelante
    Dim STRn As String
    STRn = Trim(CStr(n))
    Dim DIF As Integer
    DIF = Cifras - Len(STRn)
    If DIF > 0 Then
        Dim CEROstr As String
        CEROstr = String(DIF, "0")
        STRceros = CEROstr + STRn
    Else
        STRceros = STRn
    End If
    
End Function

Private Function FolderToRes()
    'me da la carpeta donde estan los archivo y deja todo el registro
    'el segundo parametro es el archivo al que se agregaran los resultados
    
    On Local Error GoTo FRE
    
    Dim LEE As String
    
    '****************************************************
    'son 5 archivos en 3pm/sf
    'NOMBRE        SEP                                     MEAN                               QI
    '"jumal.los"   Chr(5) + Chr(7) + Chr(6) + Chr(4)       ID1ttt + t + ID2ttt                "Ingrese su pais de residencia"
    '"guen.w"      Chr(5) + Chr(6) + Chr(6) + Chr(5)       ID2ttt + tt + ID1ttt               "Telefono o fax"
    '"japi.lon"    Chr(7) + Chr(7) + Chr(6) + Chr(5)       ID1ttt + ID2ttt + ttt              "Email tecnico"
    '"buca.rest"   Chr(4) + Chr(7) + Chr(6) + Chr(5)       T + ID1ttt + tt + ID2ttt + ttt     "Email administrativo"
    '"buda.pest"   Chr(4) + Chr(6) + Chr(6) + Chr(4)       ID1ttt + ID2ttt                    "Gracias por confiar en tbrSoft"
    
    'donde ID1ttt e ID2ttt son numeros al azar de 7 digitos SIEMPRE
    't   = path completo (de aqui sale cancion + discos + ritmo)
    'tt  = id del equipo
    'ttt = fecha y hora
    '****************************************************
    Dim ARC(5) As String
    porc 11, "Desencriptando"
    
    '*****************
    'si ya se limpio no entrar
    If fso.FileExists(AP + "sf\jumal.los") = False Then Exit Function
    
    ARC(1) = tLt(AP + "sf\jumal.los", Chr(5) + Chr(7) + Chr(6) + Chr(4), "Ingrese su pais de residencia")
    porc 17, "Desencriptando"
    ARC(2) = tLt(AP + "sf\guen.w", Chr(5) + Chr(6) + Chr(6) + Chr(5), "Telefono o fax")
    porc 29, "Desencriptando"
    ARC(3) = tLt(AP + "sf\japi.lon", Chr(7) + Chr(7) + Chr(6) + Chr(5), "Email tecnico")
    porc 51, "Desencriptando"
    ARC(4) = tLt(AP + "sf\buca.rest", Chr(4) + Chr(7) + Chr(6) + Chr(5), "Email administrativo")
    porc 91, "Desencriptando"
    ARC(5) = tLt(AP + "sf\buda.pest", Chr(4) + Chr(6) + Chr(6) + Chr(4), "Gracias por confiar en tbrSoft")
    porc 100, "Desencriptando"

    'ver que coincidan todos los archivos
    Dim SP1() As String
    Dim SP2() As String
    Dim SP3() As String
    Dim SP4() As String
    Dim SP5() As String

    SP1 = Split(ARC(1), Chr(5))
    SP2 = Split(ARC(2), Chr(5))
    SP3 = Split(ARC(3), Chr(5))
    SP4 = Split(ARC(4), Chr(5))
    SP5 = Split(ARC(5), Chr(5))

    'todos tiene que tener la misma cantidad de elementos
    Dim H As Long, EsOk As Boolean
    Dim MiniRes As String
    Dim Tef As TextStream
    
    Set Tef = fso.OpenTextFile(AP + "sf\stats.dss", ForAppending, True)
        For H = 0 To UBound(SP1)
            porc Round((H / (UBound(SP1) + 1)), 2), "Cargando" 'el +1 es para que no divida por cero
            'comparar devuelve el resumen de cada renglon (cancion vendida)
            MiniRes = Comparar(SP1(H), SP2(H), SP3(H), SP4(H), SP5(H))
            Tef.Write MiniRes + Chr(3) 'separador final (internamente es el 5)
        Next H
    Tef.Close
    
    'hay que borrar los archivos para que no se acumulen mas y se duplique!!
    fso.DeleteFile AP + "sf\jumal.los", True
    fso.DeleteFile AP + "sf\guen.w", True
    fso.DeleteFile AP + "sf\japi.lon", True
    fso.DeleteFile AP + "sf\buca.rest", True
    fso.DeleteFile AP + "sf\buda.pest", True
    
    Exit Function
    
FRE:
    'MsgBox Err.Description
    Resume Next
End Function

Private Sub ucdateF1()
    FolderToRes
    PlayInfo
    porc 0, ""
    'cargar matrices para graficar
    Dim mStr() As String, mVal() As Single
    
    G1.Descargar
    On Local Error Resume Next
    ReDim mStr(3): ReDim mVal(4)
    mStr(0) = "Escuchado"
    mStr(1) = "Venta bluetooth"
    mStr(2) = "Venta usb"
    mStr(3) = "Venta CD"
    mVal(0) = Platilla.Escuchados
    mVal(1) = Platilla.VentasB
    mVal(2) = Platilla.VentasU
    mVal(3) = Platilla.VentasC
    
    G1.Titulo = "TOTAL por modo de uso: " + CStr(Round(Platilla.Escuchados + Platilla.VentasB + Platilla.VentasU + Platilla.VentasC, 2))
    G1.LoadFromMtx mStr, mVal
    G1.Mostrar
    
    
    G2.Descargar
    Dim H As Long
    H = UBound(Ritmo)
    
    Dim mToto As Single 'otro total para control
    For H = 1 To UBound(Ritmo)
        ReDim Preserve mStr(H - 1): ReDim Preserve mVal(H - 1)
        mStr(H - 1) = Ritmo(H).Nombre
        mVal(H - 1) = Ritmo(H).Escuchados + Ritmo(H).VentasB + Ritmo(H).VentasC + Ritmo(H).VentasU
        mToto = mToto + mVal(H - 1)
    Next H
    
    G2.Titulo = "Total Por ritmo elegido: " + CStr(Round(mToto, 2))
    G2.LoadFromMtx mStr, mVal
    G2.Mostrar
    
    G1.Visible = True
    G2.Visible = True
    
End Sub

Private Sub CleanAllStats()
    'todo el acumulado esta en stats.dss
    If fso.FileExists(AP + "sf\stats.dss") Then
        Dim hoy As String 'NO VA A FALTAR el boludo que borre sin querer
        hoy = CStr(Year(Date)) + CStr(Month(Date)) + CStr(Day(Date))
        
        If fso.FileExists(AP + "sf\stats.dss." + hoy) Then fso.DeleteFile AP + "sf\stats.dss." + hoy, True
        fso.MoveFile AP + "sf\stats.dss", AP + "sf\stats.dss." + hoy
    End If
End Sub

Private Sub Form_Activate()
    ucdateF1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case TeclaDER
            CleanAllStats
            ucdateF1
        Case Else
            Unload Me
    End Select
End Sub


Private Sub Form_Load()
    
    'leer todo lo que aiga
    
    AP = App.path
    SF = fso.GetSpecialFolder(SystemFolder)
    WF = fso.GetSpecialFolder(WindowsFolder)
    
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    If Right(SF, 1) <> "\" Then SF = SF + "\"
    If Right(WF, 1) <> "\" Then WF = WF + "\"
    
    If fso.FolderExists(AP + "ACUM") = False Then fso.CreateFolder (AP + "ACUM")
    
    'defino cual es el separador decimal en este equipo
    If CSng("0,1") = 0.1 Then
        SD = ","
        SDno = "."
    End If
    
    If CSng("0.1") = 0.1 Then
        SD = "."
        SDno = ","
    End If
    
    picPorc.Width = 15
    picPorc.Visible = False
    
End Sub

Private Function Comparar(v1 As String, V2 As String, v3 As String, v4 As String, v5 As String) As String

    On Local Error GoTo errCOMP
    
    'compara los valores de los 5 logs y dice si esta ok o no!
    Dim ID1ttt(5) As String
    Dim ID2ttt(5) As String
    Dim T(2) As String, tt(2) As String, ttt(2) As String
    
    'en V1
    ID1ttt(1) = Mid(v1, 1, 7)
    ID2ttt(1) = Mid(v1, Len(v1) - 6, 7)
    T(1) = Mid(v1, 8, Len(v1) - 14)
    
    'en V2
    ID2ttt(2) = Mid(V2, 1, 7)
    ID1ttt(2) = Mid(V2, Len(V2) - 6, 7)
    tt(1) = Mid(V2, 8, Len(V2) - 14)
    
    'en V3
    ID1ttt(3) = Mid(v3, 1, 7)
    ID2ttt(3) = Mid(v3, 8, 7)
    ttt(1) = Mid(v3, 15, Len(v3) - 14)
    
    'en V4
    Dim pos As Long
    T(2) = Mid(v4, 1, Len(T(1))) 'no se que largo es de otra forma que buscando la anterior
    
    pos = Len(T(1)) + 1
    ID1ttt(4) = Mid(v4, pos, 7)
    
    pos = pos + 7
    tt(2) = Mid(v4, pos, Len(tt(1))) 'no se que largo es de otra forma que buscando la anterior
    
    pos = pos + Len(tt(1))
    ID2ttt(4) = Mid(v4, pos, 7)
    
    pos = pos + 7
    ttt(2) = Mid(v4, pos, Len(v4) - pos + 1)
    
    'en V5
    ID1ttt(5) = Mid(v5, 1, 7)
    ID2ttt(5) = Mid(v5, 8, 7)
    
    'revisar todo
    '**************************************
    Dim H As Long, TodoOk As Boolean
    TodoOk = True
    For H = 2 To 5
        If ID1ttt(1) <> ID1ttt(H) Then
            TodoOk = False
            Exit For
        End If
    Next H
    
    For H = 2 To 5
        If ID2ttt(1) <> ID2ttt(H) Then
            TodoOk = False
            Exit For
        End If
    Next H
    
    If T(1) <> T(2) Then
        TodoOk = False
    End If
    
    If tt(1) <> tt(2) Then
        TodoOk = False
    End If
    
    If ttt(1) <> ttt(2) Then
        TodoOk = False
    End If
    
    If TodoOk Then
        'si habla de plata t(x) empieza con "P"
        'si no es:
        't es el path pero la primer letra indica dispositivo
        'de este saco dispositivo + origen + disco + canción
        
        Dim sDev As String
        Dim PTH As String
        Dim PTs() As String 'partes del path
        Dim fNAME As String 'filename
        Dim sDisk As String
        Dim sOrig As String
        
        '"E" + TEMA + "*" + plata, pc, fecha 'escucha
        '"U" + TEMA + "*" + plata, pc, fecha 'compra por usb
        '"B" + TEMA + "*" + plata, pc, fecha 'compra por bluetooth
        
        Dim Platilla As String, MotivoPlata As String
        MotivoPlata = Mid(T(1), 1, 1)
        
        Dim SepAst() As String 'separado por asterisco
        SepAst = Split(T(1), "*")
        
        Platilla = SepAst(1) ' Mid(t(1), 3, Len(t(1)) - 2) de antes cuando era "PV precio" sin tema
        
        PTH = Mid(SepAst(0), 2, Len(SepAst(0)) - 1)
        PTs = Split(PTH, "\")
        fNAME = PTs(UBound(PTs))
        sDisk = PTs(UBound(PTs) - 1)
        sOrig = PTs(UBound(PTs) - 2)
        
        Comparar = ID1ttt(1) + Chr(5) + ID2ttt(1) + Chr(5) + MotivoPlata + Chr(5) + _
            Platilla + Chr(5) + tt(1) + Chr(5) + ttt(1) + _
            Chr(5) + sOrig + Chr(5) + sDisk + Chr(5) + fNAME
            'este ultimo renglon agregado para tener estadísticas de los que se escucha tambien
    Else 'hay errores !!!!!!!!!!!!!!!!!!
        Comparar = "0" + Chr(5) + "0" + Chr(5) + "shit" + Chr(5) + T(1) + Chr(5) + T(2) + Chr(5) + _
                    tt(1) + Chr(5) + tt(2) + Chr(5) + _
                    ttt(1) + Chr(5) + ttt(2) + Chr(5) + _
                    ID1ttt(1) + Chr(5) + ID1ttt(2) + Chr(5) + ID1ttt(3) + Chr(5) + ID1ttt(4) + Chr(5) + ID1ttt(5) + Chr(5) + _
                    ID2ttt(1) + Chr(5) + ID2ttt(2) + Chr(5) + ID2ttt(3) + Chr(5) + ID2ttt(4) + Chr(5) + ID2ttt(5) + Chr(5)
    End If
                
Exit Function

errCOMP:
    
    Resume Next

End Function

Private Function GetText(F As String) As String

    If fso.FileExists(F) = False Then
        GetText = "boloooo"
        Exit Function
    End If
    
    Dim TE As TextStream
    Set TE = fso.OpenTextFile(F)
        If TE.AtEndOfStream = False Then
            GetText = TE.ReadAll
        Else
            GetText = "RE-boloooo"
        End If
    TE.Close
End Function

Private Function pinchilon(Texto As String, Clave As String, Invertido As Boolean) As String
    'encriptar
    'Cargo los datos
    
    If Texto = "" Then
        pinchilon = ""
        Exit Function
    End If
    
    Dim F As Integer
    Dim Buffer() As Byte
    'Buffer = Texto 'se meten de dos en dos las letras ??? sera por algo de ascii vs unicode
    
    ReDim Buffer(Len(Texto) - 1)
    For F = 1 To Len(Texto)
        Buffer(F - 1) = Asc(Mid(Texto, F, 1))
    Next F

    Dim xClave() As Byte
    'xClave = Clave
    
    ReDim xClave(Len(Clave) - 1)
    For F = 1 To Len(Clave)
        xClave(F - 1) = Asc(Mid(Clave, F, 1))
    Next F
    
    'Encripto
    
    Dim Char1 As Integer 'Caracter Original
    Dim Char2 As Integer 'Caracter ya Modificado (char1+char3) o (char1-char3)
    Dim Char3 As Integer 'Caracter de la Clave
    
    'Voy dando vueltas por la clave asi que necesito un indice
    Dim ContadorClave As Integer 'Indice de la clave
    ContadorClave = 0
    
    Dim I As Long
    Dim NuevoDato() As Byte
    
    ReDim NuevoDato(Len(Texto) - 1)
    
    For I = 0 To UBound(Buffer)
        Char1 = Buffer(I)
        Char3 = xClave(ContadorClave)
        If Invertido = True Then
            Char2 = Char1 - Char3
        Else
            Char2 = Char1 + Char3
        End If

        If Char2 < 0 Then
            Char2 = 256 + Char2
        End If
    
        If Char2 > 255 Then
            Char2 = Char2 Mod 256
        End If
    
        NuevoDato(I) = Char2
        
        ContadorClave = ContadorClave + 1
        If ContadorClave > UBound(xClave) Then ContadorClave = 0
    Next I
    
    Dim tRES As String
    For F = 0 To UBound(NuevoDato)
        tRES = tRES + Chr(NuevoDato(F))
    Next F
    
'    Dim Ver As String
'    For F = 0 To UBound(Buffer)
'        Ver = Ver + Chr(Buffer(F)) + " - " + Chr(NuevoDato(F)) + " * "
'    Next F
'    MsgBox Ver

    pinchilon = tRES
    
End Function

Private Function tLt(Fil, sep, qi As String) As String
    'desencriptar cada una de las partes (cada separador es diferente)
    Dim ftLt As String 'temporal para leerlo todo
    ftLt = ""
    LEE = GetText(CStr(Fil))
    
    'leer cada dato
    Dim SP() As String
    SP = Split(LEE, sep)
    Dim H As Long
    For H = 0 To UBound(SP)
        ftLt = ftLt + pinchilon(SP(H), qi, True) + Chr(5) 'nuevo separador
    Next H
    
    tLt = ftLt
End Function

Private Sub PlayInfo()
    'reinicio todos los contadores
    'indices sin uso para inicializar las matrices
    Erase Equipos: Erase Ritmo: Erase DISCO: Erase Cancion 'si no hago esto el valor cero de la matriz queda con valores !!!
    ReDim Equipos(0): ReDim Ritmo(0): ReDim DISCO(0): ReDim Cancion(0)
    CleanTELEM Platilla 'asegurarse que empieze de cero !!!
    
    Dim FT As String 'texto completo de cada archivo
    Dim SP_FT() As String 'cada elemento de cada archivo
    Dim SP2() As String 'cada parte de cada elemento de los archivos
    
    'ver los filtros aplicados
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    If fso.FileExists(AP + "sf\stats.dss") = False Then Exit Sub
    
    FT = GetText(AP + "sf\stats.dss")
    SP_FT = Split(FT, Chr(3))
    
    'ver que sea un archivo de los mios
    If UBound(SP_FT) > 0 Then
        Dim J As Long
        For J = 0 To UBound(SP_FT)
            porc Round(J / (UBound(SP_FT)), 2), "Sumando"
            SP2 = Split(SP_FT(J), Chr(5))
            'algunos ejemplos
            
            'Platilla para escuchar musica
            '0 1475245 + _
             1 5469854 + _
             2 E + _
             3 0.5 + _
             4 001921F5682A|00000000|BFEBFBFF00000F49|409000F|WD-WMAM98125093| + _
             5 200801022355+ _
             6 rock + _
             7 almafuerte 10 años + _
             8 sirva otra guelta pulpero
                                
            'asegurarase que sea un renglon valido
            If UBound(SP2) >= 8 Then
                Dim NewI(5) As Long
                
                If UCase(SP2(2)) = "SHIT" Then
                    'tiene fallas el registro
                    
                Else
                    Dim T9 As String 'primera letra que indica que es
                    T9 = UCase(SP2(2))
                    SP2(3) = Replace(SP2(3), SDno, SD)
                    
                    '************SAQUE ESTADISTICAS DE DISCOS Y CANCIONES
                    '************ME PARECEN QUE DEBEN IR EN OTRO LADO
                    
                    'agregar el equipo
                    If Len(SP2(4)) > 2 Then NewI(1) = AgregarTElement(Equipos, SP2(4))
                    'registrar contador y plata para el ritmo (origen de disco)
                    If Len(SP2(6)) > 2 Then NewI(2) = AgregarTElement(Ritmo, SP2(6))
                    'registrar contador y plata para el disco
                    'If Len(SP2(7)) > 2 Then NewI(3) = AgregarTElement(Disco, SP2(7))
                    'registrar contador y plata para la cancion
                    'If Len(SP2(8)) > 2 Then NewI(4) = AgregarTElement(Cancion, SP2(8) + " DE (" + SP2(7) + ")")
                    
                    If T9 = "E" Then
                        Platilla.Escuchados = Platilla.Escuchados + CSng(SP2(3))
                        Equipos(NewI(1)).Escuchados = Equipos(NewI(1)).Escuchados + CSng(SP2(3))
                        Ritmo(NewI(2)).Escuchados = Ritmo(NewI(2)).Escuchados + CSng(SP2(3))
                        'Disco(NewI(3)).Escuchados = Disco(NewI(3)).Escuchados + CSng(SP2(3))
                        'Cancion(NewI(4)).Escuchados = Cancion(NewI(4)).Escuchados + CSng(SP2(3))
                    End If
                    
                    If T9 = "U" Then
                        Platilla.VentasU = Platilla.VentasU + CSng(SP2(3))
                        Equipos(NewI(1)).VentasU = Equipos(NewI(1)).VentasU + CSng(SP2(3))
                        Ritmo(NewI(2)).VentasU = Ritmo(NewI(2)).VentasU + CSng(SP2(3))
                        'Disco(NewI(3)).VentasU = Disco(NewI(3)).VentasU + CSng(SP2(3))
                        'Cancion(NewI(4)).VentasU = Cancion(NewI(4)).VentasU + CSng(SP2(3))
                    End If
                    
                    If T9 = "B" Then
                        Platilla.VentasB = Platilla.VentasB + CSng(SP2(3))
                        Equipos(NewI(1)).VentasB = Equipos(NewI(1)).VentasB + CSng(SP2(3))
                        Ritmo(NewI(2)).VentasB = Ritmo(NewI(2)).VentasB + CSng(SP2(3))
                        'Disco(NewI(3)).VentasB = Disco(NewI(3)).VentasB + CSng(SP2(3))
                        'Cancion(NewI(4)).VentasB = Cancion(NewI(4)).VentasB + CSng(SP2(3))
                    End If
                    
                    If T9 = "C" Then
                        Platilla.VentasC = Platilla.VentasC + CSng(SP2(3))
                        Equipos(NewI(1)).VentasC = Equipos(NewI(1)).VentasC + CSng(SP2(3))
                        Ritmo(NewI(2)).VentasC = Ritmo(NewI(2)).VentasC + CSng(SP2(3))
                        'Disco(NewI(3)).VentasC = Disco(NewI(3)).VentasC + CSng(SP2(3))
                        'Cancion(NewI(4)).VentasC = Cancion(NewI(4)).VentasC + CSng(SP2(3))
                    End If
                    
                End If
                
            End If
        Next J

    End If

    '**********************************************************
    '**********************************************************
    'este codigo ponia todo el resumen en un txtbox pero es feo, anda ok
    '**********************************************************
    'ahora esta cargado todo se puede emitir el informe
'    Dim info As String
'
'    info = "Resumen de movimiento" + vbCrLf + _
'        "Equipos cargados: " + CStr(UBound(Equipos)) + vbCrLf + _
'        "Recaudacion total: $" + _
'        CStr(Round(Platilla.Escuchados + Platilla.VentasB + Platilla.VentasU + Platilla.VentasC, 2)) + vbCrLf + _
'        "  Recaudacion por reproducción: $" + CStr(Platilla.Escuchados) + vbCrLf + _
'        "  Recaudacion ventas Bluetooth: $" + CStr(Platilla.VentasB) + vbCrLf + _
'        "  Recaudacion ventas por usb:   $" + CStr(Platilla.VentasU) + vbCrLf + _
'        "  Recaudacion ventas en CDs:   $" + CStr(Platilla.VentasC) + vbCrLf
'
'    info = info + vbCrLf + "Detalle por equipo" + vbCrLf + vbCrLf
'    Dim H As Long
'
'
'    For H = 1 To UBound(Equipos)
'        porc Round(H / (UBound(Equipos) + 1), 2), "Calculando"
'        info = info + "Nombre: " + Equipos(H).Nombre + vbCrLf + _
'            "Recaudacion total: $" + _
'        CStr(Round(Equipos(H).Escuchados + Equipos(H).VentasB + Equipos(H).VentasU + Equipos(H).VentasC, 2)) + vbCrLf + _
'        "  Recaudacion por reproducción: $" + CStr(Equipos(H).Escuchados) + vbCrLf + _
'        "  Recaudacion ventas Bluetooth: $" + CStr(Equipos(H).VentasB) + vbCrLf + _
'        "  Recaudacion ventas por usb:   $" + CStr(Equipos(H).VentasU) + vbCrLf + _
'        "  Recaudacion ventas en CD:   $" + CStr(Equipos(H).VentasC) + vbCrLf
'    Next H
'
'    info = info + vbCrLf + "Detalle por ritmo" + vbCrLf + vbCrLf
'    For H = 1 To UBound(Ritmo)
'        porc Round(H / (UBound(Ritmo) + 1), 2), "Calculando Ritmo"
'        info = info + "Ritmo: " + Ritmo(H).Nombre + vbCrLf + _
'            "Recaudacion total: $" + _
'        CStr(Round(Ritmo(H).Escuchados + Ritmo(H).VentasB + Ritmo(H).VentasU + Ritmo(H).VentasC, 2)) + vbCrLf + _
'        "  Recaudacion por reproducción: $" + CStr(Ritmo(H).Escuchados) + vbCrLf + _
'        "  Recaudacion ventas Bluetooth: $" + CStr(Ritmo(H).VentasB) + vbCrLf + _
'        "  Recaudacion ventas por usb:   $" + CStr(Ritmo(H).VentasU) + vbCrLf + _
'        "  Recaudacion ventas en CD:   $" + CStr(Ritmo(H).VentasC) + vbCrLf
'    Next H
'
'    info = info + vbCrLf + "Detalle por discos" + vbCrLf + vbCrLf
'    For H = 1 To UBound(Disco)
'        porc Round(H / (UBound(Disco) + 1), 2), "Calculando Discos"
'        info = info + "Disco: " + Disco(H).Nombre + vbCrLf + _
'            "Recaudacion total: $" + _
'        CStr(Round(Disco(H).Escuchados + Disco(H).VentasB + Disco(H).VentasU + Disco(H).VentasC, 2)) + vbCrLf + _
'        "  Recaudacion por reproducción: $" + CStr(Disco(H).Escuchados) + vbCrLf + _
'        "  Recaudacion ventas Bluetooth: $" + CStr(Disco(H).VentasB) + vbCrLf + _
'        "  Recaudacion ventas por usb:   $" + CStr(Disco(H).VentasU) + vbCrLf + _
'        "  Recaudacion ventas en CD:   $" + CStr(Disco(H).VentasC) + vbCrLf
'    Next H
'
'    info = info + vbCrLf + "Detalle por Canciones" + vbCrLf + vbCrLf
'    For H = 1 To UBound(Cancion)
'        porc Round(H / (UBound(Cancion) + 1), 2), "Calculando Canciones"
'        info = info + vbCrLf + "Cancion: " + Cancion(H).Nombre + vbCrLf + _
'            "Recaudacion total: $" + _
'            CStr(Round(Cancion(H).Escuchados + Cancion(H).VentasB + Cancion(H).VentasU + Cancion(H).VentasC, 2)) + _
'            " (" + CStr(Cancion(H).Escuchados) + " / " + _
'            CStr(Cancion(H).VentasB) + " / " + _
'            CStr(Cancion(H).VentasU) + " / " + CStr(Cancion(H).VentasC) + ")"
'    Next H
    
    'Text1.Text = info
    
    '**********************************************************
    '**********************************************************
    '**********************************************************
End Sub

Private Sub porc(xpo As Single, Dinfo As String)
    If xpo <= 1 Then xpo = 100 * xpo
    If xpo = 0 Then
        Me.Caption = "Estadisticas 3PM "
    Else
        Me.Caption = "Estadisticas 3PM      -- " + CStr(xpo) + " % (" + Dinfo + ")"
    End If
    
    If xpo >= 100 Then
        picPorc.Visible = False
    Else
        picPorc.Visible = True
        picPorc.Width = (xpo * Me.Width) / 100
        picPorc.Refresh
        Me.Refresh
    End If
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    Me.Width = Screen.Width
    Me.Left = 0
    G1.Width = Me.Width - 200
    G2.Width = G1.Width
End Sub

Private Sub CleanTELEM(tel As tElement)
    tel.Escuchados = 0
    tel.Nombre = ""
    tel.VentasB = 0
    tel.VentasC = 0
    tel.VentasU = 0
End Sub

