VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKEYS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'el sistema de claves debe tener:
    'Muchas claves distintas para futuros usos
    'Mas de una clave por tipo de licencia
    'Codigo de PC unico para dar seguridad. NO ES UNICO EL DEL WBMEME PROCID!!!!! FUCK!!!!!!
    'No importa que sea 98 o ME debe andar igual!!! con FF(98) y todo!!!
    
    

Private FSO As New Scripting.FileSystemObject

Private Declare Function GetLocaleInfo Lib "kernel32" Alias _
    "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, _
    ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Declare Sub GetMem1 Lib "msvbvm60.dll" (ByVal _
   MemAddress As Long, var As Byte)

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Private Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

'tipos de licencia que se pueden cargar
Public Enum TypeLic
    aSinCargar = 0 'todavia no puso nada
    BErronea = 1 'cuando ya erro 3 o mas veces
    CGratuita = 2 'primera que me puede pedir
    DMinima = 3
    EComun = 4
    FPremium = 5
    GFull = 6
    HSuperLicencia = 7
End Enum

Private FileWithDatosClave As String 'archivo de la clave ingresada
Private SYSfolder As String 'carpeta de sistema
Private WINfolder As String 'carpeta de Windows

Private DllHabilitada As Boolean 'se puedo o no usar la DLL
Private mAsig As String
Private mLicencia As TypeLic

'para el pais e idioma
Private Const LOCALE_USER_DEFAULT = &H400
Private Const LOCALE_SENGCOUNTRY = &H1002 '  English name of country
Private Const LOCALE_SENGLANGUAGE = &H1001  '  English name of language
Private Const LOCALE_SNATIVELANGNAME = &H4  '  native name of language
Private Const LOCALE_SNATIVECTRYNAME = &H8  '  native name of country

Private mVariacionClave As Long 'variacion de la clave. es 0 si no hay clave validada

Public Property Get FileData() As String
    FileData = FileWithDatosClave
End Property

Private Sub Class_Initialize()
    SYSfolder = FSO.GetSpecialFolder(SystemFolder)
    WINfolder = FSO.GetSpecialFolder(WindowsFolder)
    If Right(SYSfolder, 1) <> "\" Then SYSfolder = SYSfolder + "\"
    If Right(WINfolder, 1) <> "\" Then WINfolder = WINfolder + "\"
        
    FileWithDatosClave = SYSfolder + "dciLib22.dll" 'DatosClaveIngresada.DLL
    
    
    'Dim N As Long
    'Dim H As String
    'Dim HtL As String
    'N = Val(InputBox("Numero?:"))
    'H = Hex(N)
    'HtL = HEXtoLONG(H)
    'MsgBox "N= " + CStr(N) + vbCrLf + _
        "H= " + H + vbCrLf + _
        "HtL= " + HtL
    
    
End Sub

Public Property Let ClaveDLL(NewClave As String)
    If NewClave = "" Then
        DllHabilitada = True
    Else
        If NewClave = "ashjdklahsJKLHASL65456456456" Then DllHabilitada = True
    End If
End Property

Public Function ReleerLICENCIA() As TypeLic
    'para saber si o si vuelve a leer. Usarlo solo una vez por secion de 3PM v5
    ReleerLICENCIA = IngresaClave(LeerClaveIngresadaAntes)
End Function

Public Function LICENCIA() As TypeLic
    'lee lo grabado para que sea mas rapido
    LICENCIA = mLicencia
End Function


Public Function IngresaClave(ClaveIngresadaConSeparadores As String) As TypeLic
    'aqui el tipo ingresa y tengo que saber si es alguna de las validas
    'devuelve el tipo de licencia que envoco
    
    Dim TmpLic As TypeLic
    
    'si no puso clave deben venir los separadores
    If ClaveIngresadaConSeparadores = "-------" Or _
        ClaveIngresadaConSeparadores = "" Then
        
        mLicencia = aSinCargar
        IngresaClave = aSinCargar
        Exit Function
        
    End If
    'si despues no entra a ninguna es erronea
    TmpLic = BErronea
    
    Dim A As TypeLic
    Dim B As Long
    
    mVariacionClave = 0
    
    For A = 2 To 7 'todos los tipos de licencia que dependen de claves
        '--------------
        'dar vueltas para que pueda pasar por las 50 variaciones de cada tipo de clave
        '--------------
        
        For B = 1 To 50
        
        
        
            If ClaveIngresadaConSeparadores = CLAVE(A, B) Then
            
                mVariacionClave = B 'variacion elegida
                
                Select Case A
                    Case 2
                        TmpLic = CGratuita
                    Case 3
                        TmpLic = DMinima
                    Case 4
                        TmpLic = EComun
                    Case 5
                        TmpLic = FPremium
                    Case 6
                        TmpLic = GFull
                    Case 7
                        TmpLic = HSuperLicencia
                End Select
                
                'ya encontro, al pedo seguir...
                GoTo Listo
                
            End If
        Next B
        
    Next A
Listo:
    mLicencia = TmpLic
    IngresaClave = TmpLic
    'escribir lo que el escribio en un archivo
    '!Encriptado!!!! (un caracter valido - otro inventado)
    'solo es para que no lo busquen por el texto de la clave
    If FSO.FileExists(FileWithDatosClave) Then FSO.DeleteFile FileWithDatosClave, True
    
    Dim TE As TextStream
    Dim ClaveEncr As String
    ClaveEncr = SimpleEncript(ClaveIngresadaConSeparadores)
    Set TE = FSO.CreateTextFile(FileWithDatosClave, True)
        TE.WriteLine ClaveEncr
    TE.Close

End Function

Public Property Get VariacionClave()
    VariacionClave = mVariacionClave
End Property

Public Function LeerClaveIngresadaAntes() As String
    
    If FSO.FileExists(FileWithDatosClave) = False Then
        LeerClaveIngresadaAntes = ""
    Else
        Dim TE As TextStream
        Dim ClaveEncr As String
        
        Set TE = FSO.OpenTextFile(FileWithDatosClave, ForReading, False)
            ClaveEncr = TE.ReadLine
        TE.Close
        
        LeerClaveIngresadaAntes = SimpleUnEncript(ClaveEncr)
        
    End If

End Function

Private Function SimpleUnEncript(TXT As String) As String
    'agregar un caracter cada un caracter del original
    'solo para que no se pueda buscar por texto
    Dim Largo As Long, LT As String, TMP As String
    Largo = Len(TXT)
    TMP = ""
    For A = 1 To Largo
        
        LT = Mid(TXT, A, 1)
        'solo los caracteres impares valen
        If A / 2 <> A \ 2 Then
            'y los acumulo
            TMP = TMP + LT
        End If
    Next
    SimpleUnEncript = TMP
End Function

Private Function SimpleEncript(TXT As String) As String
    'agregar un caracter cada un caracter del original
    'solo para que no se pueda buscar por texto
    Dim Largo As Long, LT As String, TMP As String, s As Integer
    Largo = Len(TXT)
    TMP = ""
    For A = 1 To Largo
        'primer caracter original
        LT = Mid(TXT, A, 1)
        'y otro al azar
        Randomize Timer
        s = Int(Rnd * 9)
        'los sumo
        LT = LT + CStr(s)
        'y los acumulo
        TMP = TMP + LT
    Next
    SimpleEncript = TMP
End Function

Public Property Get UniquePC() As String
    If DllHabilitada = False Then
        UniquePC = "Información no disponible"
        Exit Property
    End If
    
    UniquePC = GetCountry + "." + SimpleEncript(GetProcID) + "." + SimpleEncript(GetInfoReserved) + "." + SimpleEncript(GetBiosNumber)
End Property


Public Property Get CLAVE(nClave As TypeLic, Variacion As Long, Optional ForUnique As String = "") As String
    'nClave es el numero de clave (FULL, DEMO, SL, etc)
    'Variacion es una de las 20 claves válidas para cada nClave para cada PC. _
        Esto hace que si me manda el mismo numero pueda yo darle claves distintas
    'ForUnique es para el codigo de PC que se solicita la clave
    
    If DllHabilitada = False Then
        CLAVE = "Información no disponible"
        Exit Property
    End If
    
    Dim Uniq As String
    'ver si es para esta PC o para otra
    If ForUnique = "" Then
        Uniq = UniquePC
    Else
        Uniq = ForUnique
    End If
    
    '--------------------------------------------
    'del numero de PC solo tengo en cuenta el BiosDate y el Reserved _
        ya que la otra bosta no se si es confiable!
    
    Uniq = GetOldFromNew(Uniq)
    'Dim PartesUniq() As String
    'PartesUniq = Split(Uniq, ".")
    ''ahora transformo el uniq en HEX a mi viejo y querido formato bios-reserved
    'Uniq = HEXtoLONG(PartesUniq(4)) + "-" + HEXtoLONG(PartesUniq(3))
    '--------------------------------------------
    
        
    'carcteres validos
    Dim ChrValid As String
    ChrValid = "789abcdeyzABgggGcdeyzABgGc3636363a6rtHjKkEdeyzABgGGGCDEFc262688WW876GUP44deyzABCGfgcdeyzABChkmnpcdeyzABCqr2346tucdeyz" + _
        "c262688WW876GUP44deyzABCGfgcdeyzABChkmnpcdeyzABCqr2346tucdeyzeBCvExHceEyzAeE234gKcdGyzAwCMxWzAf346FCDEbG23B6H6NP62336R3892" + _
        "ABCvwxHcdeyzABC2346KcdeyzABCMxyzA2346BCDEFG2346HKX46abcdXf23X6ghXmnpqXT4TXTVW23T6FYTF346GtuvGM23G6NPQG234GTUGWNPQ2346R7892" + _
        "346abcdeAkCvwKHcdkyzABK2346kcdeKzABdMxyDA234dBCDtFG2T46HtNPQT346T7892f2346ghkmnpq2346TUVW2346XYZ2346rtuvwM2346NPQR2346TUVW" + _
        "AkC3y6aycdey234yghkynpqy346yUVWy34yXYZy34yrtyvwy23y6NPYR2Y46YUVYvwKHcdkyzABK2346kcdeKzABdMxyDA234dBCDtFG2T46HtNPQT346T7892" + _
        "eBCvExHceEyzAeE234gKcdGyzAwCMxWzAf346FCDEbG23B6HeVCvVxHVeEVzAeN23MgKcdMFzAFCMxMzAfV46VCDECG2CBCH6NZ62Z36Z38Z26NP62336R3892" + _
        "ch6268hWWh76GhP44heyzaBCafgcaeyzA9Ch9mn9cde9zAB9qr2f46tfcFFFw" + _
        "ABgvwvHWWeyzWWC2346KjdjyjABkMxkzAM34mBCmEFd23d6HcNPc23c6Rccc2" + _
        "3y6aycdey234yghkynpqy346yUVWy34yXYZy34yrtyvwy23y6NPYR2Y46YUVY" + _
        "AbCvbKHcbBBBABK23sskcdeKeABdexyDe234hBCDtFGtT4tHtNtQTt46T7897" + _
        "Tw6Tw8TwW876GUPw4dewzABChfghdehzAhChhmhphFeFzFBkqk2F4FtFkdeyz" + _
        "3BC3wx3cdeTzAT323T6Tcd7yzAB7Mxy9A2946FCD9FGa34aAKNaQa34aRa892" + _
        "X46abcdXf23X6ghXmnpqXT4TXTVW23T6FYTF346GtuvGM23G6NPQG234GTUGW" + _
        "AHCvHKHcHkyzHBK2H46kcHeKzAHdMFyHA23HdTTDTTG2N46HNNPQN346N78NN" + _
        "eVCvVxHVeEVzAeN23MgKcdMFzAFCMxMzAfV46VCDECG2CBCH6NZ62Z36Z38Z2" + _
        "Qh6Q68QWQhQ6GQPFQFeQFaQCafgPaePzAPCP9Pn9PTTTzPB9Pr2fR6tRcFFRR" + _
        "ARgvRvHWReFzRWCFR46KRdjyRABkRxkzRM3RmBRmEHd2Hd6HcNP72376R7c77" + _
        "8y68ycTTT284yg7k7np7y749y9V9y949X8Zy344rt4vwy43T64PYR4Y46ZUVY" + _
        "XTCTbTHcXBBXABX23XskXdeKXAXdXxXDX2X4hBKDtKGKTKtKtKtQ8t8687898" + _
        "Qh6Q68QWQhQ6GQPFQFeQFaQCafgPaePzAPCP9Pn9PTTTzPB9Pr2fR6tRcFFRR" + _
        "AWgvWvHWReWzRWCFr46rRdryRArkRrkzuM3umBumuHu2Hy6HyNy7y376y7c77" + _
        "6y68yc6T6284ug7kjnp7h749r9V9k949k8Zyp449t4vfy43d64PYa4Ya6ZaVY" + _
        "rTCTbrHcXyBXyBXy3XskydeKtAXdhxXDXgX4hfKDtKdKTKdKtKdQ8td687999" + _
        "2346XYZZZZZZABC2346KcdeyzABCMxyzA2mnpq2346TUV3636363a6rtHjKk346TU"
    
    Dim Largo As Long
    Largo = Len(Uniq)
    'pasar por todos los caracteres
    Dim LT As String 'letras
    Dim Res As String 'resultado parcial
    Dim ResFull As String 'resultado acumulado
    ResFull = ""
    mAsig = ""
    Dim pos As Long
    'otro numero mas que es fijo y que es otra posibilidad mas
    Dim OtroMas As Long
    OtroMas = 0
    'el do y loop es simplemente para segurar 40 caracteres de clave
    Do
        For A = 1 To Largo
            OtroMas = OtroMas + 3
            LT = Mid(Uniq, A, 1)
            'si es el gion o algo raro pasar de largo
            If Not IsNumeric(LT) Then GoTo SIG
            'a es de 1 a 15 (aprox)
            'val(LT) es de 0 a 9
            'variacion es de 1 a 50
            'otro mas es de 3 hasta 0??? aprox??
            'nClaves es de 2 a 7 que son las claves que se pueden pedir
            '--------------------------
            'ASEGURARME QUE NO HAYA DOS CLAVES QUE SIRVAN PARA DOS COSAS DISTINTAS!!
            '--------------------------
            Dim PosInChars As Long
            '--------------------------
            'ASEGURARME QUE NO HAYA DOS CLAVES QUE SIRVAN PARA DOS COSAS DISTINTAS!!
            'PARA TODOS OtroMas,A y Val(LT) son iguales
            'la diferencia deb estar en nClave y Variacion
            '--------------------------
            'me aseguro ademas que sea positivo
            PosInChars = Abs(nClave * 100 + Variacion - (OtroMas + A + Val(LT)))
            'desde 200+1- (120+15+9)= 57 ---------menor sumando y mayior restando
            'hasta 700+50-(3+1+0)=746
                       
            'ver que la posicion copn el largo a usar esten dentro de ChValid
            Do
                'solo antes de la mitad del largo(chrValid) para asegurar siempre caracteres
                If PosInChars < Len(ChrValid) / 2 Then Exit Do
                PosInChars = PosInChars - 1
            Loop
            Res = Mid(ChrValid, PosInChars, 1)
            
            ResFull = ResFull + Res
            
            'grabar las asignaciones
            'HABILITARLAS SI LAS NECESITO
            mAsig = mAsig + LT + "=" + Res + vbCrLf
            
            If Len(ResFull) >= 40 Then
                ResFull = Left(ResFull, 40)
                Exit Do
            End If
SIG:
        Next
    Loop
    CLAVE = MostraDeA5(ResFull)
    
End Property

Public Function GetOldFromNew(OLD As String) As String
    On Local Error GoTo NoGetNew
    '--------------------------------------------
    'del numero de PC solo tengo en cuenta el BiosDate y el Reserved _
        ya que la otra bosta no se si es confiable!
    Dim PartesUniq() As String
    PartesUniq = Split(OLD, ".")
    'esta matriz no siempre tiene 5, puede ser 98!! y tener solo FF.Reserved.Bios!!!!
    'esto daría error
    Dim PtosImport As Long 'puntos importantes
    PtosImport = UBound(PartesUniq) - 1
    'ahora transformo el uniq en HEX a mi viejo y querido formato bios-reserved
    GetOldFromNew = HEXtoLONG(SimpleUnEncript(PartesUniq(PtosImport + 1))) + "-" + HEXtoLONG(SimpleUnEncript(PartesUniq(PtosImport)))
    '--------------------------------------------
    Exit Function
    
NoGetNew:
    GetOldFromNew = "00"
End Function

Public Property Get Asignaciones() As String
    Asignaciones = mAsig
End Property

Public Property Get ExistsDataLic() As Boolean
    If DllHabilitada = False Then
        ExistsDataLic = False
        Exit Property
    End If
    'solo si existen los 2 archivos
    ExistsDataLic = False
    
    If FSO.FileExists(FileWithDatosClave) Then
        ExistsDataLic = True
    End If
    
End Property

Private Function GetBIOSDate() As String
  Dim P As Byte, MemAddr As Long, sBios As String
  Dim i As Integer
  'start of bios serial number ?&HFE0C0
  MemAddr = &HFE000
  For i = 0 To 331
      Call GetMem1(MemAddr + i, P)
      'get printable characters
      If P > 31 And P <= 128 Then
      sBios = sBios & Chr$(P)
    End If
  Next i
  GetBIOSDate = sBios
End Function

Private Function GetBiosNumber() As String
    If DllHabilitada = False Then
        GetBiosNumber = "Información no disponible"
        Exit Function
    End If
    
    Dim CPM As String
    CPM = CStr(SumaCHRtxt(GetBIOSDate))
    'pasar a HEX para que no sea igual a antes
    CPM = Hex(CLng(CPM))
    GetBiosNumber = CPM
End Function

Private Function GetInfoReserved() As String
    If DllHabilitada = False Then
        GetInfoReserved = "Información no disponible"
        Exit Function
    End If
    
    '-------------------
    'RESERVED
    '-------------------
    Dim RSV As String
    
    Dim INFO As SYSTEM_INFO
    
    GetSystemInfo INFO
    
    Dim GUIDtmp As String 'no es guid, es un valor unico para cada PC
    'este reserved es un numero entre 50.000.000 y 140.000.000
    RSV = CStr(INFO.dwReserved)
    If Len(RSV) < 3 Then
        'no es compatible en esta PC
        'meter un valor fijo
        RSV = "FF"
    Else
        'pasar a hex para que no sea igual a antes
        RSV = Hex(CLng(RSV))
    End If
    GetInfoReserved = RSV
    
End Function

Private Function GetProcID() As String
    
    If DllHabilitada = False Then
        GetProcID = "Información no disponible"
        Exit Function
    End If
    
    Dim TMP As String
    
    On Error GoTo NoWBmem
    
    Dim ObjSet As SWbemObjectSet
    Dim SERV As SWbemServices
    Set SERV = GetObject("WinMgmts:")
    Set ObjSet = Nothing
    Set ObjSet = SERV.InstancesOf("Win32_Processor")
    If ObjSet.Count = 1 Then
        For Each MICRO In ObjSet
            TMP = CStr(NN(MICRO.ProcessorType)) + "." + _
                CStr(NN(MICRO.Revision)) + "." + _
                CStr(NN(MICRO.ProcessorId))
            
        Next
    End If
NoInst:
    
    'si es el 98 como no tiene el procID dejara en FF!!
    GetProcID = TMP
    
    Exit Function
        
NoWBmem:
    TMP = "FF"
    GoTo NoInst
End Function

Private Function SumaCHRtxt(TXT As String) As Long
    'sumar el valor CHR de los caracteres de un texto
    Dim Caracter As String
    Dim TMP As Long
    
    For j = 1 To Len(TXT)
      Caracter = Mid(TXT, j, 1)
      TMP = TMP + Asc(Caracter)
    Next j
    SumaCHRtxt = TMP
End Function

Private Function LeerArch1Linea(Arch As String) As String
    'If Dir(Arch) = "" Then
    If FSO.FileExists(Arch) = False Then
        LeerArch1Linea = "No existe archivo"
        Exit Function
    End If
    'Dim Libre As Long, TMP As String
    'Libre = FreeFile
    'Open Arch For Input As Libre
    Set TE = FSO.OpenTextFile(Arch, ForReading, False)
        'Line Input #Libre, TMP
    LeerArch1Linea = TE.ReadLine
    'Close Libre
    TE.Close
    'LeerArch1Linea = TMP
End Function

Private Function MostraDeA5(TXT As String)
    Dim c As Long, Letra As String, newTXT As String
    c = 0
    Do While c < Len(TXT)
        Letra = Mid(TXT, c + 1, 5)
        newTXT = newTXT + Letra
        c = c + 5
        If c < Len(TXT) Then newTXT = newTXT + "-"
    Loop
    MostraDeA5 = newTXT
End Function

Private Function HEXtoLONG(N As String)
    'recibe el hex en str y devuelve un numero en str
    
    Dim Letra As String
    Dim c As Long
    Dim NumeroActual As Long
    Dim ACUM ' As Double
    For c = 1 To Len(N)
        Letra = Mid(N, c, 1)
        Select Case Letra
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                NumeroActual = Val(Letra)
            Case "A"
                NumeroActual = 10
            Case "B"
                NumeroActual = 11
            Case "C"
                NumeroActual = 12
            Case "D"
                NumeroActual = 13
            Case "E"
                NumeroActual = 14
            Case "F"
                NumeroActual = 15
        End Select
        Dim ToSum ' As Double
        ToSum = NumeroActual * (16 ^ (Len(N) - c))
        ACUM = ACUM + ToSum
        'Label10 = Label10 + "LETRA: " + Letra + "=" + CStr(ToSum) + vbCrLf
        
    Next
    
    HEXtoLONG = CStr(ACUM)
End Function

Private Function NN(Val, Optional DEfault = "NULO")
    'No Nulo
    If IsNull(Val) Then
        NN = DEfault
    Else
        NN = Val
    End If
End Function

Public Function GetCountry() As String
    Dim TMP As String
    TMP = GetInfo(LOCALE_SENGCOUNTRY) ' + "(" + GetInfo(LOCALE_SNATIVECTRYNAME) & ")"
    
    GetCountry = TMP
End Function

Private Function GetInfo(ByVal lInfo As Long) As String
    Dim Buffer As String, Ret As String
    Buffer = String$(256, 0)
    Ret = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, Buffer, Len(Buffer))
    If Ret > 0 Then
        GetInfo = Left$(Buffer, Ret - 1)
    Else
        GetInfo = ""
    End If
End Function


