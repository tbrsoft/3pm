Attribute VB_Name = "Mod"
Public Const LOCALE_USER_DEFAULT = &H400
Public Const LOCALE_SENGCOUNTRY = &H1002 '  English name of country
Public Const LOCALE_SENGLANGUAGE = &H1001  '  English name of language
Public Const LOCALE_SNATIVELANGNAME = &H4  '  native name of language
Public Const LOCALE_SNATIVECTRYNAME = &H8  '  native name of country
Public Const LOCALE_ICOUNTRY = &H5 'Country/region code, based on _
    international phone codes, also referred to as IBM country codes. _
    The maximum number of characters allowed for this string is si

Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Declare Sub GetMem1 Lib "msvbvm60.dll" (ByVal _
   MemAddress As Long, var As Byte)

Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Public Type SYSTEM_INFO
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

Public Function GetBIOSDate() As String
  Dim p As Byte, MemAddr As Long, sBios As String
  Dim i As Integer
  'start of bios serial number ?&HFE0C0
  MemAddr = &HFE000
  For i = 0 To 331
      Call GetMem1(MemAddr + i, p)
      'get printable characters
      If p > 31 And p <= 128 Then
      sBios = sBios & Chr$(p)
    End If
  Next i
  GetBIOSDate = sBios
End Function

Public Function NN(Val, Optional DEfault = "NULO")
    'No Nulo
    If IsNull(Val) Then
        NN = DEfault
    Else
        NN = Val
    End If
End Function

Public Function GetBiosNumber()
    Dim CPM As String
    CPM = CStr(SumaCHRtxt(GetBIOSDate))
    GetBiosNumber = CPM
End Function

Public Function GetInfoProc() As String
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
        'generar un numero al azar y dejarlo grabado
        'si se formatea debera pedir de vuelta
        'para que no me caguen estos numeros deberán empezar con 111.000.000
        Dim ArchUniqueAzar As String
        ArchUniqueAzar = SYSfolder + "\razaGUID.dll"
        'ver si ya se genero el archivo para esta formateada
        If FSO.FileExists(ArchUniqueAzar) = False Then
            Dim A As Long
            Randomize Timer
            A = Int(Rnd * 10000)
            A = 111000000 + A
            
            Set TE = FSO.CreateTextFile(ArchUniqueAzar, True)
            TE.WriteLine CStr(A)
            TE.Close
            RSV = CStr(A)
        End If
        'leer el archivo
        Dim UnicoAzar  As String
        UnicoAzar = LeerArch1Linea(ArchUniqueAzar)
        RSV = UnicoAzar
        
        'MsgBox "Esta PC no es compatible con las superlicencias. Pruebe en un Pentium I o superior"
        'Exit Function
    End If
    
    GetInfoProc = RSV
End Function

Public Function GetProcID() As String
    
    Dim TMP As String
    
    On Error GoTo NoWBmem
    
    Dim ObjSet As SWbemObjectSet
    Dim SERV As SWbemServices
    Set SERV = GetObject("WinMgmts:")
    Set ObjSet = Nothing
    Set ObjSet = SERV.InstancesOf("Win32_Processor")
    If ObjSet.Count = 1 Then
        For Each MICRO In ObjSet
            TMP = CStr(NN(MICRO.ProcessorId))
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

Public Function SumaCHRtxt(TXT As String) As Long
    'sumar el valor CHR de los caracteres de un texto
    Dim Caracter As String
    Dim TMP As Long
    
    For j = 1 To Len(TXT)
      Caracter = Mid(TXT, j, 1)
      TMP = TMP + Asc(Caracter)
    Next j
    SumaCHRtxt = TMP
End Function

Public Function LeerArch1Linea(Arch As String) As String
    If FSO.FileExists(Arch) = False Then
        LeerArch1Linea = "No existe archivo"
        Exit Function
    End If
    Set TE = FSO.OpenTextFile(Arch, ForReading, False)
    LeerArch1Linea = TE.ReadLine
    TE.Close
End Function

Public Function MostraDeA5(TXT As String)
    Dim C As Long, Letra As String, newTXT As String
    C = 0
    Do While C < Len(TXT)
        Letra = Mid(TXT, C + 1, 5)
        newTXT = newTXT + Letra
        C = C + 5
        If C < Len(TXT) Then newTXT = newTXT + "-"
    Loop
    MostraDeA5 = newTXT
End Function

Private Function HEXtoLONG(n As String)
    'recibe el hex en str y devuelve un numero en str
    
    Dim Letra As String
    Dim C As Long
    Dim NumeroActual As Long
    Dim ACUM ' As Double
    For C = 1 To Len(n)
        Letra = Mid(n, C, 1)
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
        ToSum = NumeroActual * (15 ^ (Len(n) - C))
        ACUM = ACUM + ToSum
        Label10 = Label10 + "LETRA: " + Letra + "=" + CStr(ToSum) + vbCrLf
        
    Next
    
    HEXtoLONG = CStr(ACUM)
End Function

Public Function GetInfo(ByVal lInfo As Long) As String
    Dim Buffer As String, Ret As String
    Buffer = String$(256, 0)
    Ret = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, Buffer, Len(Buffer))
    If Ret > 0 Then
        GetInfo = Left$(Buffer, Ret - 1)
    Else
        GetInfo = ""
    End If
End Function



