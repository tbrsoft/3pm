Attribute VB_Name = "SuperLicencia"
'para obtener info del procesador
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

Public Declare Sub GetMem1 Lib "msvbvm60.dll" (ByVal MemAddress As Long, Var As Byte)

Public Function GetGuidSL() As String
    'obtener identificador unico de equipo
    'copymem+reserved
    '-----------------
    'COPYMEM
    '-----------------
    Dim CPM As String
    CPM = CStr(SumaCHRtxt(GetBIOSDate))
    
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
        ArchUniqueAzar = GPF("rempres44")
        'ver si ya se genero el archivo para esta formateada
        If fso.FileExists(ArchUniqueAzar) = False Then
            Dim A As Long
            Randomize Timer
            A = Int(Rnd * 10000)
            A = 111000000 + A
            
            Set TE = fso.CreateTextFile(ArchUniqueAzar, True)
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
    
    Dim FullCOD As String
    FullCOD = CPM + "-" + RSV
    GetGuidSL = FullCOD
End Function

Private Function GetBIOSDate() As String
  Dim P As Byte, MemAddr As Long, sBios As String
  Dim I As Integer
  'comienzo del serial de la BIOS &HFE0C0
  MemAddr = &HFE000
  For I = 0 To 331
      Call GetMem1(MemAddr + I, P)
      'get printable characters
      If P > 31 And P <= 128 Then
      sBios = sBios & Chr$(P)
    End If
  Next I
  GetBIOSDate = sBios
End Function

Public Function SumaCHRtxt(TXT As String) As Long
    'sumar el valor CHR de los caracteres de un texto
    Dim Caracter As String
    Dim TMP As Long
    
    For J = 1 To Len(TXT)
      Caracter = Mid(TXT, J, 1)
      TMP = TMP + Asc(Caracter)
    Next J
    SumaCHRtxt = TMP
End Function
