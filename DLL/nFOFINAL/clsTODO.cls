VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTODO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim ACUMTTT As String
Dim FSO As New Scripting.FileSystemObject

'sacado de 3PM con sistema de licencia anterior
Private Function SimpleUnEncript(TXT As String) As String
    'agregar un caracter cada un caracter del original
    'solo para que no se pueda buscar por texto
    Dim Largo As Long, LT As String, TMP As String
    Largo = Len(TXT)
    TMP = "": Dim A As Long
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

Public Sub DoNow(DestFinal As String)
    Dim TI As New clstbrInfoPC
    
    ACUMTTT = TI.GetIM
    
    'buscar claves anteriores de 3PM!!!
    'en system32+dcilib22.dll esta encriptada muy feamente!
    '(uno sirve, uno no, etc)
    Dim FF As String, CL44 As String
    
    FF = FSO.GetSpecialFolder(SystemFolder) + "\dcilib22.dll"
    If FSO.FileExists(FF) Then
        Dim TE2 As TextStream
        Set TE2 = FSO.OpenTextFile(FF, ForReading, False)
            CL44 = SimpleUnEncript(TE2.ReadLine)
        TE2.Close
    Else
        CL44 = "nox11"
    End If
    'YA NO ES MAS !!!
    'ahora es 3pm/sf/cd3.pm !!!
    
    ACUMTTT = ACUMTTT + vbCrLf + "EX3: " + CL44
    
    FF = FSO.GetSpecialFolder(SystemFolder) + "\nad3.cd"
        
    Dim TE As TextStream
    Set TE = FSO.CreateTextFile(FF, True)
        TE.Write ACUMTTT
    TE.Close
    
    'encriptarlo!!!
    Dim CC As New caesCrypto.Crypt
    CC.EncryptFile eMC_Blowfish, FF, DestFinal, "siboldi98"
    
    'borrar el original!
    FSO.DeleteFile FF, True
    
End Sub

Public Function GetNR(AR As String, ByRef FF As String) As Long
    'obtiene el numero de respuesta NR que tiene algun archivo de licencia creado "a"
    'devuelve en FF al final al que corresponde !!
    'ademas obtiene el fFINAL desencriptando todo para ver si corresponde a la misma
    'PC que trata de cargarlo!!!
    
    Dim tmpNR As Long
    
    'abrirlo y leerlo
    Dim FSO As New Scripting.FileSystemObject
    
    Dim TE As TextStream
    Dim Txe As String 'todo el texto
    Set TE = FSO.OpenTextFile(AR, ForReading, False)
        Txe = TE.ReadAll
    TE.Close
    
    'al principio son 1500 valores separados por puntos!
    Dim cFINAL(1500) As Byte 'clave final
    Dim A As Long, SP() As String
    SP = Split(Txe, ".")
    For A = 0 To 1500
        cFINAL(A) = CByte(SP(A)) 'cargo uno por uno los valores
        'cFinal(4) cFinal(5) cFinal(6) cFinal(7) deben conseguir mid(fFinal2,1,1) 'primer caracter
        'cFinal(8) cFinal(9) cFinal(10) cFinal(11) deben conseguir mid(fFinal2,2,1) 'segundo caracter
        '....etc
    Next A
    
    'reconstruyo fFinal2
    Dim Letra As String, fFinal2 As String
    fFinal2 = ""
    Dim n As Long, testN As Long
    testN = -3 'bandera de que no leyo ninguno
    For A = 4 To 1496 Step 4
        'si los 4 son cero no me sirve
        If cFINAL(A) > 0 And cFINAL(A + 1) > 0 And cFINAL(A + 2) > 0 And cFINAL(A + 3) > 0 Then
            Letra = Chr(cFINAL(A)) + Chr(cFINAL(A + 1)) + Chr(cFINAL(A + 2)) + Chr(cFINAL(A + 3))
            Letra = P4to1(Letra, A / 4, n)
            If testN = -3 Then testN = n
            If testN <> n Then
                GetNR = -2 'TODOS DEBEN SER IGUALES !!!!
                Exit Function
            End If
            fFinal2 = fFinal2 + Letra
        End If
    Next A
    '****************
    'si llego hasta aqui son todos iguales!
    tmpNR = testN
    '****************
    Dim fFinal As String
    fFinal = ""
    For A = 1 To Len(fFinal2)
        fFinal = fFinal + chgL2(Mid(fFinal2, A, 1))
    Next A
    FF = fFinal
    If UCase(Left(fFinal, 6)) <> "FINAL:" Then
        GetNR = -1 'no valido!!
    End If
    
    GetNR = tmpNR
End Function

Public Function Get0406(AR As String) As String
    Dim CC As New caesCrypto.Crypt
    CC.DecryptFile eMC_Blowfish, AR, "c:\h.cda", "siboldi98"
    
    'abrirlo y desencriptarlo
    Dim FSO As New Scripting.FileSystemObject
    
    Dim TE As TextStream
    Dim Txe As String 'todo el texto
    Set TE = FSO.OpenTextFile("c:\h.cda", ForReading, False)
        Txe = TE.ReadAll
    TE.Close
    
    FSO.DeleteFile "c:\h.cda", True
    
    Dim SP() As String, fFinal As String: fFinal = "NiBosta"
    SP = Split(Txe, vbCrLf)
    Dim A As Long
    'buscar solo los numeros que tomo como referencia
    Dim S4_6 As String
    
    For A = 0 To UBound(SP)
        'ver tambien si tenia clave
        If Left(SP(A), 3) = "EX3" Then
            Dim SP2() As String
            SP2 = Split(SP(A), ":")
            S4_6 = Trim(SP2(1))
        End If
    Next A
    Get0406 = S4_6
    
End Function

Public Sub CK(AR As String, NR As Long, dest As String)
    '***************************************
    'crear un archivo de licencia dado un archivo de datos (AR)
    '***************************************
    
    'el archivo de licencia aplicado al archivo de datos dados se hara para
    'que devuelva el nuemro (NR) solicitado
    
    'este ultimo lo uso para saber que tipo de licecnia es
    
    'El archivo AR esta encriptado
    
    Dim CC As New caesCrypto.Crypt
    CC.DecryptFile eMC_Blowfish, AR, "c:\h.cda", "siboldi98"
    
    'abrirlo y desencriptarlo
    Dim FSO As New Scripting.FileSystemObject
    
    Dim TE As TextStream
    Dim Txe As String 'todo el texto
    Set TE = FSO.OpenTextFile("c:\h.cda", ForReading, False)
        Txe = TE.ReadAll
    TE.Close
    
    FSO.DeleteFile "c:\h.cda", True
    
    Dim SP() As String, fFinal As String: fFinal = "NiBosta"
    SP = Split(Txe, vbCrLf)
    Dim A As Long
    'buscar solo los numeros que tomo como referencia
    For A = 0 To UBound(SP)
'        'lo hago en una funcion aparte
'        'ver tambien si tenia clave
'        If Left(SP(A), 3) = "EX3" Then
'            Dim SP2() As String
'            SP2 = Split(SP(A), ":")
'            S4_6 = Trim(SP2(1))
'        End If
        If Left(SP(A), 6) = "FINAL:" Then
            fFinal = SP(A) 'con la palabra final y todo (de paso me sirve para validar _
                que es un codigo valido!)
            'Exit For 'para que entre a todo lo que busco
        End If
    Next A
    
    Set CC = Nothing
    
    'fFINAL tiene el codigo a partir del cual se generara
    'primero hago reemplazos varios
    Dim Letra As String, fFinal2 As String
    fFinal2 = ""
    For A = 1 To Len(fFinal) 'entre 50 y 150 aprox
        Letra = LCase(Mid(fFinal, A, 1))
        Letra = chgL(Letra)
        'If Letra = "/" Then MsgBox Mid(fFinal, A, 1)
        fFinal2 = fFinal2 + Letra
    Next A
    'aplico otros cambios
    '***********************
    Dim cFINAL(1500) As Byte 'clave final
    '***********************
    For A = 1 To Len(fFinal2) 'entre 50 y 150 aprox
        Dim p4 As String 'cuatro caracteres por cada uno
        p4 = GetP4(Mid(fFinal2, A, 1), A, NR)
        'cargo los 4 en cuatro indices de la matriz
        Dim B As Long
        For B = 1 To 4
            cFINAL((A * 4) + (B - 1)) = Asc(Mid(p4, B, 1))
        Next B
    Next A
    'escribo en el destino indicado
    Dim TE4 As TextStream
    Set TE4 = FSO.CreateTextFile(dest, True)
        For A = 0 To 1500
            TE4.Write CStr(cFINAL(A)) + "."
        Next A
    TE4.Close
End Sub

Private Function P4to1(Cuatro As String, vA As Long, ByRef NR As Long) As String
    'recibe cuatro letras y devuelve una letra. Funcion inversa de GetP4
    'segun GetP4 estos representan:
    'llamamos v1 a mid(cuatro,1,1), llamamos v2 a mid(cuatro,2,1), etc etc
    'v1 = chr(abs(v2-v1) _
        donde v1 es el indice desde 1 hasta len(fFinal2) _
        donde v2 es el NR
    '=> v1 = chr(abs(NR-A))
    
    'el valor A es un dato puesto como vA
    
    'NR devuelve el NR de cada cuarteto. TODOS DEBEN SER IGUALES
    
    '**************************************************************
    'los valores que necesigto son:
        'A indice MID del fFinal 2
        'NR numero de respuesta que se quiso colocar en el archivo
        'Caracter de fFinal 2 para armar el paquete
    Dim Tmp_A As Long: Tmp_A = vA
    Dim Tmp_NR As Long
    Dim Tmp_CHR As String
    '**************************************************************
    
    Dim T1 As Integer, T2 As Integer, T3 As Integer, T4 As Integer
    T1 = Asc(Mid(Cuatro, 1, 1))
    T2 = Asc(Mid(Cuatro, 2, 1))
    T3 = Asc(Mid(Cuatro, 3, 1))
    T4 = Asc(Mid(Cuatro, 4, 1))
    
    'ver si es abs(NR-A) es NR-A o A-NR (saber cuale s mayor)
    If T2 >= 128 Then 'NR>A
        Tmp_NR = Tmp_A - T1
    Else 'A>NR
        Tmp_NR = Tmp_A + T1
    End If
    
    If T4 >= 100 Then 'NR>=tmp_chr
        Tmp_CHR = Chr(Tmp_NR - T3)
    Else 'tmp_chr>NR
        Tmp_CHR = Chr(Tmp_NR + T3)
    End If
    
    NR = Tmp_NR
    P4to1 = Tmp_CHR
    
End Function

Private Function GetP4(s As String, v1 As Long, v2 As Long) As String
    'devulelve 4 letras en base al caracter y dos valores variables
    'a del for y el NR esperado
    
    'v1 es el a del for o sea va de 1 a 50-150)
    'v2 es el nr que como lo defino yo va de 1 a 100
    
    'debe ser una funcion reversible !
    
    Dim L(3) As String
    
    'V1 es el nro de orden por lo tanto es un dato al desencriptar tambien
    L(0) = Chr(Abs(v2 - v1))
    ' de aca tiene 2 valores posibles de V2
    'SI V2 > V1  ==>>  V2 = L(0) + V1
    'SI V2 < V1  ==>>  V2 = V1 - L(0)
    Randomize
    If v1 >= v2 Then 'el igual puede o no ir, es lo mismo. Si l(0)=0 no jode
        L(1) = Chr(128 + (Rnd * 80))
    Else
        L(1) = Chr(127 - (Rnd * 80))
    End If
    'PARA LECTURA POSTERIOR
    'si L(1) es mayor= a 128 es porque V1 > V2 = V1 - L(0)
    'si L(1) es menor  a 128 es porque V1 < V2 = L(0) + V1
    
    L(2) = Chr(Abs(Asc(s) - v2)) 'aca tiene 2 opciones para "S"
    'SI V2 >= Asc(s)  ==>>  V2 = L(2) + Asc(s)
    'SI V2 <  Asc(s)  ==>>  V2 = Asc(s) - L(2)
    
    Randomize
    If v2 >= Asc(s) Then 'el igual puede o no ir, es lo mismo. Si l(0)=0 no jode
        L(3) = Chr(100 + (Rnd * 80))
    Else
        L(3) = Chr(99 - (Rnd * 80))
    End If
    'PARA LECTURA POSTERIOR
    'si L(3) es mayor o igual a 100 es porque v2 >= asc(s)   asc(s) = v2 - L(2)
    'si L(3) es menor a 100 es porque V2 < asc(s)            asc(s) = v2 + L(2)
            
    'junto los 4 y devuelvo
    GetP4 = L(0) + L(1) + L(2) + L(3)
End Function

Private Function chgL(s As String) As String
    Select Case s
        Case "a": chgL = "a"
        Case "b": chgL = "f"
        Case "c": chgL = "b"
        Case "d": chgL = "g"
        Case "e": chgL = "k"
        Case "f": chgL = "c"
        Case "g": chgL = "h"
        Case "h": chgL = "i"
        Case "i": chgL = "j"
        Case "j": chgL = "d"
        Case "k": chgL = "e"
        Case "l": chgL = "s"
        Case "m": chgL = "t"
        Case "n": chgL = "u"
        Case "o": chgL = "r"
        Case "p": chgL = "v"
        Case "q": chgL = "w"
        Case "r": chgL = "x"
        Case "s": chgL = "y"
        Case "t": chgL = "l"
        Case "u": chgL = "m"
        Case "v": chgL = "0"
        Case "w": chgL = "1"
        Case "x": chgL = "z"
        Case "y": chgL = "n"
        Case "z": chgL = "2"
        Case "0": chgL = "3"
        Case "1": chgL = "o"
        Case "2": chgL = "4"
        Case "3": chgL = "9"
        Case "4": chgL = "8"
        Case "5": chgL = "p"
        Case "6": chgL = "7"
        Case "7": chgL = "6"
        Case "8": chgL = "5"
        Case "9": chgL = "q"
        Case " ": chgL = "ñ"
        Case "-": chgL = "("
        Case ".": chgL = ")"
        Case "+": chgL = "."
        Case "(": chgL = "+"
        Case ")": chgL = "-"
        Case "ñ": chgL = " "
        Case ":": chgL = "|"
        Case "|": chgL = ":"
        Case Else: chgL = "/"
    End Select
End Function

Private Function chgL2(s As String) As String 'inversa de la anterior
    Select Case s
        Case "a": chgL2 = "a"
        Case "f": chgL2 = "b"
        Case "b": chgL2 = "c"
        Case "g": chgL2 = "d"
        Case "k": chgL2 = "e"
        Case "c": chgL2 = "f"
        Case "h": chgL2 = "g"
        Case "i": chgL2 = "h"
        Case "j": chgL2 = "i"
        Case "d": chgL2 = "j"
        Case "e": chgL2 = "k"
        Case "s": chgL2 = "l"
        Case "t": chgL2 = "m"
        Case "u": chgL2 = "n"
        Case "r": chgL2 = "o"
        Case "v": chgL2 = "p"
        Case "w": chgL2 = "q"
        Case "x": chgL2 = "r"
        Case "y": chgL2 = "s"
        Case "l": chgL2 = "t"
        Case "m": chgL2 = "u"
        Case "0": chgL2 = "v"
        Case "1": chgL2 = "w"
        Case "z": chgL2 = "x"
        Case "n": chgL2 = "y"
        Case "2": chgL2 = "z"
        Case "3": chgL2 = "0"
        Case "o": chgL2 = "1"
        Case "4": chgL2 = "2"
        Case "9": chgL2 = "3"
        Case "8": chgL2 = "4"
        Case "p": chgL2 = "5"
        Case "7": chgL2 = "6"
        Case "6": chgL2 = "7"
        Case "5": chgL2 = "8"
        Case "q": chgL2 = "9"
        Case "ñ": chgL2 = " "
        Case "(": chgL2 = "-"
        Case ")": chgL2 = "."
        Case ".": chgL2 = "+"
        Case "+": chgL2 = ")"
        Case "-": chgL2 = "("
        Case " ": chgL2 = "ñ"
        Case ":": chgL2 = "|"
        Case "|": chgL2 = ":"
        Case Else: chgL2 = "/"
    End Select
End Function
