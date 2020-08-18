Attribute VB_Name = "modValidar"

'el codigo que se pide esta grabado en BasePath + "cpd.dor"

Private txtClaves As String

Public Sub CrearNuevoCodigoValidar()
    'tiene que ser de la lista de codigos que se pueden pedir!!!!
    'si no fuera asi no encontraría nunca el pedido para responder
    
    Dim TX As String
    Dim TE As TextStream
    
    Set TE = fso.OpenTextFile(GPF("dalivmp2")) 'archivo encriptado con las claves
        TX = TE.ReadAll
    TE.Close
    
    Dim pos As Long 'posicion del archivo que voy leyendo
    pos = 1
    pos = pos + 16
    'los siguentes 2 digitos especifican el largo del texto
    Dim LN As Long, LN2 As Long, LN3 As Long 'temporales
    TP = Mid(TX, pos, 2)
    LN = CLng(TP)
    pos = pos + 2
    pos = pos + (LN * 4)
    'listo ahora solo los numeros. cada 16 hay 2 grupos de 8 encriptados
    
    Dim AZAR As Long, Limite As Long
    
    Randomize
    Limite = CLng(Rnd * 200)
    
    Dim Pedir As String
    
    For LN = (pos - 1) To (Len(TX) - 16) Step 16
        
        TP = Mid(TX, LN + 2, 1) + Mid(TX, LN + 15, 1) + Mid(TX, LN + 11, 1) + Mid(TX, LN + 16, 1) + _
             Mid(TX, LN + 6, 1) + Mid(TX, LN + 4, 1) + Mid(TX, LN + 14, 1) + Mid(TX, LN + 10, 1)
        
        Pedir = TP 'por si llega al final de todo y no eligio nada
        
        TP2 = Mid(TX, LN + 3, 1) + Mid(TX, LN + 5, 1) + Mid(TX, LN + 9, 1) + Mid(TX, LN + 1, 1) + _
             Mid(TX, LN + 7, 1) + Mid(TX, LN + 8, 1) + Mid(TX, LN + 13, 1) + Mid(TX, LN + 12, 1)

        Randomize
        AZAR = AZAR + CLng(Rnd * 8)
        If AZAR > Limite Then
            Pedir = TP
            Exit For
        End If
    Next LN
    
    EscribirArch1Linea GPF("clavevalid"), Pedir
End Sub


Public Function CodigoParaClaveActual() As String
    'codigo que se entrega al dueño para cargar la clave
    'no debe cambiar para que tenga tiempo de pedirla
    
    'PARA QUE NO CAMBIE ESTA EN UNA ARCHIVO
    Dim Cod As String
    Cod = LeerArch1Linea(GPF("clavevalid"))
    
    CodigoParaClaveActual = Cod
    
End Function

Public Sub RegistroDiario()
    'registra cada inicio de 3PM y el numero que indica el contador
    Dim TE As TextStream
    Set TE = fso.OpenTextFile(GPF("rdcday"), ForAppending, True)
        SumarContadorCreditos 0 'me aseguro que se carge la variable contador
        SumarContadorCarrito 0
        TR.SetVars CONTADOR, CONTADOR2, Contador_Cart, CONTADOR2_Cart
        
        TE.WriteLine CStr(Date) + " - " + CStr(time) + _
            TR.Trad(" Contador R en: %01% Contador H en: %02%%98%Contador " + _
                "R es el contador Reiniciable y contador H es el contador" + _
                "Histórico de monedas insertadas%99%")
        TE.WriteLine TR.Trad(" Contador R(carrito) en: %03% Contador H en: %04%%98%Contador " + _
                "R es el contador Reiniciable y contador H es el contador" + _
                "Histórico de creditos usados en el carrito de compras%99%")
    TE.Close
    
    If FileLen(GPF("rdcday")) > 50000 Then
        'si es muy grande achicarlo.
        Dim TE431 As String
        Set TE = fso.OpenTextFile(GPF("rdcday"), ForReading, True)
            TE431 = TE.ReadAll
        TE.Close
        'le saco al mitad
        Dim MIT As Long
        MIT = Len(TE431) / 2
        TE431 = Right(TE431, MIT)
        Set TE = fso.CreateTextFile(GPF("rdcday"), True)
            TE.WriteLine TE431
        TE.Close
    End If

End Sub

Public Function PasarStrToClave4Teclas(TXT As String)
    'se para una cadena (clave) y se transforma en una clave con las 4 teclas de desplazamiento
    Dim ButonIzq As String
    Dim ButonDer As String
    
    ButonIzq = Chr(TeclaIZQ)
    ButonDer = Chr(TeclaDER)
    
    Dim Largo As Long, CC As Long
    Largo = Len(TXT)
    CC = 1
    Dim Letra As String, LetraNew As String, ClaveNew As String
    ClaveNew = ""
    Dim nLetraNew As Long
    Do While CC <= Largo
        Letra = Mid(TXT, CC, 1)
        ' a cada letra le divido su ASC por cuatro y _
            segun el resto le asigno cada una de las 4 posibilidades
        nLetraNew = Asc(Letra) - (4 * (Asc(Letra) \ 4))
        Select Case nLetraNew
            Case 0: LetraNew = ButonIzq
            Case 1: LetraNew = ButonDer
            Case 2: LetraNew = ButonIzq
            Case 3: LetraNew = ButonDer
            'no debe pasar NUNCA!!!
            Case Else: LetraNew = ButonDer
        End Select
        ClaveNew = ClaveNew + LetraNew
        CC = CC + 1
    Loop
    PasarStrToClave4Teclas = ClaveNew
    
End Function

Public Function NumToTec(nS As String) As String
    
    Dim LL As Long
    LL = Len(nS)
    Dim J As Long, res As String, Letra As String
    res = ""
    For J = 1 To LL
        Letra = Mid(nS, J, 1)
        Select Case Letra
            Case "0": Letra = "IZQ"
            Case "1": Letra = "DER"
            Case "2": Letra = "DER"
            Case "3": Letra = "IZQ"
            Case "4": Letra = "DER"
            Case "5": Letra = "DER"
            Case "6": Letra = "IZQ"
            Case "7": Letra = "IZQ"
            Case "8": Letra = "IZQ"
            Case "9": Letra = "DER"
        End Select
        
        res = res + Letra
        If J < LL Then res = res + " - "
        
    Next J
    
    NumToTec = res
End Function

Public Function TexToTec(nS As String) As String
    
    Dim LL As Long
    LL = Len(nS)
    Dim J As Long, res As String, Letra As String
    res = ""
    For J = 1 To LL
        Letra = Mid(nS, J, 1)
        Select Case Asc(Letra)
            Case TeclaIZQ: Letra = "IZQ"
            Case TeclaDER: Letra = "DER"
            Case Else: Letra = "NO"
        End Select
        
        res = res + Letra
        If J < LL Then res = res + " - "
        
    Next J
    
    TexToTec = res
End Function

Public Function ClaveParaValidar(CodigoSolicitado As String, _
                                 Optional ByRef Usos As Long, _
                                 Optional ByRef PreAviso As Long, _
                                 Optional ByRef RecPC As String) As String

    ClaveParaValidar = 0

    Dim TX As String
    Dim TE As TextStream
    
    Set TE = fso.OpenTextFile(GPF("dalivmp2")) 'archivo encriptado con las claves
        TX = TE.ReadAll
    TE.Close
    
    Dim pos As Long 'posicion del archivo que voy leyendo
    pos = 1
    'las primeras 16 son dos numeros de 8 mezclados que informan de cuantos creditos
    'se valida y con que preaviso
    Dim TP As String, TP2 As String, TP3 As String 'temporales
    
    TP = Mid(TX, pos, 16)
    TP2 = Mid(TP, 3, 1) + Mid(TP, 15, 1) + Mid(TP, 9, 1) + Mid(TP, 1, 1) + _
          Mid(TP, 13, 1) + Mid(TP, 7, 1) + Mid(TP, 11, 1) + Mid(TP, 5, 1)
          
    Usos = CLng(TP2 / 8)
          
    TP2 = Mid(TP, 4, 1) + Mid(TP, 6, 1) + Mid(TP, 8, 1) + Mid(TP, 12, 1) + _
          Mid(TP, 10, 1) + Mid(TP, 14, 1) + Mid(TP, 16, 1) + Mid(TP, 2, 1)
          
    PreAviso = CLng(TP2 / 6)
    
    pos = pos + 16
    
    'los siguentes 2 digitos especifican el largo del texto
    
    Dim LN As Long, LN2 As Long, LN3 As Long 'temporales
    
    TP = Mid(TX, pos, 2)
    LN = CLng(TP)
    pos = pos + 2
    For LN2 = 0 To LN - 1 'cuatro numeros cada letra
        LN3 = CLng(Mid(TX, pos + (LN2 * 4), 4)) / (LN2 + 1)
        TP2 = Chr(LN3)
        RecPC = RecPC + TP2
    Next LN2
      
    pos = pos + (LN * 4)
    'listo ahora solo los numeros. cada 16 hay 2 grupos de 8 encriptados
    
    For LN = pos - 1 To (Len(TX) - 16) Step 16
        TP = Mid(TX, LN + 2, 1) + Mid(TX, LN + 15, 1) + Mid(TX, LN + 11, 1) + Mid(TX, LN + 16, 1) + _
             Mid(TX, LN + 6, 1) + Mid(TX, LN + 4, 1) + Mid(TX, LN + 14, 1) + Mid(TX, LN + 10, 1)
        TP2 = Mid(TX, LN + 3, 1) + Mid(TX, LN + 5, 1) + Mid(TX, LN + 9, 1) + Mid(TX, LN + 1, 1) + _
             Mid(TX, LN + 7, 1) + Mid(TX, LN + 8, 1) + Mid(TX, LN + 13, 1) + Mid(TX, LN + 12, 1)
            
        TP3 = NumToTec(TP2)
        
        If CLng(CodigoSolicitado) = CLng(TP) Then
            'este es el que pidio!"
            ClaveParaValidar = TP2 'devuelvo string para que no olvide los ceros de _
                                    adelante que si cuentan!
            Exit For
        End If

    Next LN
    
End Function


