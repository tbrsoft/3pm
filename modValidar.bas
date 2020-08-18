Attribute VB_Name = "modValidar"
'el codigo que se pide esta grabado en SYSFOLDER+"\CodPed.cfg"
Private txtClaves As String

Public Sub CrearNuevoCodigoValidar()
    'cuando carga OK una clave se genera el proximo codigo a pedir
    'esto debe ser al azar si son siempre la misma serie de codigos un
    'tipo con varis máquinas solo pide las clavs una vez
    Dim A As Long
    Randomize Timer
    A = Int(Rnd * 1000000) '1 millon
    EscribirArch1Linea SYSfolder + "codped.cfg", CStr(A)
End Sub


Public Function CodigoParaClaveActual() As String
    'codigo que se entrega al dueño para cargar la clave
    'no debe cambiar para que tenga tiempo de pedirla
    
    'PARA QUE NO CAMBIE ESTA EN UNA ARCHIVO
    Dim Cod As String
    Cod = LeerArch1Linea(SYSfolder + "codped.cfg")
    
    CodigoParaClaveActual = Cod
    
End Function

Public Function ClaveParaValidar(CodigoSolicitado As String) ' es siempre un numero del 1 al 1 millon
    txtClaves = "hfhefy487rgh8dX88734grefvberhg8y3487yfeg8hjYJWEGFWUGWVUXHWHJE3FVB8CM8XJD8FYHWUXUX"
    txtClaves = txtClaves + "HGHFGDFGGyqye76462763674838e8dghfvvGW6263GFBVXVCZFADQRgay2738rXgopgj"
    txtClaves = txtClaves + "dhdbdgGQ6327498R09FD83YDBHGDH3626EHDB8Y71637r9d0dujcbdte535detr8wfw68"
    txtClaves = txtClaves + "GDGFCYYW62739FRJXBy17387dhdh38438759fhdbxvzbnxmxcm8hjwu274eg8g727376D"
    txtClaves = txtClaves + "6374HFDBDGHg173849rufcbsvzgytwy37egcy364yfgdvxbv8uy3743yt8g7273HFHXBVV"
    txtClaves = txtClaves + "CXCBNCNXMGDGD652163ydbxnzj8uXe726475859d8098ghdgYEWY3YDHDYW273HRHJDJ"
    txtClaves = txtClaves + "SHCGBEY26153748772737degdb2u74374gdbv8y273ytdgdg32736degdv2636regfg8g26"
    txtClave8 = txtClaves + "cvxcg8ywuywu8jau1837437772737EHDGDH3HDH8JHJQXQOOP8K3MVACg8agd8y6wehd"
    txtClaves = txtClaves + "vcvxbnzmajwX3837dytdg8twq62763gd8gh8jaka3aowqu8bnahGQ5265373828283737646"
    txtClaves = txtClaves + "23098732097094327569263576573432643556325643275684327568326554363284534"
    txtClaves = txtClaves + "1907309874986214987698732648648632432487567465032465065508716508DF8D8D"
    txtClaves = txtClaves + "vxbcnvshdfuwetruywtuXdsygfhj8dfahjfahjfdwquXyetfaghjdfghj8adfghj3476586358fe87a"
    txtClaves = txtClaves + "8djhgXutr726354765w876ftu8dahgfhjd8agf13764532187645876w5f76dtfuXt276547621"
    txtClaves = txtClaves + "297346987qwdfthjd8agfchjd8GA7X2165R9762FTDGUYWGFDHJAg8f762refuytfdc376r23"
    txtClaves = txtClaves + "8cnbv8adfhgd8ahjgdvuyqtruywetuyuXqXqXXoqwXuoX8afdhjk8adgfka7a36816448787tfyaaa"
    txtClaves = txtClaves + "wer8769876trywetgfuhd8agf87t4r9t2f76wgfuywgfeug92167gf916gf9763gf976ewgfuy"
    txtClaves = txtClaves + "weuyrwXuerghhjkgd88768752354WE87TD8GT8DUFHG218RT987TGFUY8GDF7632TR732"
    txtClaves = txtClaves + "DWFGD8F7326T76ATD76T76T76t76t76rR5DF865d65de6D865d865D8865D865DuytdUY"
    txtClaves = txtClaves + "vdcbvxbvcxghjfdUYDUye65RE65eufdJHGHJFDhjgfXTrXfrKFGHJfghjFHGFGHJfghfghFDHJTR"
    txtClaves = txtClaves + "ghfghjkfaxghjdf8ad76UYTRY486548654765465E65EDUYTedytrdeuyrdytrDUYTRDUYTdutr"
    txtClaves = txtClaves + "rm3vf200328177891k3j8fh8aoXduyuXUXYTUYXTRUTFYTERTEUTRDUYfXruyfDYTRDErdXduyrDX"
    txtClaves = txtClaves + "276543321759832745983276987659875217645321764321493246543253245324656666"
    txtClaves = txtClaves + "190730987498ghjg76579873264864863gt524875gt56503246gt55508716508DF8D8D"
    txtClaves = txtClaves + "weuyrwXuerghhgt5d887687523gt5E87TD8GT8DUgt5218RT987Tgt5Y8GDF7632TR732"
    
    Dim Cod As Long
    Cod = CLng(CodigoSolicitado)
    'la clave debe tener tambien algo personalizado para que cada licenciatario tenga sus propias claves
    'la forma mejor es utilizar su clave de administrador. Esta se repite para todas las licencias de un mismo tipo.
    Cod = Cod + SumaCHRtxt(ClaveAdmin)
    Dim Largo As Long
    Largo = Len(txtClaves)
    A = Cod \ Largo 'resultado entero
    Dim Resto As Long
    Resto = Cod - (A * Largo)
    If Resto < 0 Then Resto = -Resto
    'asegurarme que este dentro del txclaves
    Do While Resto > Largo - 20
        Resto = Resto - 20
    Loop
    
    'una clave con 4 valores posibles debe ser de 15 caracteres
    'para que haya 1000 millones de claves posibles
    
    'si coloco una clave con texto el que alquila debera tener un teclado!!!!
    'no debe funcionar así!!!!!
    
    'transformo entonces la clave en algo que el tipo pueda escribir
    'no puedo usar ni el enter ni el escape!
    'me quedan cuatro
    'IZQ DER PagAd y PagAt
    Dim ClaveIntermedia As String
    ClaveIntermedia = Mid(txtClaves, Resto, 15)
    ClaveParaValidar = PasarStrToClave4Teclas(ClaveIntermedia)
End Function

Public Sub RegistroDiario()
    'registra cada inicio de 3PM y el numero que indica el contador
    Dim TE As TextStream
    Set TE = FSO.OpenTextFile(SYSfolder + "daily.cfg", ForAppending, True)
    SumarContadorCreditos 0 'me aseguro que se carge la variable contador
    TE.WriteLine CStr(Date) + " - " + CStr(time) + " Contador R en: " + CStr(CONTADOR) + " Contador H en: " + CStr(CONTADOR2)
    TE.Close

End Sub

Public Function PasarStrToClave4Teclas(TXT As String)
    'se para una cadena (clave) y se transforma en una clave con las 4 teclas de desplazamiento
    Dim ButonIzq As String
    Dim ButonDer As String
    Dim ButonPagAd As String
    Dim ButonPagAt As String
    ButonIzq = Chr(TeclaIZQ)
    ButonDer = Chr(TeclaDER)
    ButonPagAd = Chr(TeclaPagAd)
    ButonPagAt = Chr(TeclaPagAt)
    
    Dim Largo As Long, CC As Long
    Largo = Len(TXT)
    CC = 1
    Dim Letra As String, LetraNew As String, ClaveNew As String
    ClaveNew = ""
    Dim nLetraNew As Long
    Do While CC <= Largo
        Letra = Mid(TXT, CC, 1)
        ' a cada letra le divido su ASC por cuatro y segun el resto le asigno cada una de las 4 posibilidades
        nLetraNew = Asc(Letra) - (4 * (Asc(Letra) \ 4))
        Select Case nLetraNew
            Case 0: LetraNew = ButonIzq
            Case 1: LetraNew = ButonDer
            Case 2: LetraNew = ButonPagAd
            Case 3: LetraNew = ButonPagAt
            'no debe pasar NUNCA!!!
            Case Else: LetraNew = ButonPagAt
        End Select
        ClaveNew = ClaveNew + LetraNew
        CC = CC + 1
    Loop
    PasarStrToClave4Teclas = ClaveNew
    
End Function





