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
Private fso As New Scripting.FileSystemObject
Private nFOt As New tbrDATA.clsTODO

'tipos de licencia que se pueden cargar
Public Enum TypeLic
    ArchivoNoValidoComoLic = -3
    ParaOtraPC = -2 'no fue hecha para este equipo
    a_SinLeer = -1
    aSinCargar = 0 'todavia no puso nada
    BErronea = 1 'cuando ya erro 3 o mas veces
    CGratuita = 2 'primera que me puede pedir
    DMinima = 3
    EComun = 4
    FPremium = 5
    GFull = 6
    Supsabseee = 7
End Enum

Private mLicencia As TypeLic 'basica de 3pm
'cada plugin tendra tambien su licencia independiente y su archivo de licencia
'ademas cada uno podrá ser licencia superlicencia ext ext para que cada plugin tenga sus niveles


Private mLicencia3PMVtaMusica As TypeLic
Private mLicencia3PMOrigMusicaFTP As TypeLic
Private mLicencia3PMConfigOnline As TypeLic

Private mLicenciaCD001Kar As TypeLic
Private mLicenciaCD002Kar As TypeLic
Private mLicenciaCD003Kar As TypeLic
Private mLicenciaCD004Kar As TypeLic
Private mLicenciaCD005Kar As TypeLic
Private mLicenciaCD006Kar As TypeLic

Private Sub Class_Initialize()
    'GPF("cd3pm") 'DatosClaveIngresada.DLL
    'asegurarme que exista "dciLib22.dll"
    mLicencia = a_SinLeer
    
    mLicencia3PMVtaMusica = a_SinLeer
    mLicencia3PMOrigMusicaFTP = a_SinLeer
    mLicencia3PMConfigOnline = a_SinLeer
    
    mLicenciaCD001Kar = a_SinLeer
    mLicenciaCD002Kar = a_SinLeer
    mLicenciaCD003Kar = a_SinLeer
    mLicenciaCD004Kar = a_SinLeer
    mLicenciaCD005Kar = a_SinLeer
    mLicenciaCD006Kar = a_SinLeer
    
    'generar el archivo de esta PC !
    nFOt.SetSF "3pm" 'nuevo agosto 2007 para no mezclar con karaokes ni con programas de artime y manu
    nFOt.SetLog AP + "kc.log"
    nFOt.DoNow GPF("cd4pm")
    
    'validarlo con la clave si hubiera !
    tERR.Anotar "IC11"
    sabseee "3pm"
    'que ademas se revise todo
    sabseee "mLicencia3PMVtaMusica"
    sabseee "mLicencia3PMOrigMusicaFTP"
    sabseee "mLicencia3PMConfigOnline"
    sabseee "mLicenciaCD001Kar"
    sabseee "mLicenciaCD002Kar"
    sabseee "mLicenciaCD003Kar"
    sabseee "mLicenciaCD004Kar"
    sabseee "mLicenciaCD005Kar"
    sabseee "mLicenciaCD006Kar"
    tERR.Anotar "IC12ff"
    
End Sub

Public Sub SetFileLic(pARCH As String)
    'recibo el archivo de respuesta
    'ponerlo en la ubicacion deseada!
    fso.CopyFile pARCH, GPF("cd7pm")
    'le hago copia de seguridad
    fso.CopyFile GPF("cd7pm"), GPF("cd8pm")
End Sub

Public Function sabseee_STR() As String
    'lee lo grabado para que sea mas rapido,
    'si es la primera vez leo todo
    If mLicencia = a_SinLeer Then
        tERR.Anotar "IC09g22"
        IngresaClave "3pm", False
    End If
    
    If mLicencia = a_SinLeer Then sabseee_STR = TR.Trad("Sin licencia%98%No tiene archivo de licencia de soft %99%")
    If mLicencia = ArchivoNoValidoComoLic Then sabseee_STR = TR.Trad("Archivo no valido%98%El archivo de licencia usado no es un archivo de licencia%99%")
    If mLicencia = aSinCargar Then sabseee_STR = TR.Trad("Sin cargar%98%El archivo de licencia aún no se insertó%99%")
    If mLicencia = BErronea Then sabseee_STR = TR.Trad("Licencia erronea%98%El archivo de licencia es erróneo%99%")
    If mLicencia = CGratuita Then sabseee_STR = TR.Trad("Gratuita%98%Tiene un archivo de licencia gratuito%99%")
    If mLicencia = DMinima Then sabseee_STR = TR.Trad("Licencia mínima%98%No se usa aún pero es un nivel basico de habilitacion para uso del soft%99%")
    If mLicencia = EComun Then sabseee_STR = TR.Trad("Licencia Comun%99%")
    If mLicencia = FPremium Then sabseee_STR = TR.Trad("Licencia Premium%99%")
    If mLicencia = GFull Then sabseee_STR = TR.Trad("Licencia Full%99%")
    If mLicencia = Supsabseee Then sabseee_STR = TR.Trad("Super Licencia%99%")
    If mLicencia = ParaOtraPC Then sabseee_STR = TR.Trad("Licencia para otra PC%98%El archivo es de licencia pero fue diseñado para otra PC%99%")
    
End Function

Public Function sabseee(ParaQue As String) As TypeLic
    'lee lo grabado para que sea mas rapido,
    'si es la primera vez leo todo

    'dentro de ingresa clave se encarga de validar segun la 2h
    
    Dim tmpLLIC As TypeLic
    
    Select Case LCase(ParaQue)
        Case "3pm":                              tmpLLIC = mLicencia
        Case LCase("mLicencia3PMVtaMusica"):     tmpLLIC = mLicencia3PMVtaMusica
        Case LCase("mLicencia3PMOrigMusicaFTP"): tmpLLIC = mLicencia3PMOrigMusicaFTP
        Case LCase("mLicencia3PMConfigOnline"):  tmpLLIC = mLicencia3PMConfigOnline
        Case LCase("mLicenciaCD001Kar"):         tmpLLIC = mLicenciaCD001Kar
        Case LCase("mLicenciaCD002Kar"):         tmpLLIC = mLicenciaCD002Kar
        Case LCase("mLicenciaCD003Kar"):         tmpLLIC = mLicenciaCD003Kar
        Case LCase("mLicenciaCD004Kar"):         tmpLLIC = mLicenciaCD004Kar
        Case LCase("mLicenciaCD005Kar"):         tmpLLIC = mLicenciaCD005Kar
        Case LCase("mLicenciaCD006Kar"):         tmpLLIC = mLicenciaCD006Kar
    End Select
    
    'reviso a ver si esta
    If tmpLLIC = a_SinLeer Then
        tERR.Anotar "IC09g", ParaQue
        'solo me importa registrar si es lic base
        If LCase(ParaQue) = "3pm" Then
            'primero de todo!
            IngresaClave ParaQue, True
        Else
            IngresaClave ParaQue, False
        End If
    End If
    
    'veo y cargo el resultado
    Select Case LCase(ParaQue)
        Case "3pm":                              sabseee = mLicencia
        Case LCase("mLicencia3PMVtaMusica"):     sabseee = mLicencia3PMVtaMusica
        Case LCase("mLicencia3PMOrigMusicaFTP"): sabseee = mLicencia3PMOrigMusicaFTP
        Case LCase("mLicencia3PMConfigOnline"):  sabseee = mLicencia3PMConfigOnline
        Case LCase("mLicenciaCD001Kar"):         sabseee = mLicenciaCD001Kar
        Case LCase("mLicenciaCD002Kar"):         sabseee = mLicenciaCD002Kar
        Case LCase("mLicenciaCD003Kar"):         sabseee = mLicenciaCD003Kar
        Case LCase("mLicenciaCD004Kar"):         sabseee = mLicenciaCD004Kar
        Case LCase("mLicenciaCD005Kar"):         sabseee = mLicenciaCD005Kar
        Case LCase("mLicenciaCD006Kar"):         sabseee = mLicenciaCD006Kar
    End Select
    
End Function

Private Function GetFileOrBUP(F1 As String, F2 As String) As String
    'entre 2 archivos devuelve el original si existe, si no la copia reemplazando _
        el original o nada si no puede solucionarlo
    
    If fso.FileExists(F1) = False Then
        tERR.Anotar "IC02"
        'a ver si esta la copia!!
        If fso.FileExists(F2) = False Then
            tERR.Anotar "IC03"
            GetFileOrBUP = ""
            Exit Function
        Else 'restaurar al original la copia
            tERR.Anotar "IC04"
            fso.CopyFile F2, F1, True
        End If
    End If
    
    GetFileOrBUP = F1

End Function

Public Sub IngresaClave(ParaQue As String, RegFallo As Boolean)
    'ingresa archivo de licencia y devuelve el tipo esperado
    'para todos las licencias posibles (karaokes y plugins tambien)
    tERR.Anotar "IC01"
    
    Dim TmpAR As String
    
    Select Case LCase(ParaQue)
        Case "3pm"
            '****************LIC 3PM ********************************
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("cd7pm"), GPF("cd8pm"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicencia = garrala(TmpAR, "3PM", RegFallo)
                'NUEVO ene 08 para meter el id en registro de venta de musica
                If mLicencia >= DMinima Then 'para los logs!
                    Dim nS As Long
                    nS = nFOt.GetNR(TmpAR, dwQU_See)
                    dwQU_See = UCase(dwQU_See)
                End If
            Else
                mLicencia = aSinCargar
            End If
        
        Case LCase("mLicencia3PMVtaMusica")
            '****************mLicencia3PMVtaMusica*********************
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("plin1"), GPF("plin2"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicencia3PMVtaMusica = garrala(TmpAR, "mLicencia3PMVtaMusica", RegFallo)
            Else
                mLicencia3PMVtaMusica = aSinCargar
            End If
        
        Case LCase("mLicencia3PMOrigMusicaFTP")
            '****************mLicencia3PMOrigMusicaFTP**************
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("plin3"), GPF("plin4"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicencia3PMOrigMusicaFTP = garrala(TmpAR, "mLicencia3PMOrigMusicaFTP", RegFallo)
            Else
                mLicencia3PMOrigMusicaFTP = aSinCargar
            End If
    
        Case LCase("mLicencia3PMConfigOnline")
            '****************mLicencia3PMConfigOnline**************
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("plin5"), GPF("plin6"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicencia3PMConfigOnline = garrala(TmpAR, "mLicencia3PMConfigOnline", RegFallo)
            Else
                mLicencia3PMConfigOnline = aSinCargar
            End If
    
        Case LCase("mLicenciaCD001Kar")
            '****************mLicenciaCD001Kar
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("plin7"), GPF("plin8"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicenciaCD001Kar = garrala(TmpAR, "mLicenciaCD001Kar", RegFallo)
            Else
                mLicenciaCD001Kar = aSinCargar
            End If
    
        Case LCase("mLicenciaCD002Kar")
            '****************mLicenciaCD002Kar
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("plin9"), GPF("plin10"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicenciaCD002Kar = garrala(TmpAR, "mLicenciaCD002Kar", RegFallo)
            Else
                mLicenciaCD002Kar = aSinCargar
            End If
    
        Case LCase("mLicenciaCD003Kar")
            '****************mLicenciaCD003Kar
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("plin11"), GPF("plin12"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicenciaCD003Kar = garrala(TmpAR, "mLicenciaCD003Kar", RegFallo)
            Else
                mLicenciaCD003Kar = aSinCargar
            End If
    
        Case LCase("mLicenciaCD004Kar")
            '****************mLicenciaCD004Kar
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("plin13"), GPF("plin14"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicenciaCD004Kar = garrala(TmpAR, "mLicenciaCD004Kar", RegFallo)
            Else
                mLicenciaCD004Kar = aSinCargar
            End If
    
        Case LCase("mLicenciaCD005Kar")
            '****************mLicenciaCD005Kar
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("plin15"), GPF("plin16"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicenciaCD005Kar = garrala(TmpAR, "mLicenciaCD005Kar", RegFallo)
            Else
                mLicenciaCD005Kar = aSinCargar
            End If
    
        Case LCase("mLicenciaCD006Kar")
            '****************mLicenciaCD006Kar
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("plin17"), GPF("plin18"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicenciaCD006Kar = garrala(TmpAR, "mLicenciaCD006Kar", RegFallo)
            Else
                mLicenciaCD006Kar = aSinCargar
            End If
            
    End Select
    
    'darle una oportunidad con la interfase 2-H
    If (NP > 0) And (Wueltas >= 7) Then
        
        dwQU_See = dwQU_See + "**" + CStr(NP) + "**"
        
        Select Case NP 'SEGUN  LOS CLIENTES ESPECIFICAR
            Case 41837215 'cesar pizarro BsAs
                mLicencia = Supsabseee
            Case 45757651 'diego fontana buenos aires
                mLicencia = GFull
            Case 46523511 'sergio soza (fue en un paquete con 3 mas. Esta es la única con karaoke)
                mLicencia = Supsabseee
                mLicenciaCD001Kar = Supsabseee
                mLicenciaCD002Kar = Supsabseee
            Case 47536194 'sergio soza la mando para que le pongamos karaoke
                mLicencia = Supsabseee
                mLicenciaCD001Kar = Supsabseee
                mLicenciaCD002Kar = Supsabseee
            Case 49509525 'guillermo rosario USB
                mLicencia = Supsabseee
                mLicencia3PMVtaMusica = Supsabseee 'PAGO LICENCIA DE CARRITO
            Case 49547493 'Alejandro molina se lo llevo del taller
                mLicencia = Supsabseee
            Case 50072908 'rocchio
                mLicencia = Supsabseee
            Case 54144024 'ale molina
                mLicencia = Supsabseee
            Case 55923640 'ale molina
                mLicencia = Supsabseee
            Case 56634485 'ale molina
                mLicencia = Supsabseee
            Case 55287301 'cesar pizarro
                mLicencia = Supsabseee
            Case 70108830 'oscar brito licencia simple
                mLicencia = GFull
            Case 82453846 'rodolfo martino (justiniano pose / bell ville)
                mLicencia = Supsabseee
                mLicenciaCD001Kar = Supsabseee
                mLicenciaCD002Kar = Supsabseee
                mLicenciaCD003Kar = Supsabseee
                mLicencia3PMVtaMusica = Supsabseee
            Case Else
                'las que no estan especificadas tienen al menos superlicencia _
                    de 3PM si no especifica lo contrario
                mLicencia = Supsabseee
        End Select
    End If
    
End Sub

'la siguiente funcion recibe un archivo de licencia y la indicacion de para que software o plugin
'fue hecho y en base a eso da como resultado el tipo de licencia que corresponde

Private Function garrala(ArchTest As String, sForSF As String, EsErrorSiNoVa As Boolean) As TypeLic

    tERR.Anotar "setML", ArchTest, sForSF

    Dim TEZ As String 'lista "FINAL:" original
    Dim mNR As Long
    mNR = nFOt.GetNR(ArchTest, TEZ)  'el tez queda en ni bosta!!!
    
    Dim ForSF As String 'saber para que soft fue hecho
    ForSF = nFOt.GetSF
    tERR.Anotar "forsf", ForSF
    
    If ForSF <> "" And LCase(ForSF) <> LCase(sForSF) Then
        ' "" van a dar las anteriores supuestamente
        ' las mantengo por compatibilidad.
        mNR = -2333
    End If
    
    tERR.Anotar "IC05", mNR
    'si da -2 es que puso un archivo que NO es de licencia
    
    Dim TmpLic As TypeLic
    'si mnr es <0 no sirve el archivo !!!
    If mNR < 0 Then
        TmpLic = ArchivoNoValidoComoLic
        If EsErrorSiNoVa Then tERR.AppendLog "IC09c"
        GoTo FIM
    End If

    TEZ = Trim(UCase(TEZ))
    tERR.Anotar "IC06", TEZ
    
    Dim FFdeLaClave As String
    FFdeLaClave = nFOt.GetRF 'da ok el de esta misma PC
    
    FFdeLaClave = Trim(UCase(FFdeLaClave))
    tERR.Anotar "IC07", FFdeLaClave
    
'    'lo escribo por las dudas
'    Dim POP4 As TextStream
'    Set POP4 = FSO.OpenTextFile(AP + "sf\fur.ia", ForWriting, True)
'        POP4.WriteLine TEZ
'        POP4.WriteLine FFdeLaClave
'    POP4.Close
    
    If TEZ <> FFdeLaClave Then
    
        'VER SI NO SON TAN DISTINTAS LAS 2 PCs
        Dim DF As String
        'primero esta pc y 2º para la que fue hecha la clave
        DF = nFOt.GetDiff2PC(FFdeLaClave, TEZ)
        
        Dim RET(20) As Long
        Dim L As Long
        For L = 1 To Len(DF)
            RET(L - 1) = CLng(Mid(DF, L, 1))
        Next L
        
        Dim PT1 As Long 'puntos para medir la igualdad de 2 equipos
        'ret(0):si es cero la PC1 tiene el minimo de indices necesarios para arrancar
        'ret(1):si es cero la PC2 tiene el minimo de indices necesarios para arrancar
        
        'ret(2):cantidad de placas de red de la PC1
        'ret(3):cantidad de placas de red de la PC2
        'ret(4):cantidad de coindicencias en placas de red
        
        'ret(5):dif en la bios si es 4 esta ok, si es 6 esta mal. NO puede ser otro resultado
        'ret(6):dif en el micro si es 2 esta ok, si es 1 esta mal. NO puede ser otro resultado
        'ret(7):dif en el micro reserved si es 5 esta ok, si es 3 esta mal. NO puede ser otro resultado
        
        'ret(8):cantidad de discos en la pc1
        'ret(9):cantidad de discos en la pc2
        'ret(10):coincidencias en los discos
        
        '00 111 4 1 3 111 0000000000
        '001921F5682A|00000000|BFEBFBFF00000F49|409000F|WD-WMAM98125093|
        '001921F5682A|00000000|BFEBFBFF00000F64|604000F|WD-WMAM98125093|
        If RET(0) = 1 Then PT1 = 0: GoTo VALIDAR
        If RET(1) = 1 Then PT1 = 0: GoTo VALIDAR
        
        If RET(2) = RET(3) Then PT1 = PT1 + 10
        PT1 = PT1 + (RET(4) * 50)
        
        If RET(5) = 4 Then PT1 = PT1 + 20
        If RET(6) = 2 Then PT1 = PT1 + 20
        If RET(7) = 5 Then PT1 = PT1 + 20
        
        If RET(8) = RET(9) Then PT1 = PT1 + 10
        PT1 = PT1 + (RET(10) * 50)
        
        Dim CM As String 'COMO MOSTRAR
        CM = "DF:" + DF + vbCrLf + _
            "PT:" + CStr(PT1) + vbCrLf
            
        Dim H As Long
        Dim TJ As String 'todo junto
        TJ = "B" + TEZ + "WEX" + FFdeLaClave
        Dim Mx As Long
        Mx = Len(TJ)
        
        For H = 1 To 200
            If (H < 27) Or (H > (Mx + 27)) Then
                Randomize
                CM = CM + CStr(Int(Rnd * 7))
                CM = CM + CStr(Int(Rnd * 7))
            Else
                If (H - 27) <= Len(TJ) Then CM = CM + Mid(TJ, (H - 26), 1)
            End If
        Next H
        'que se esconda bien
        CM = Replace(CM, "|", "A")
        
        If EsErrorSiNoVa Then tERR.AppendSinHist CM
        
VALIDAR:
        If PT1 >= 100 Then
            GoTo ValeTAMBIEN
        Else
            TmpLic = ParaOtraPC
            If EsErrorSiNoVa Then tERR.AppendLog "IC08" 'para que me envien el registro en cualquier caso
        End If
    Else 'el archivo fue hecho para esta PC
    
ValeTAMBIEN:
    
        tERR.Anotar "IC09e", mNR
        Select Case mNR
            Case Is < 0
                TmpLic = ArchivoNoValidoComoLic
                If EsErrorSiNoVa Then tERR.AppendLog "IC09f"
            Case 10, 19, 81:                        TmpLic = CGratuita
            Case 30, 69, 82:                        TmpLic = DMinima
            Case 40, 44, 55, 22, 65:                TmpLic = EComun
            Case 50, 2, 9:                          TmpLic = FPremium 'el 65 estaba acá!!!
            Case 60, 67, 66, 98, 11:                TmpLic = GFull
            Case 70, 71, 72, 3, 5, 7, 8, 9, 42, 79: TmpLic = Supsabseee
            
            Case Else
                TmpLic = BErronea
                If EsErrorSiNoVa Then tERR.AppendLog "IC10"
        End Select
        
    End If
    
FIM:
    '**********************
    garrala = TmpLic
    '**********************
End Function
