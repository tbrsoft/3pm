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
    LicenciaVencida = -4 'mm94
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
    SuperLicencia = 8 'CRACK PROPIO
End Enum

Private mLicencia As TypeLic 'basica de 3pm
'cada plugin tendra tambien su licencia independiente y su archivo de licencia
'ademas cada uno podrá ser licencia superlicencia ext ext para que cada plugin tenga sus niveles

Private mLicencia3PMVtaMusica As TypeLic
Private mLicencia3PMOrigMusicaFTP As TypeLic
Private mLicencia3PMConfigOnline As TypeLic

Private mLicenciaCD001Kar As TypeLic

'**************************************************************
'valores que esta vueno saber en todo momento una vez leido
'todos referidos a validacion por fecha o versión de licencia
Private nnFecha As Long 'fecha del archivo L37
Private nnVers As Long
Private nnVers2 As Long 'version del 3pm que estoy usando
Private nnFecha2 As Long
Private difXaSoporte As Long
Private difXaVersion As Long
'**************************************************************

Private Sub Class_Initialize()
    'GPF("cd3pm") 'DatosClaveIngresada.DLL
    'asegurarme que exista "dciLib22.dll"
    mLicencia = a_SinLeer
    
    mLicencia3PMVtaMusica = a_SinLeer
    mLicencia3PMOrigMusicaFTP = a_SinLeer
    mLicencia3PMConfigOnline = a_SinLeer
    
    mLicenciaCD001Kar = a_SinLeer
    
    'generar el archivo de esta PC !
    nFOt.SetSF dcr("1Vx0YVGhEoIisHPLAZMHXw==") 'nuevo agosto 2007 para no mezclar con karaokes ni con programas de artime y manu
    nFOt.SetLog AP + "kc.log"
    
    'ANTES DE CREAR EL ARCHIVO DEJO INDICADO DE QUE VERSION ES PARA QUE AL CREAR LA LICENCIA USE ESTA VERSION
    Dim nnVers3 As Long
    nnVers3 = App.Revision
    nnVers3 = nnVers3 + CLng(App.Minor) * 1000
    nnVers3 = nnVers3 + CLng(App.Major) * 100000
    
    nFOt.SetSV nnVers3
    nFOt.DoNow GPF("cd4pm")
    
    'validarlo con la clave si hubiera !
    tERR.Anotar "IC11"
    sabseee dcr("q44KmdDBQ+IB8dTOX8F+VA==")
    'que ademas se revise todo
    sabseee dcr("MCuVh38359iRH+GBaAkXedz8Pl38peUqZHKs0a0SpMe+QLrW9mKdnA==")
    sabseee dcr("yTSbeYe2oWp2ydIUpGyes+DYNN6qU8l9pMMGAAqH+wBg8bBgTQ+/hw==")
    sabseee dcr("VZPmSDtgWIj2UthiVZfN1LsFHe7IZv/K/ue9/JPXBYNJosAztaasKg==")
    sabseee dcr("OqgcJfckN8975IVShi0xrqPphoO7CJfy1bRk3zQnHno=")
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
        IngresaClave dcr("1Vx0YVGhEoIisHPLAZMHXw=="), False 'solo para asegurarme que esta cargada
    End If
    
    
    If mLicencia = LicenciaVencida Then
        '"Licencia vencida%98%La licencia esta vencida %99%"
        tSTR = dcr("fcyo88AcnaWKp+49MX/u7PO4s3B0jrkUw81yvait8bYLD/zDgC3aS0QoYo7iG3cGviOg5R7QeFgLLtZ8FQg/MQ==")
    End If
    
    If mLicencia = a_SinLeer Then
        '"Sin licencia%98%No tiene archivo de licencia de soft %99%"
        tSTR = dcr("+VqjbTpH6IUOjP76N1HFR7SaIBWIBDcganv2v3UXKSgcB5TS9VPdG7j7hI8IicTwntDj/KeSUl7xe9VEFEi4B4XRQeuvBk47")
    End If
    
    If mLicencia = ArchivoNoValidoComoLic Then
        '"Archivo no valido%98%El archivo de licencia usado no es un archivo de licencia%99%"
        tSTR = dcr("dWaDd9maDMPgYcpF+NoT2SYcGIKNLO47Mu31QnIzGWP6pE6djm3Rt3Gpl6MCdQPeiF4FeywD3DSvfBi8iH2DT+69Ui3apiOfegk2vTGorLLE3OlSWqGSFYg06EOODJ26")
    End If
    
    If mLicencia = aSinCargar Then
        'Sin cargar%98%El archivo de licencia aún no se insertó%99%
        tSTR = dcr("xLuAlBXuWdN+CL5Xda9TGu7EV+BuQZX6ShZHUpJwsAqYNi9Lau3U8BfGUH7e0K2X2aczY9q0gxXv2oiPhUnrkbp8XeHKvJhd")
    End If
    
    If mLicencia = BErronea Then
        'Licencia erronea%98%El archivo de licencia es erróneo%99%
        tSTR = dcr("F6aTfCLsBgkxzNwYXEkOzcWIuMwkAYBM+Om3yg5qKUX8nisLCZWVVz5U+mOawnicdwPyjRCGCwIFAyMsFLvB1t+ETzRGwxSS")
    End If
    
    If mLicencia = CGratuita Then
        'Gratuita%98%Tiene un archivo de licencia gratuito%99%
        tSTR = dcr("JNdn8MvV0P2OoitQ/U2n9yYZp8LEWCdOV4xVmg644QHtAZppEXD+eOCFS9w67K769Y+gZWHKKmVABPNMC19m5zKt8HthdFlq")
    End If
    
    If mLicencia = DMinima Then
        'Licencia mínima%98%No se usa aún pero es un nivel basico de habilitacion para uso del soft%99%
        tSTR = dcr("GC5cm/kRNvwVlX4LFdTKDYPPGL0arCK4aQqEd6ZlMHOy9u4/jWC0TUJ0rFWeZ8c5hHN3ulg4pHVZJInO9wc40eDaNTDLfAK2PSTGEXwrc5n1nGERnzx9f9eh4VLv6/yv+zpxa/1krnA7l0YzOcLALw==")
    End If
        
    If mLicencia = EComun Then
        'Licencia Comun%99%
        tSTR = dcr("z4kgPOAU95B4Mf8LqjJYH3RftDvk3mgtsIg0EELxM5U=")
    End If
    
    If mLicencia = FPremium Then
        'Licencia Premium%99%
        tSTR = dcr("8RCmPOEgXUkJup/WK24oEHgMZ06Rc/kuCciUSI05Rfw=")
    End If
    
    If mLicencia = GFull Then
        'Licencia Full%99%
        tSTR = dcr("zpY+sIjhiE6K3ztV5KwihvVNHCJmuZqPhldBHVFMyyY=")
    End If
    
    If mLicencia = Supsabseee Then
        'Super Licencia%99%
        tSTR = dcr("4PXXE1RY0qtIW2V9WdMaiP/+Z/Ai7M5a7jFweoENBDg=")
    End If
    
    If mLicencia = SuperLicencia Then
        'crack malditos
        'libertad al 3PM !!!!
        tSTR = dcr("BE1s+L/UQcVFZoCB2K5uK1YXnVpNO+7xC0MVlRoELwVkAUAq+Gby/Q==")
    End If
    
    If mLicencia = ParaOtraPC Then
        'Licencia para otra PC%98%El archivo es de licencia pero fue diseñado para otra PC%99%
        tSTR = dcr("2ODmQxipWg9YZCfTTH+nlLzFaiXAcDDLHdjFKkDGpWKvs4BgOcEtmh/hAU2t9LxRYDWU2l0cTLHxnzT+JMpJeqBqfI6YDKBa/+chvCV+wT/ulJJ3tmuTSwYaM+ZYUBqmhhR1yb2eUz0=")
    End If
    
    '///////////////////////////////////////////////////
    sabseee_STR = TR.Trad(tSTR)
    
End Function

Public Function sabseee(ParaQue As String) As TypeLic
    'lee lo grabado para que sea mas rapido,
    'si es la primera vez leo todo

    'dentro de ingresa clave se encarga de validar segun la 2h
    
    Dim tmpLLIC As TypeLic
    
    Select Case LCase(ParaQue)
        Case dcr("1Vx0YVGhEoIisHPLAZMHXw=="):                                         tmpLLIC = mLicencia
        Case LCase(dcr("MCuVh38359iRH+GBaAkXedz8Pl38peUqZHKs0a0SpMe+QLrW9mKdnA==")):  tmpLLIC = mLicencia3PMVtaMusica
        Case LCase(dcr("yTSbeYe2oWp2ydIUpGyes+DYNN6qU8l9pMMGAAqH+wBg8bBgTQ+/hw==")):  tmpLLIC = mLicencia3PMOrigMusicaFTP
        Case LCase(dcr("VZPmSDtgWIj2UthiVZfN1LsFHe7IZv/K/ue9/JPXBYNJosAztaasKg==")):  tmpLLIC = mLicencia3PMConfigOnline
        Case LCase(dcr("OqgcJfckN8975IVShi0xrqPphoO7CJfy1bRk3zQnHno=")):              tmpLLIC = mLicenciaCD001Kar
    End Select
    
    'reviso a ver si esta
    If tmpLLIC = a_SinLeer Then
        tERR.Anotar "IC09g", ParaQue
        'solo me importa registrar si es lic base
        If LCase(ParaQue) = dcr("q44KmdDBQ+IB8dTOX8F+VA==") Then
            'primero de todo!
            IngresaClave ParaQue, True
        Else
            IngresaClave ParaQue, False
        End If
    End If
    
    'veo y cargo el resultado
    Select Case LCase(ParaQue)
        Case dcr("1Vx0YVGhEoIisHPLAZMHXw=="):                                          sabseee = mLicencia
        Case LCase(dcr("MCuVh38359iRH+GBaAkXedz8Pl38peUqZHKs0a0SpMe+QLrW9mKdnA==")):   sabseee = mLicencia3PMVtaMusica
        Case LCase(dcr("yTSbeYe2oWp2ydIUpGyes+DYNN6qU8l9pMMGAAqH+wBg8bBgTQ+/hw==")):   sabseee = mLicencia3PMOrigMusicaFTP
        Case LCase(dcr("VZPmSDtgWIj2UthiVZfN1LsFHe7IZv/K/ue9/JPXBYNJosAztaasKg==")):   sabseee = mLicencia3PMConfigOnline
        Case LCase(dcr("OqgcJfckN8975IVShi0xrqPphoO7CJfy1bRk3zQnHno=")):               sabseee = mLicenciaCD001Kar
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
        Case dcr("1Vx0YVGhEoIisHPLAZMHXw==")
            '****************LIC 3PM ********************************
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("cd7pm"), GPF("cd8pm"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicencia = garrala(TmpAR, dcr("q44KmdDBQ+IB8dTOX8F+VA=="), RegFallo) 'este es el UNICO lugar donde reviso la licencia directo al archivo
                'NUEVO ene 08 para meter el id en registro de venta de musica
                If mLicencia >= DMinima Then 'para los logs!
                    Dim nS As Long
                    nS = nFOt.GetNR(TmpAR, dwQU_See)
                    dwQU_See = UCase(dwQU_See)
                End If
            Else
                mLicencia = aSinCargar
            End If
        
        Case LCase(dcr("MCuVh38359iRH+GBaAkXedz8Pl38peUqZHKs0a0SpMe+QLrW9mKdnA=="))
            '****************mLicencia3PMVtaMusica*********************
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("plin1"), GPF("plin2"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicencia3PMVtaMusica = garrala(TmpAR, dcr("MCuVh38359iRH+GBaAkXedz8Pl38peUqZHKs0a0SpMe+QLrW9mKdnA=="), RegFallo)
            Else
                mLicencia3PMVtaMusica = aSinCargar
            End If
        
        Case LCase(dcr("yTSbeYe2oWp2ydIUpGyes+DYNN6qU8l9pMMGAAqH+wBg8bBgTQ+/hw=="))
            '****************mLicencia3PMOrigMusicaFTP**************
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("plin3"), GPF("plin4"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicencia3PMOrigMusicaFTP = garrala(TmpAR, dcr("yTSbeYe2oWp2ydIUpGyes+DYNN6qU8l9pMMGAAqH+wBg8bBgTQ+/hw=="), RegFallo)
            Else
                mLicencia3PMOrigMusicaFTP = aSinCargar
            End If
    
        Case LCase(dcr("VZPmSDtgWIj2UthiVZfN1LsFHe7IZv/K/ue9/JPXBYNJosAztaasKg=="))
            '****************mLicencia3PMConfigOnline**************
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("plin5"), GPF("plin6"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicencia3PMConfigOnline = garrala(TmpAR, dcr("VZPmSDtgWIj2UthiVZfN1LsFHe7IZv/K/ue9/JPXBYNJosAztaasKg=="), RegFallo)
            Else
                mLicencia3PMConfigOnline = aSinCargar
            End If
    
        Case LCase(dcr("OqgcJfckN8975IVShi0xrqPphoO7CJfy1bRk3zQnHno="))
            '****************mLicenciaCD001Kar
            'si no cargo el archivo no leer nada
            TmpAR = GetFileOrBUP(GPF("plin7"), GPF("plin8"))
            'hasta aqui solo se define el archivo que se va a probar
            If TmpAR <> "" Then
                mLicenciaCD001Kar = garrala(TmpAR, dcr("OqgcJfckN8975IVShi0xrqPphoO7CJfy1bRk3zQnHno="), RegFallo)
            Else
                mLicenciaCD001Kar = aSinCargar
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
            Case 47536194 'sergio soza la mando para que le pongamos karaoke
                mLicencia = Supsabseee
                mLicenciaCD001Kar = Supsabseee
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
            Case 68838655 'la trajo anzil gustavo para reparar
                mLicencia = Supsabseee
            Case 70108830 'oscar brito licencia simple
                mLicencia = GFull
                
            Case 82453846, 86542785, 85721433, 85620164, 87381243, 89463531 'rodolfo martino (justiniano pose / bell ville)
                mLicencia = Supsabseee
                'mLicenciaCD001Kar = Supsabseee' se arrepintio al final ? compro solo expendedores y un karaoke
                mLicencia3PMVtaMusica = Supsabseee
            
            Case 85888147 'oscar brito licencia simple
                mLicencia = GFull
            Case 90566933 'martino pide especial con karaoke
                mLicencia = Supsabseee
                mLicenciaCD001Kar = Supsabseee
            Case 95641791 ' gruccos 14/11/08
                mLicencia = Supsabseee
            Case 90609013 'dario settimo (antes era solo t-store, ahora pago la diferencia y es con expendedor)
                mLicencia = Supsabseee 'ParaOtraPC
                mLicencia3PMVtaMusica = Supsabseee
            Case 93685209 'roberto abrahan bsas
                mLicencia = Supsabseee
            Case 94939172 ' ariel schell compro interfase para un equipo
                '001BFCD5BB81.101111111111|SYS-1234567890|BFEBFBFF00000F65|605000F|S08EJDSP516514|
                mLicencia = Supsabseee 'se la active por que fallaban varias cosas y me arto
                'no le pongo nada de licencia, que salga del archivo de texto que ya tiene
            Case 93549609 'creo que es del ale veiga
                mLicencia = Supsabseee
                mLicenciaCD001Kar = Supsabseee
            Case 93633532 ' carlos salas compro interfase para un equipo que ya tenia licencia de archivo
                'mLicencia = Supsabseee
                'no le pongo nada de licencia, que salga del archivo de texto que ya tiene
                
                mLicencia = Supsabseee 'la formateo y compro el bingo, vamos a confiar en el !!! (19/11/2009)
                
            Case 93934714 'carlos salas t-Store !
                'nada !
            Case 96104800 'antes era de tStore para chicago pero la cambio por 3pm
                mLicencia = Supsabseee
                
            Case 9651396 'las chelaz bsas Juan Carlos Mendoza
                mLicencia = Supsabseee
                
            Case 96369266 'oscar britto chaco 16/01/2009
                mLicencia = Supsabseee
                
            '15 completas para victor grande de arias
            '********************************************************************************
            Case 102996647, 101296007, 119202121, 97670972, 102271795, 102872252
                'maldita sea !!!! se habia bloqueado y ya habia pagado !!!
                'If Year(Date) > 2008 And Month(Date) < 6 Then
                    mLicencia = Supsabseee
                    mLicenciaCD001Kar = Supsabseee
                    mLicencia3PMVtaMusica = Supsabseee
                'End If
            'Case 97670972 'ahora esta en manos de raul pogonza, estaba registrada como "grande de arias" ???
            'puede ser que luis haya hecho dos con el mismo numero ?????????????
            'si luis estuvo arreglando apurado y saco algunas con numeros repetidos
            
            '15 completas para victor grande de arias
            Case 98369777, 99625825, 98225533, 101871311, 99924147, 99354684, 103290080
                'maldita sea !!!! se habia bloqueado y ya habia pagado !!!
                'If Year(Date) > 2008 And Month(Date) < 8 Then
                    mLicencia = Supsabseee
                    mLicenciaCD001Kar = Supsabseee
                    mLicencia3PMVtaMusica = Supsabseee
                'End If
            
            '15 completas para victor grande de arias
            Case 100677430, 98804175, 100154459
                mLicencia = Supsabseee
                mLicenciaCD001Kar = Supsabseee
                mLicencia3PMVtaMusica = Supsabseee
            '********************************************************************************
                
            Case 104329645, 103389501, 107510805 'martino, solo mpRock-(genesis-MP6) [14 abr 2009]
                mLicencia = ParaOtraPC
            
            Case 90213537 'PRUEBA EXPENDEDOR
                mLicencia = Supsabseee
                mLicenciaCD001Kar = Supsabseee
                mLicencia3PMVtaMusica = Supsabseee
                 
            Case 105929791 'gustavo lopez rosario, full chiche
                mLicencia = Supsabseee
                mLicenciaCD001Kar = Supsabseee
                mLicencia3PMVtaMusica = Supsabseee
            Case 106218576 'ale veiga
                mLicencia = Supsabseee
            
            Case 109995186 'dardo paso de mpRock a 3Pm y la actualizamos aqui a full
                mLicencia = Supsabseee
                mLicenciaCD001Kar = Supsabseee
                mLicencia3PMVtaMusica = Supsabseee
                
            Case 109529197 'ale veiga 1/2 activacion de extras
                mLicencia = Supsabseee
                mLicenciaCD001Kar = Supsabseee
                mLicencia3PMVtaMusica = Supsabseee
                
            Case 110959589 ''dardo paso de mpRock a 3Pm y la actualizamos aqui a full
                mLicencia = Supsabseee
                mLicenciaCD001Kar = Supsabseee
                mLicencia3PMVtaMusica = Supsabseee
                
            Case Else
                'las que no estan especificadas tienen al menos superlicencia _
                    de 3PM si no especifica lo contrario
                mLicencia = Supsabseee
        End Select
        
        tERR.AppendSinHist "TryNp:" + CStr(NP) + ":" + CStr(mLicencia) + ":" + CStr(mLicenciaCD001Kar) + ":" + CStr(mLicencia3PMVtaMusica)
        
    End If
    
End Sub

'la siguiente funcion recibe un archivo de licencia y la indicacion de para que software o plugin
'fue hecho y en base a eso da como resultado el tipo de licencia que corresponde

Private Function garrala(ArchTest As String, sForSF As String, EsErrorSiNoVa As Boolean) As TypeLic

    tERR.Anotar "setML", ArchTest, sForSF

    'version crackeada PROPIA
    If MDCN2 > 0 Then
        garrala = SuperLicencia
        Exit Function
    End If
    
    Dim TEZ As String 'lista "FINAL:" original
    Dim mNR As Long
    mNR = nFOt.GetNR(ArchTest, TEZ)    'el tez queda en ni bosta!!!
    
    Dim ForSF As String 'saber para que soft fue hecho
    ForSF = nFOt.GetSF
    tERR.Anotar "forsf", ForSF
    
    If ForSF <> "" And LCase(ForSF) <> LCase(sForSF) Then
        ' "" van a dar las anteriores supuestamente
        ' las mantengo por compatibilidad.
        mNR = -2333
        tERR.Anotar "IC05a", ForSF, sForSF
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
    
    '///////////////////////////////////////////////////////////////////////
    'mm94 verque la licencia no este vencida (SOLO EN 3PM, las otras licencias no importan por ahora)
    If LCase(sForSF) = dcr("1Vx0YVGhEoIisHPLAZMHXw==") Then
        Dim resVersionAndTime As String
        'set 08 nuevo
        'getNR carga si lo tuviera el L37 la fecha y version de la licencia
        
        nnFecha = nFOt.GetFH
        resVersionAndTime = "FH:" + CStr(nnFecha)
        tERR.Anotar "IC09L", "FH:" + CStr(nnFecha)
        
        'VER QUE VERSION ES la que tiene el archivo de licencia!
        
        nnVers = nFOt.getSV
        resVersionAndTime = resVersionAndTime + vbCrLf + "SV:" + CStr(nnVers)
        tERR.Anotar "IC09m", "SV:" + CStr(nnVers)
        
        
        nnVers2 = App.Revision
        nnVers2 = nnVers2 + CLng(App.Minor) * 1000
        nnVers2 = nnVers2 + CLng(App.Major) * 100000
        resVersionAndTime = resVersionAndTime + vbCrLf + "SV2:" + CStr(nnVers2)
        tERR.Anotar "IC09n", "SV2:" + CStr(nnVers2)
        
        
        nnFecha2 = (Year(Date) - 2000) * 10000 'año 2008 es = 80.000
        nnFecha2 = nnFecha2 + Month(Date) * 100 'setiembre de 2008 seria 80.000 + 900 = 80.900
        nnFecha2 = nnFecha2 + Day(Date) '23 set 08 = 80.923
        resVersionAndTime = resVersionAndTime + vbCrLf + "FH2:" + CStr(nnFecha2)
        tERR.Anotar "IC09o", "FH2:" + CStr(nnFecha2)
        
        'ejemplos: 31/12/08 = 81231
        '          01/01/09 = 90101
        '          01/10/09 = 91001
        '          31/12/09 = 91231
        'en resumen si la diferencia es mayor a 10.000 ya no tiene soporte
        
        'como las versiones tienen fechas de salida debo verificar si la fecha del tipo esta mal!!
        'por ejemplo si el tipo tiene version 7.18.233 y la fecha dice agosto de 2008 esta mal !!! es de diciembre de 2007
        '23/09/2008 = 7.42.000 = 742000
        '12/09/2008 = 7.39.000 = 739000
        '20/08/2008 = 7.22.720 = 722720
        '24/03/2008 = 7.20.010 = 720010
        '09/12/2007 = 7.18.049 = 718049
        '06/12/2007 = 7.16.226 = 716226
        '27/08/2007 = 7.00.155 = 700155
        'mm94
        If (nnVers2 > 700155) And (nnFecha2 < 70827) Then nnFecha = 70827
        If (nnVers2 > 716226) And (nnFecha2 < 71206) Then nnFecha = 71206
        If (nnVers2 > 718049) And (nnFecha2 < 71209) Then nnFecha = 71209
        If (nnVers2 > 720010) And (nnFecha2 < 80324) Then nnFecha = 80324
        If (nnVers2 > 722720) And (nnFecha2 < 80820) Then nnFecha = 80820
        If (nnVers2 > 739000) And (nnFecha2 < 80912) Then nnFecha = 80912
        'la lista hasta aqui no tiene sentido pero es demostrativa ya que esto se crea en la version 7.42
        resVersionAndTime = resVersionAndTime + vbCrLf + "FH3:" + CStr(nnFecha)
        tERR.Anotar "IC09p", "FH3:" + CStr(nnFecha)
        
        
        difXaSoporte = nnFecha2 - nnFecha
        resVersionAndTime = resVersionAndTime + vbCrLf + "DfSupp:" + CStr(difXaSoporte)
        tERR.Anotar "IC09q", "DfSupp:" + CStr(difXaSoporte)
        
        If difXaSoporte < 0 Then
            'el tipo tiene problemas con el reloj o me quiere cagar
            'o cuando se creo el archivo de licencia se creo con fecha del futuro por error
            tERR.AppendSinHist "SUPP0089" + vbCrLf + resVersionAndTime
        End If
        
        If difXaSoporte > 10000 Then
            'ya no tiene soporte'
            tERR.AppendSinHist "SUPP0099" + vbCrLf + resVersionAndTime
        End If
            
        
        'en este mometo la version pública es 7.39 (739000) por lo que no puedo cagar a los anteriores
        'ademas si el archivo de licencia no tienes esto (versiones menores a 7.42) dara solo 6 (0.00.006)
        If nnVers < 739000 Then nnVers = 739000
        difXaVersion = nnVers2 - nnVers '(imposible menor que cero!!)
        resVersionAndTime = resVersionAndTime + vbCrLf + "DfVs:" + CStr(difXaVersion)
        tERR.Anotar "IC09r", "DfVs:" + CStr(difXaVersion)
        
        If difXaVersion > 100000 Then
            'ya no tiene actualziación!!!
            'SU LICENCIA DE 3PM LO CUBRE POR ACTUALIZACIONES HASTA LA VERSION "nnvers"+100.000 y usted esta usando "nnvers2"
            TmpLic = LicenciaVencida
            If EsErrorSiNoVa Then tERR.AppendLog "IC09k"
            GoTo FIM
        End If
        
        'los limites de 3PM son 1 año o 1 punto de version major
        'por ejemplo si la fecha de la licencia es 80923 = 2008 / 09 / 23 y date es 2008 / 09 / 24 ya no se valida,
        'no quiere decir que no tiene licencia sino que no tiene soporte
        'por ejemplo
    End If
    
    '///////////////////////////////////////////////////////////////////////
    TEZ = Trim(UCase(TEZ))
    tERR.Anotar "IC06", TEZ 'final del L37 abierto
    
    FFdeLaClave = nFOt.GetRF 'da ok el de esta misma PC
    
    FFdeLaClave = Trim(UCase(FFdeLaClave))
    tERR.Anotar "IC07", FFdeLaClave 'final de esta PC
    
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
            If EsErrorSiNoVa Then tERR.AppendLog "IC08", CStr(PT1) 'para que me envien el registro en cualquier caso
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

Public Function GETnnFecha() As Long 'fecha del archivo L37
    GETnnFecha = nnFecha
End Function

Public Function GETnnFecha2() As Long 'fecha del archivo L37
    GETnnFecha2 = nnFecha2
End Function

Public Function GETnnVers() As Long 'fecha del archivo L37
    GETnnVers = nnVers
End Function

Public Function GETnnVers2() As Long 'fecha del archivo L37
    GETnnVers2 = nnVers2
End Function

Public Function GETdifXaSoporte() As Long 'fecha del archivo L37
    GETdifXaSoporte = difXaSoporte
End Function

Public Function GETdifXaVersion() As Long 'fecha del archivo L37
    GETdifXaVersion = difXaVersion
End Function
