Attribute VB_Name = "Archivos"
Option Explicit
Option Compare Text

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' API declaración (utilizada por la rutina EsperarPorProceso)
Private Declare Function EsperarUnicoObjeto Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilisegundos As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwAccess As _
    Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Public Sub AbrirArchivo(Arch As String, FrmSolicita As Form)
    tERR.Anotar "001-0001"
    ShellExecute FrmSolicita.HWND, vbNullString, Arch, vbNullString, vbNullString, vbMaximizedFocus
End Sub

Public Function GPF(TXT As String) As String 'GetPathFile
    Dim TMP As String
    Select Case LCase(TXT)
        Case "origs": TMP = "pindo.nga" 'lista de origenes de discos 'EX: sf+ "oddtb.jut"
        Case "creditosactuales": TMP = "kund.era" 'Creditos actuales para usar 'EX: AP + "creditos.tbr"
        Case "config": TMP = "marad.ona" 'sf + "3pmcfg.tbr"
        Case "clavevalid": TMP = "cpd.dor" 'Archivo con la clave de validacion sf +"codped.cfg"
        Case "radliv": TMP = "atak.e77" 'Codigos contados para validacion sf + "radilav.cfg"
        Case "chdc01": TMP = "chd.c01" 'contadores de creditos
        Case "chdc02": TMP = "chd.c02" ' sf + "cc891.dll" hasta 894
        Case "chdc03": TMP = "chd.c03" '
        Case "chdc04": TMP = "chd.c04" '
        Case "chdc05": TMP = "chd.c05" 'contadores de carrito ventas
        Case "chdc06": TMP = "chd.c06" '
        Case "chdc07": TMP = "chd.c07" '
        Case "chdc08": TMP = "chd.c08" '
        
        Case "config2": TMP = "eber.lud" 'copia de seguridad de la config sf + "autoSave3PM.cfg"
        Case "cd3pm": TMP = "cd3.pm" 'Clave sf + "dciLib22.dll"
        Case "cccd3pm": TMP = "cccd3.pm" 'Copia clave sf + "c2LK.dll"
        Case "cd4pm": TMP = "cd4.pm" 'Archivo de licencia 3pm 7.0 (GENERADO)
        Case "cd5pm": TMP = "cd5.pm" 'Archivo RECIBIDO de licencia 3pm 7.0 EN DESUSO
        Case "cd6pm": TMP = "cd6.pm" 'Archivo RECIBIDO de licencia 3pm 7.0 (Copia SEG) EN DESUSO
        
        Case "cd7pm": TMP = "cd7.pm" 'Archivo RECIBIDO de licencia 3pm 7.0 COREGIDO Y EN USO
        Case "cd8pm": TMP = "cd8.pm" 'Archivo RECIBIDO de licencia 3pm 7.0 (Copia SEG) COREGIDO Y EN USO
        
        Case "plin1": TMP = "plin1.ppm" 'mLicencia3PMVtaMusica
        Case "plin2": TMP = "plin2.ppm" 'mLicencia3PMVtaMusica BUP
        
        Case "plin3": TMP = "plin3.ppm" 'mLicencia3PMOrigMusicaFTP
        Case "plin4": TMP = "plin4.ppm" 'mLicencia3PMOrigMusicaFTP BUP
        
        Case "plin5": TMP = "plin5.ppm" 'mLicencia3PMConfigOnline
        Case "plin6": TMP = "plin6.ppm" 'mLicencia3PMConfigOnline BUP
        
        Case "plin7": TMP = "plin7.ppm" 'mLicenciaCD001Kar
        Case "plin8": TMP = "plin8.ppm" 'mLicenciaCD001Kar BUP
        
        Case "plin9": TMP = "plin9.ppm" 'mLicenciaCD002Kar
        Case "plin10": TMP = "plin10.ppm" 'mLicenciaCD002Kar BUP
        
        Case "plin11": TMP = "plin11.ppm" 'mLicenciaCD003Kar
        Case "plin12": TMP = "plin12.ppm" 'mLicenciaCD003Kar BUP
        
        Case "plin13": TMP = "plin13.ppm" 'mLicenciaCD004Kar
        Case "plin14": TMP = "plin14.ppm" 'mLicenciaCD004Kar BUP
        
        Case "plin15": TMP = "plin15.ppm" 'mLicenciaCD005Kar
        Case "plin16": TMP = "plin16.ppm" 'mLicenciaCD005Kar BUP
            
        Case "plin17": TMP = "plin17.ppm" 'mLicenciaCD006Kar
        Case "plin18": TMP = "plin18.ppm" 'mLicenciaCD006Kar BUP
        
        Case "rdcday": TMP = "rdc.day" 'registro diario del contador sf + "daily.cfg"
        Case "pdis233"
            TMP = "pdis.233" 'paquete de imagenes sf + "nev.man"
        Case "extr233": TMP = "" 'sf sola para extraer el paquete de imagenes
            Case "extr233_56": TMP = "f56.dlw" 'logo.sys
            Case "extr233_57": TMP = "f57.dlw" 'logos.sys
            Case "extr233_58": TMP = "f58.dlw" 'logow.sys
            
        Case "233_56_b": TMP = "233.56b" '56 modificado por SL sf + "f5yaSL.nam"
        Case "233_58_b": TMP = "233.58b" '58 modificado por SL sf + "f7yaSL.nam"
        Case "233_57_b": TMP = "233.57b" '57 modif sf + "f6yaSL.nam"
        Case "rempres44": TMP = "rempres.44" 'reemplazo del reserved cuando no hay sf + "razaGUID.dll"
        Case "rempmon45": TMP = "rempmon.45" 'reemplazo para señales de monedero sf + "teclaesp.fas"
        
        Case "telcnot": TMP = "telc.not" 'texto en config No Tbr Wf + "SL\txtCFG.tbr"
        Case "ildw9m": TMP = "ildw9m.811" 'imagen logo.sys del win98/me wf + "img3pm\w\logo.sys"
        Case "ildw9m2": TMP = "ildw9m.812" 'imagen logos.sys del win98/me Wf+ "img3pm\w\logos.sys"
        Case "ildw9m3": TMP = "ildw9m.813" 'imagen logow.sys del win98/me Wf+ "img3pm\w\logow.sys"
        
        Case "ild3pm": TMP = "ild3pm.811" 'imagen logo.sys del 3pm wf + "img3pm\3\logo.sys"
        Case "ild3pm2": TMP = "ild3pm.812" 'imagen logos.sys del 3pm Wf+ "img3pm\3\logos.sys"
        Case "ild3pm3": TMP = "ild3pm.813" 'imagen logow.sys del 3pm Wf+ "img3pm\3\logow.sys"
        
        Case "sequeda32": TMP = "sequed.a32" 'Claves de uso desde afuera de la fonola
        Case "iisl67": TMP = "iis.l67" 'imagen del inicio de la SL Wf+ "SL\imgbig.tbr"
        Case "iischu": TMP = "iis.chu" 'imagen de fondo de las portadas
        Case "tddp322": TMP = "tddp.322" 'tapa predeterminada
        Case "tddp323": TMP = "tddp.323" 'tapa ranking
        Case "tslpri112": TMP = "tslpri.112" 'txtSL principal Wf + "SL\txtIDX.tbr"
        Case "rd3_444": TMP = "rd3.444" 'AP + "ranking.tbr"
        Case "iit17222": TMP = "iit17.222" 'info sobre las imagenes de inicio de win
        Case "adpdp2323": TMP = "adpdp2.323" 'algo del protector
        Case "casc1001": TMP = "casc1.001" 'canciones a seguir cantando
        Case "dalivmp2": TMP = "daliv.mp2" 'archivo con las claves para validar
        Case "cart987": TMP = "cart.987"
        Case "promocart": TMP = "pcraor.mto"
        Case Else: MsgBox "NO SE ENCUENTRA EL ARCHIVO:" + TXT
    End Select
    
    GPF = GetBasePath + TMP
    
End Function

Public Sub BuscarArchivosUbicVieja()

    Dim BasePath As String
    BasePath = GetBasePath

    On Local Error GoTo ErrBAV
    
    Dim ArchAnt As String, ArchNew As String
    Dim SF As String: SF = SYSfolder: Dim WF As String: WF = WINfolder
    
    'lista de origenes de discos
    ArchAnt = SF + "oddtb.jut": ArchNew = BasePath + "pindo.nga"
    ReplaceArch ArchAnt, ArchNew
    
    'creditos actuales
    ArchAnt = AP + "creditos.tbr": ArchNew = BasePath + "kund.era"
    ReplaceArch ArchAnt, ArchNew
    
    'config
    ArchAnt = SF + "3pmcfg.tbr": ArchNew = BasePath + "marad.ona"
    ReplaceArch ArchAnt, ArchNew
    
    'clave de validacion
    ArchAnt = SF + "codped.cfg": ArchNew = BasePath + "cpd.dor"
    ReplaceArch ArchAnt, ArchNew
    
    'Codigos contados para validacion
    ArchAnt = SF + "radilav.cfg": ArchNew = BasePath + "atak.e77"
    ReplaceArch ArchAnt, ArchNew
    
    'contadores de creditos sf + "cc891.dll" hasta 894
    ArchAnt = SF + "cc891.dll": ArchNew = BasePath + "chd.c01"
    ReplaceArch ArchAnt, ArchNew
    ArchAnt = SF + "cc892.dll": ArchNew = BasePath + "chd.c02"
    ReplaceArch ArchAnt, ArchNew
    ArchAnt = SF + "cc893.dll": ArchNew = BasePath + "chd.c03"
    ReplaceArch ArchAnt, ArchNew
    ArchAnt = SF + "cc894.dll": ArchNew = BasePath + "chd.c04"
    ReplaceArch ArchAnt, ArchNew
    
    'copia de seguridad de la config
    ArchAnt = SF + "autoSave3PM.cfg": ArchNew = BasePath + "eber.lud"
    ReplaceArch ArchAnt, ArchNew
    
    'Clave XXXX dejarla del lado freezado ???
    ArchAnt = SF + "dciLib22.dll": ArchNew = BasePath + "cd3.pm"
    ReplaceArch ArchAnt, ArchNew
    
    'dejarla en system para que la version 7 la lea!!!!
    ArchNew = SF + "dciLib22.dll"
    ArchAnt = BasePath + "cd3.pm"
    If fso.FileExists(ArchAnt) Then fso.CopyFile ArchAnt, ArchNew
    
    'Copia clave
    ArchAnt = SF + "c2LK.dll": ArchNew = BasePath + "cccd3.pm"
    ReplaceArch ArchAnt, ArchNew
    
    'registro diario del contador
    ArchAnt = SF + "daily.cfg": ArchNew = BasePath + "rdc.day"
    ReplaceArch ArchAnt, ArchNew
    
    'paquete de imagenes
    ArchAnt = SF + "nev.man": ArchNew = BasePath + "pdis.233"
    ReplaceArch ArchAnt, ArchNew
    
    '56 modificado por SL
    ArchAnt = SF + "f5yaSL.nam": ArchNew = BasePath + "233.56b"
    ReplaceArch ArchAnt, ArchNew
    
    '58 modificado por SL
    ArchAnt = SF + "f7yaSL.nam": ArchNew = BasePath + "233.58b"
    ReplaceArch ArchAnt, ArchNew
    
    'modif 57 (233)
    ArchAnt = SF + "f6yaSL.nam": ArchNew = BasePath + "233.57b"
    ReplaceArch ArchAnt, ArchNew
    
    'modif 54
    ArchAnt = SF + "f9yaSL.nam": ArchNew = BasePath + "233.54b"
    ReplaceArch ArchAnt, ArchNew
    
    ''reemplazo del reserved cuando no hay
    ArchAnt = SF + "razaGUID.dll": ArchNew = BasePath + "rempres.44"
    ReplaceArch ArchAnt, ArchNew
    
    'reemplazo para señales de monedero
    ArchAnt = SF + "teclaesp.fas": ArchNew = BasePath + "rempmon.45"
    ReplaceArch ArchAnt, ArchNew
    
    'la 61 config en SL 'Wf + "SL\indexCHI.tbr"
    ArchAnt = WF + "SL\indexCHI.tbr": ArchNew = BasePath + "61con.f"
    ReplaceArch ArchAnt, ArchNew
    
    'texto en config no Tbrsoft
    ArchAnt = WF + "SL\txtCFG.tbr": ArchNew = BasePath + "telc.not"
    ReplaceArch ArchAnt, ArchNew
    
    'imagen logo.sys del win98/me
    ArchAnt = WF + "img3pm\w\logo.sys": ArchNew = BasePath + "ildw9m.811"
    ReplaceArch ArchAnt, ArchNew
    
    'imagen logos.sys del win98/me Wf + "img3pm\w\logos.sys"
    ArchAnt = WF + "img3pm\w\logos.sys": ArchNew = BasePath + "ildw9m.812"
    ReplaceArch ArchAnt, ArchNew
    
    'imagen logow.sys del win98/me Wf + "img3pm\w\logow.sys"
    ArchAnt = WF + "img3pm\w\logow.sys": ArchNew = BasePath + "ildw9m.813"
    ReplaceArch ArchAnt, ArchNew
    
    'imagen logo.sys del win98/me
    ArchAnt = WF + "img3pm\3\logo.sys": ArchNew = BasePath + "ild3pm.811"
    ReplaceArch ArchAnt, ArchNew
    
    'imagen logos.sys del win98/me Wf + "img3pm\w\logos.sys"
    ArchAnt = WF + "img3pm\3\logos.sys": ArchNew = BasePath + "ild3pm.812"
    ReplaceArch ArchAnt, ArchNew
    
    'imagen logow.sys del win98/me Wf + "img3pm\w\logow.sys"
    ArchAnt = WF + "img3pm\3\logow.sys": ArchNew = BasePath + "ild3pm.813"
    ReplaceArch ArchAnt, ArchNew
    
    'Claves de uso desde afuera de la fonola Wf + "sevalc.dll"
    ArchAnt = WF + "sevalc.dll": ArchNew = BasePath + "sequed.a32"
    ReplaceArch ArchAnt, ArchNew
    
    'imagen del inicio de la SL Wf + "SL\imgbig.tbr"
    ArchAnt = WF + "SL\imgbig.tbr": ArchNew = BasePath + "iis.l67"
    ReplaceArch ArchAnt, ArchNew
    
    'txtSL principal Wf + "SL\txtIDX.tbr"
    ArchAnt = WF + "SL\txtIDX.tbr": ArchNew = BasePath + "tslpri.112"
    ReplaceArch ArchAnt, ArchNew
    
    'ranking de 3pm
    ArchAnt = AP + "ranking.tbr": ArchNew = BasePath + "rd3.444"
    ReplaceArch ArchAnt, ArchNew
    
    'info sobre las imagenes de inicio de win
    ArchAnt = AP + "imgini.tbr": ArchNew = BasePath + "iit17.222"
    ReplaceArch ArchAnt, ArchNew
    
    'algo del protecto de pantalla de tbr
    ArchAnt = AP + "protect.tbr": ArchNew = BasePath + "adpdp2.323"
    ReplaceArch ArchAnt, ArchNew
    
    'canciones a seguir cantando
    ArchAnt = AP + "reini.tbr": ArchNew = BasePath + "casc1.001"
    ReplaceArch ArchAnt, ArchNew
        
    Exit Sub
    
ErrBAV:
    tERR.AppendLog tERR.ErrToTXT(Err), "BAV: " + ArchAnt + " / " + ArchNew
    Resume Next

End Sub

Private Function ReplaceArch(Orig As String, DEST As String) As Long
    
    On Local Error GoTo ErrBAV2
    
    If fso.FileExists(Orig) Then
        If fso.FileExists(DEST) Then fso.DeleteFile DEST, True
        fso.MoveFile Orig, DEST
        ReplaceArch = 0
    End If
    Exit Function
    
ErrBAV2:
    ReplaceArch = 1
    tERR.AppendLog tERR.ErrToTXT(Err), "BAV: " + Orig + " / " + DEST
    Resume Next
    
End Function

Public Function GetBasePath() As String
    
    Dim BasePath As String                   'carpeta de cada archivo usado por 3pm
    
    If fso.FileExists(AP + "sf\bp3.bas") Then
        BasePath = LeerArch1Linea(AP + "sf\bp3.bas")
    Else
        BasePath = AP + "sf\"
    End If
    
    If fso.FolderExists(BasePath) = False Then fso.CreateFolder BasePath
    
    GetBasePath = BasePath
    
End Function

Public Sub ExportarCFG(Optional DestArch As String = "")
    
    Dim F As String
    If DestArch <> "" Then
        'para modo automático
        If fso.FileExists(DestArch) Then fso.DeleteFile DestArch, True
        F = DestArch
    Else
        Dim CmdLg As New CommonDialog
        CmdLg.DialogTitle = TR.Trad("Exportar Archivo de configuración de 3PM%99%")
        CmdLg.ShowSave
    
        F = CmdLg.FileName
        If F = "" Then Exit Sub
    
        If fso.FileExists(F) Then
            TR.SetVars F
            If MsgBox(TR.Trad("Ya existe el archivo " + vbCrLf + _
                "%01%" + vbCrLf + _
                "¿Desea reemplazarlo?%99%"), vbQuestion + vbYesNo) = _
                vbNo Then Exit Sub
        End If
    End If
    
    fso.CopyFile GPF("config"), F, True
    'solo mostrar el mensaje si lo habia abierto el usuario
    If DestArch = "" Then MsgBox TR.Trad("El archivo se exporto correctamente%99%")

End Sub

Public Sub MostrarCursor(Mostrar As Boolean)
    
    'si estoy en el IDE NOLO HAGO!
    'necesito el mouse para depurar!
    If LCase(AP) = "d:\dev\3pm\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pmkun~1\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pmkun~2\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pm kundera 68300\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pm kundera 69000\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pm kundera 69000\dlllistarep\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pm kundera 70000\dlllistarep\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pm kundera 70000\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pm kundera 71000\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pm kundera 71500\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pm kundera 716226\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pm kundera 718047\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pmkun~3\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pmkun~4\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pmkun~5\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pmkun~6\" Then Exit Sub
    If LCase(AP) = "c:\windows\system32\" Then Exit Sub
    If LCase(AP) = "D:\BUP MANUEL\3PM kundera 71500 ale\" Then Exit Sub
    
    
    tERR.Anotar "001-0002"
    Dim A As Long, CONT As Long 'para que no de muchas vueltas !!!
    If Mostrar Then
        A = 0: CONT = 0
        Do While A < 1
            CONT = CONT + 1
            tERR.Anotar "001-0003"
            A = ShowCursor(1) 'suma uno
            ' a clifton se le clavo muchas veces subiendo _
                y subiendo (parece que no llegaba!!!)
            If CONT > 1 Then
                tERR.AppendSinHist "NoShowCur!"
                Exit Sub
            End If
        Loop
    Else
        tERR.Anotar "001-0004"
        A = 1: CONT = 0
        Do While A >= 0
            tERR.Anotar "001-0005"
            A = ShowCursor(0) 'suma uno
            
            If CONT > 1 Then
                tERR.AppendSinHist "NoShowCur!2"
                Exit Sub
            End If
        Loop
    End If
End Sub

' devuelve los atributos de un archivo en un formato legible
' esta rutina también funciona con archivos abiertos
' provoca un error si el archivo no existe

Function ObtAtribDescrip(nombrearch As String) As String
    tERR.Anotar "001-0006"
    Dim Resultado As String, attr As Long
    tERR.Anotar "001-0007"
    attr = GetAttr(nombrearch)
    ' GetAttr también funciona con directorios
    tERR.Anotar "001-0008"
    If attr And vbDirectory Then Resultado = Resultado & " Directorio"
    tERR.Anotar "001-0009"
    If attr And vbReadOnly Then Resultado = Resultado & " Sólo lectura"
    tERR.Anotar "001-0010"
    If attr And vbHidden Then Resultado = Resultado & " Oculto"
    tERR.Anotar "001-0011"
    If attr And vbSystem Then Resultado = Resultado & " Sistema"
    tERR.Anotar "001-0012"
    If attr And vbArchive Then Resultado = Resultado & " Archivo"
    ' descarta el primer espacio
    tERR.Anotar "001-0013"
    ObtAtribDescrip = Mid$(Resultado, 2)
End Function

Function ObtenerArchivos(path As String, EXT As String) As String()
        ' proporciona un array de cadenas que almacenan todos los nombres de archivo que
        ' coinciden con una especificación de archivo dada y unos atributos de búsqueda.
        'devuelve path,nombrearchivo
        tERR.Anotar "001-0014"
        If Right(path, 1) <> "\" Then path = path + "\"
        tERR.Anotar "001-0015"
        Dim Resultado() As String
        Dim NombreArchivo As String, ContadorArch As Long, Ruta2 As String
        Const ALLOC_CHUNK = 50
        tERR.Anotar "001-0016"
        ReDim Resultado(0 To ALLOC_CHUNK) As String
        tERR.Anotar "001-0017"
        NombreArchivo = Dir$(path + EXT)
        tERR.Anotar "001-0018"
        Do While Len(NombreArchivo)
            tERR.Anotar "001-0019"
            ContadorArch = ContadorArch + 1
            tERR.Anotar "001-0020"
            If ContadorArch > UBound(Resultado) Then
                ' cambia el tamaño del array resultado, si es necesario
                tERR.Anotar "001-0021"
                ReDim Preserve Resultado(0 To ContadorArch + ALLOC_CHUNK) As String
            End If
            tERR.Anotar "001-0022"
            Resultado(ContadorArch) = path + NombreArchivo + "," + NombreArchivo
            ' queda preparado para la siguiente iteración
            tERR.Anotar "001-0023"
            NombreArchivo = Dir$
        Loop
        'devuelve el array resultado
        tERR.Anotar "001-0024"
        ReDim Preserve Resultado(0 To ContadorArch) As String
        tERR.Anotar "001-0025"
        ObtenerArchivos = Resultado
End Function

' analiza la existencia de un archivo
Function ExisteArch(NombreArchivo As String) As Boolean
    tERR.Anotar "001-0026"
    On Error Resume Next
    tERR.Anotar "001-0027"
    ExisteArch = (Dir$(NombreArchivo) <> "")
End Function

' verificar si existe un directorio

Function ExisteDir(ruta As String) As Boolean
    tERR.Anotar "001-0028"
    On Error Resume Next
    ExisteDir = (Dir$(ruta & "\nul") <> "")
End Function

' Devuelve un array de cadenas que incluye todos los subdirectorios
' contenidos en una ruta

'corrige ademas los puntos que pueda tener, los saca

'tiene (tenia) metido el mostrador de avance de proceso

Function ObtenerDir(ruta As String) As String()
    tERR.Anotar "001-0029"
    Dim NewName As String 'nuevo nombre si hay que corregir puntos metidos en el nombre de la carpeta

    Dim ParaMatriz As String 'para generar cada elemento de la matriz
    Dim ContadorCarp As Long, CantMP3 As Long
    Dim Resultado() As String
    Dim NombreDir As String, ContadorArch As Long, Ruta2 As String
    tERR.Anotar "001-0035"
    Const ALLOC_CHUNK = 50
    ReDim Resultado(ALLOC_CHUNK) As String
    ' genera el nombre de ruta + barra invertida
    Ruta2 = ruta
    tERR.Anotar "001-0038"
    If Right$(Ruta2, 1) <> "\" Then Ruta2 = Ruta2 & "\"
    tERR.Anotar "001-0039"
    NombreDir = Dir$(Ruta2 & "*.*", vbDirectory)
    tERR.Anotar "001-0040"
    Do While Len(NombreDir)
        tERR.Anotar "001-0041"
        If NombreDir = "." Or NombreDir = ".." Then
            ' excluir las entradas "." y ".."
            tERR.Anotar "001-0042"
        ElseIf (GetAttr(Ruta2 & NombreDir) And vbDirectory) = 0 Then
            ' este es un archivo normal
            tERR.Anotar "001-0043"
        Else
            ' es un directorio
            tERR.Anotar "001-0044"
            If RankToPeople = False And NombreDir = "_Los mas escuchados" Then
                'pasar al que sigue
                tERR.Anotar "001-0045"
                GoTo NextCarp
            End If
            tERR.Anotar "001-0046"
            ContadorArch = ContadorArch + 1
            
            'frmINI.lblINI = "Contando Discos: " + Trim(CStr(ContadorArch))
            tERR.Anotar "001-0047"
            'frmINI.lblINI.Refresh
            tERR.Anotar "001-0048"
            If ContadorArch > UBound(Resultado) Then
                ' cambia el tamaño del array resultante, si
                ' en necesario
                tERR.Anotar "001-0049"
                ReDim Preserve Resultado(ContadorArch + ALLOC_CHUNK) As String
            End If
            
            ContadorCarp = ContadorCarp + 1
            'corregir el nombre del tema
            NewName = Replace(NombreDir, ".", "")
            NewName = Replace(NewName, "#", "")
            
            tERR.Anotar "001-0054"
            If NombreDir <> NewName Then
            
                tERR.Anotar "001-0055", Ruta2 + NombreDir, Ruta2 + NewName
                'si la carpeta de destino ya exista da un error!!!
                If fso.FolderExists(Ruta2 + NewName) Then
                    Dim BB As Long, tmpNewName As String
                    'busco un numero que al ponerlo al final no este duplicado
                    For BB = 2 To 1000
                        tmpNewName = NewName + CStr(BB)
                        If fso.FolderExists(Ruta2 + tmpNewName) = False Then
                            NewName = tmpNewName
                            Exit For
                        End If
                    Next BB
                    NewName = tmpNewName
                End If
            
                fso.MoveFolder Ruta2 + NombreDir, Ruta2 + NewName

                tERR.Anotar "001-0057", NewName
                NombreDir = NewName
            End If
            tERR.Anotar "001-0058"
            ParaMatriz = Ruta2 & NombreDir + "," + NombreDir
            tERR.Anotar "001-0059"
            Resultado(ContadorArch) = ParaMatriz
            
            frmINI.lblINI.Caption = ParaMatriz
            frmINI.lblINI.Refresh
            frmINI.pBAR.Width = (frmINI.lblINI.Width * ContadorArch / 100) Mod frmINI.lblINI.Width
NextCarp:
            
        End If
        tERR.Anotar "001-0066"
        NombreDir = Dir$
        
    Loop
    
solo12: 'solo los 12 primeros
    tERR.Anotar "001-0067"
    'frmINI.PBar.Width = MaxPBAR
    tERR.Anotar "001-0068"
    
    ' proporciona el array resultante
    tERR.Anotar "001-0069"
    ReDim Preserve Resultado(ContadorArch) As String
    
    'tomar la matriz (con valores separador) y ordenala en base a la columna indicada.
    'en este caso el separador es "," y la columna es 0.
    'seria los mismo que tomara 1 ya que todos tienen el mismo path
    tERR.Anotar "001-0070"
    Dim MinSTR As String 'comparacoin de cadenas. Empiezo con el máximo
    Dim ubicMIN As Long 'indice en la matriz del menor encontrado cada vuelta
    tERR.Anotar "001-0071"
    MinSTR = "zzzzzzzzzzzzzzzz"
    tERR.Anotar "001-0072"
    Dim C As Long, mtx As Long, ValComp As String
    C = 0 'cantidad de minimos encontrados
    tERR.Anotar "001-0073"
    Dim Ordenados() As Long 'matriz con los indices ordenados
    Do
        tERR.Anotar "001-0074"
        For mtx = LBound(Resultado) + 1 To UBound(Resultado)
            ValComp = txtInLista(Resultado(mtx), 0, ",")
            If ValComp < MinSTR Then
                MinSTR = ValComp
                ubicMIN = mtx
            End If
        Next
        
        frmINI.lblINI.Caption = CStr(C)
        frmINI.lblINI.Refresh
        frmINI.pBAR.Width = (frmINI.lblINI.Width * mtx / 100) Mod frmINI.lblINI.Width
        
        tERR.Anotar "001-0079"
        Resultado(ubicMIN) = "zzzzzzzzzz," + Resultado(ubicMIN)
        C = C + 1
        ReDim Preserve Ordenados(C)
        tERR.Anotar "001-0080"
        Ordenados(C) = ubicMIN
        tERR.Anotar "001-0081"
        If C >= UBound(Resultado) Then Exit Do
        tERR.Anotar "001-0082"
        MinSTR = "zzzzzzzzzz"
    Loop
    'cargar todos y sacar la primera columna de las zetas
    tERR.Anotar "001-0083"
    Dim MTXsort() As String
    
EntreAlPedo:
    tERR.Anotar "001-0089[" + CStr(LBound(Resultado)) + ":" + CStr(UBound(Resultado)) + "]"
    'si es 0:0 (me pasa en varios)!
    'en ese caso sale del for directamente!
    'entonces dimensiono mtxsort por las dudas!
    
    ReDim MTXsort(0)
    For mtx = LBound(Resultado) + 1 To UBound(Resultado)
        tERR.Anotar "001-0090"
        ReDim Preserve MTXsort(mtx)
        tERR.Anotar "001-0091"
        Dim CarpFull As String, NameCarp As String
        tERR.Anotar "001-0092"
        CarpFull = txtInLista(Resultado(Ordenados(mtx)), 1, ",")
        tERR.Anotar "001-0093"
        NameCarp = txtInLista(Resultado(Ordenados(mtx)), 2, ",")
        tERR.Anotar "001-0094"
        MTXsort(mtx) = CarpFull + "," + NameCarp
    Next
    ObtenerDir = MTXsort
End Function

Public Sub MostrarDiscosMTX()
    'ya debe estar sumada matriz_discos
    
    TOTAL_DISCOS = UBound(MATRIZ_DISCOS) + 1
    
    'ver que haya alguna carpeta
    tERR.Anotar "001-0084"
    If TOTAL_DISCOS < 2 Then
        'VER SI HAY UN DISCO Y NO ES EL RANKING
        tERR.Anotar "001-0085"
        If RankToPeople = False And TOTAL_DISCOS = 1 Then GoTo EntreAlPedo
        tERR.Anotar "001-0086"
'        MsgBox "NO HAY DISCOS PARA MOSTRAR." + vbCrLf + _
'        "Una vez iniciado el sistema presione la tecla " + _
'        "'C' para ingresar a la configuracion y utilize el " + _
'        "asistente para cargar multimedia al sistema."
    End If
    
EntreAlPedo:
    Dim MaxPBAR As Long
    tERR.Anotar "001-0030"
    'MaxPBAR = frmINI.PBar.Width
    
    Dim AY As Long
    Dim nTAPAcd As Integer
    nTAPAcd = 0
    For AY = 0 To UBound(MATRIZ_DISCOS)
        
        tERR.Anotar "001-0095", UBound(MATRIZ_DISCOS)
        UbicDiscoActual = txtInLista(MATRIZ_DISCOS(AY), 0, ",")
        
        Dim CarpFull As String, NameCarp As String
        CarpFull = txtInLista(MATRIZ_DISCOS(AY), 0, ",")
        NameCarp = txtInLista(MATRIZ_DISCOS(AY), 1, ",")
        
        tERR.Anotar "001-0097", CarpFull, nTAPAcd
        'el L es el de los discos en modo texto!
        If nTAPAcd > 0 Then
            tERR.Anotar "001-0098"
            Load frmIndex.L(nTAPAcd)
            tERR.Anotar "001-0099"
            frmIndex.L(nTAPAcd).Top = frmIndex.L(nTAPAcd - 1).Top + frmIndex.L(nTAPAcd - 1).Height
            tERR.Anotar "001-0100"
            frmIndex.L(nTAPAcd).Visible = True
        End If
        tERR.Anotar "001-0101"
        '????¿¿¿¿
        frmIndex.L(nTAPAcd) = NameCarp
        tERR.Anotar "001-0102"
        'frmINI.lblPROCES.AddItem NameCarp, 0
        tERR.Anotar "001-0103"
        'frmINI.lblPROCES.Refresh
        nTAPAcd = nTAPAcd + 1
    Next AY
    
solo12:
    tERR.Anotar "001-0118"
    'frmINI.lblINI = "Proceso terminado, cargando 3PM..."
    'frmINI.lblINI.Refresh
    'frmINI.PBar.Width = MaxPBAR
    
End Sub

'cuenta los archivos de determinada extension contenidos en una carpeta
Public Function ContarArch(ByVal ruta As String, EXT As String, VerSubDIRs As Boolean) As Long
    Dim nombres() As String, I As Long, CONT As Long
    tERR.Anotar "001-0122"
    ' asegurarse de que existe una barra invertida inicial
    If Right(ruta, 1) <> "\" Then ruta = ruta & "\"
    ' obtener la lista de archivos ejecutables
    tERR.Anotar "001-0123"
    nombres() = ObtenerArchivos(ruta, EXT)
    'aqui esta la lista, por ahora no la uso
    'For i = 1 To UBound(nombres)
    '    lst.AddItem Ruta & nombres(i)
    'Next
    tERR.Anotar "001-0124"
    CONT = CONT + UBound(nombres)
    tERR.Anotar "001-0125"
    If VerSubDIRs Then
        ' obtener la lista de subdirectorios, incluyendo los ocultos
        ' y ejecutar recursivamente esta rutina en todos ellos.
        tERR.Anotar "001-0126"
        nombres() = ObtenerDir(ruta)
        tERR.Anotar "001-0127"
        For I = 1 To UBound(nombres)
            tERR.Anotar "001-0128"
            ContarArch ruta & nombres(I), EXT, True
        Next
        tERR.Anotar "001-0129"
    End If
End Function

' carga un archivo de texto en un control TextBox

Sub cargarArchivoEnTextBox(NombreArchivo As String, TXT As TextBox)
    tERR.Anotar "001-0130", NombreArchivo
    Dim numlib As Integer, isOpen As Boolean
    Dim lineatexto As String, Texto As String
    tERR.Anotar "001-0131"
    On Error GoTo Manejador_Error
    ' obtiene el siguiente número libre de archivo
    tERR.Anotar "001-0132"
    numlib = FreeFile()
    tERR.Anotar "001-0133"
    Open NombreArchivo For Input As #numlib
    ' si el flujo llega hasta aquí, se habrán abierto los archivos
    ' sin que se produzca ningún error
    tERR.Anotar "001-0134"
    isOpen = True
    tERR.Anotar "001-0135"
    Do Until EOF(numlib)
        tERR.Anotar "001-0136"
        Line Input #numlib, lineatexto
        tERR.Anotar "001-0137"
        Texto = Texto & lineatexto & vbCrLf
    Loop
    tERR.Anotar "001-0138"
    ' cargar en el cuadro de texto
    TXT.Text = Texto
    ' se cae intencionadamente en el manejador de error para
    ' cerrar el archivo
Manejador_Error:
    tERR.Anotar "001-0139"
    ' se provoca un error(si es que hay alguno), pero primero
    ' se cierra el archivo
    If isOpen Then Close #numlib
    If Err Then Err.Raise Err.Number, , Err.Description
End Sub

' proporciona el contenido completo de un archivo de texto en una cadena

Function LeerContenidoArchTexto(NombreArchivo As String) As String
    tERR.Anotar "001-0140"
    Dim numlib As Integer, isOpen As Boolean
    On Error GoTo Manejador_Error
    ' obtiene el siguiente número libre de archivo
    tERR.Anotar "001-0141"
    numlib = FreeFile()
    tERR.Anotar "001-0142"
    Open NombreArchivo For Input As #numlib
    ' si el flujo llega hasta aquí, se habrán abierto los archivos
    ' sin que se produzca ningún error
    tERR.Anotar "001-0143"
    isOpen = True
    ' leer todo el contenido en una única operación
    tERR.Anotar "001-0144"
    LeerContenidoArchTexto = Input(LOF(numlib), numlib)
    ' se cae intencionadamente en el manejador de error para
    ' cerrar el archivo
Manejador_Error:
    ' se provoca un error(si es que hay alguno), pero primero
    ' se cierra el archivo
    tERR.Anotar "001-0145"
    If isOpen Then Close #numlib
    tERR.Anotar "001-0146"
    If Err Then Err.Raise Err.Number, , Err.Description
End Function

' escribe el contenido de una cadena en un archivo, opcionalmente
' en modo Append

Sub EscribirContenidoArchTexto(Texto As String, NombreArchivo As String, _
    Optional ModoAppend As Boolean)
    tERR.Anotar "001-0147"
        Dim numlib As Integer, isOpen As Boolean
        On Error GoTo Manejador_Error
        ' obtiene el siguiente número libre de archivo
        tERR.Anotar "001-0148"
        numlib = FreeFile()
        tERR.Anotar "001-0149"
        If ModoAppend Then
            tERR.Anotar "001-0150"
            Open NombreArchivo For Append As #numlib
        Else
            tERR.Anotar "001-0151"
            Open NombreArchivo For Output As #numlib
        End If
        ' si el flujo de ejecución llega hasta aquí es que el archivo
        ' se ha abierto correctamente
        tERR.Anotar "001-0152"
        isOpen = True
        ' imprime al archivo en una única operación
        tERR.Anotar "001-0153"
        Print #numlib, Texto
    ' se cae intencionadamente en el manejador de error para
    ' cerrar el archivo
Manejador_Error:
    ' se provoca un error(si es que hay alguno), pero primero
    ' se cierra el archivo
        tERR.Anotar "001-0154"
        If isOpen Then Close #numlib
        tERR.Anotar "001-0155"
        If Err Then Err.Raise Err.Number, , Err.Description
End Sub

' proporciona el contenido de un archivo de texto
' como un array de cadenas.
' NOTA: requiere el empleo de la rutina LeerContenidoArchTexto

Function ObtenLineasArchTexto(NombreArchivo As String, Optional DropEmpty As Boolean, _
    Optional Limite As Variant) As String()
        tERR.Anotar "001-0156"
        Dim ArchTexto As String, elementos() As String, I As Long
        ' lee el contenido del archivo, salir si hay un error
        tERR.Anotar "001-0157"
        ArchTexto = LeerContenidoArchTexto(NombreArchivo)
        ' esto es necesario porque Split() sólo acepta delimitadores de 1 carácter
        tERR.Anotar "001-0158"
        ArchTexto = Replace(ArchTexto, vbCrLf, vbCr)
        ' divide al archivo en líneas individuales de texto
        tERR.Anotar "001-0159"
        elementos() = Split(ArchTexto, vbCr, Limite)
        ' proporciona líneas sencillas, si se solicita
        tERR.Anotar "001-0160"
        If DropEmpty Then
            ' llena las líneas vacías con algo que otros elementos
            ' no contienen
            tERR.Anotar "001-0161"
            For I = 0 To UBound(elementos)
                tERR.Anotar "001-0162"
                    If Len(elementos(I)) = 0 Then elementos(I) = vbCrLf
            Next
            ' utiliza la función Filter() para soltar rápidamente
            ' las líneas vacías
            tERR.Anotar "001-0163"
            elementos() = Filter(elementos(), vbCrLf, False)
        End If
        tERR.Anotar "001-0164"
        ObtenLineasArchTexto = elementos()
End Function

' proporciona el contenido de un archivo de texto delimitado como un
' array de arrays de cadenas.
' NOTA: requiere el empleo de las rutinas LeerContenidoArchTexto y
' ObtenLineasArchTexto

Function ImportarArchDelimitado(NombreArchivo As String, _
    Optional delimitador As String = vbTab) As Variant()
        tERR.Anotar "001-0165"
        Dim lineas() As String, I As Long
        ' obtiene todas las lineas contenidas en el archivo,
        ' ignorando las líneas en blanco
        tERR.Anotar "001-0166"
        lineas() = ObtenLineasArchTexto(NombreArchivo, True)
        ' crea un array de cadena por cada línea de texto
        ' y lo almacena en un elemento Variant
        tERR.Anotar "001-0167"
        ReDim Valores(0 To UBound(lineas)) As Variant
        tERR.Anotar "001-0168"
        For I = 0 To UBound(lineas)
            tERR.Anotar "001-0169"
            Valores(I) = Split(lineas(I), delimitador, -1)
        Next
        tERR.Anotar "001-0170"
        ImportarArchDelimitado = Valores()
End Function

' escribir el contenido de un array de arrays de cadenas a un
' archivo de texto deliminado.
' NOTA: necesita la rutina EscribirContenidoArchTexto

Sub ExportarArchDelimitado(Valores() As Variant, NombreArchivo As String, _
    Optional delimitador As String = vbTab)
        tERR.Anotar "001-0171"
        Dim I As Long, J As Long, ArchTexto As String
        ' reconstruye las líneas individuales de texto del archivo
        tERR.Anotar "001-0172"
        ReDim lineas(0 To UBound(Valores)) As String
        tERR.Anotar "001-0173"
        For I = 0 To UBound(Valores)
            tERR.Anotar "001-0174"
            lineas(I) = Join(Valores(I), delimitador)
        Next
        ' introduce CRLFs entre registros
        tERR.Anotar "001-0175"
        ArchTexto = Replace(Join(lineas, vbCr), vbCr, vbCrLf)
        tERR.Anotar "001-0176"
        EscribirContenidoArchTexto ArchTexto, NombreArchivo
End Sub

' duplica el árbol de directorios sin copiar los archivos

' llamar a esta rutina para iniciar el proceso de copia
' NOTA: la carpeta destino se creará en caso necesario
'       utiliza el procedimiento Private Sub DuplicarDirArbolSub

Sub DuplicarDirArbol(rutaOrigen As String, rutaDest As String)
    tERR.Anotar "001-0177"
    Dim CarpOrigen As Scripting.folder, CarpDest As Scripting.folder
    ' la carpeta origen debe existir
    tERR.Anotar "001-0178"
    Set CarpOrigen = fso.GetFolder(rutaOrigen)
    ' la carpeta destino se creará en caso necesario
    tERR.Anotar "001-0179"
    If fso.FolderExists(rutaDest) Then
        tERR.Anotar "001-0180"
        Set CarpDest = fso.GetFolder(rutaDest)
    Else
        tERR.Anotar "001-0181"
        Set CarpDest = fso.CreateFolder(rutaDest)
    End If
    ' saltar a la rutina recursiva para realizar el trabajo real
    tERR.Anotar "001-0181"
    DuplicarDirArbolSub CarpOrigen, CarpDest
End Sub

' Procedimiento recursivo privado utilizado por DuplicarDirArbol

Private Sub DuplicarDirArbolSub(origen As folder, destino As folder)
    tERR.Anotar "001-0182"
    Dim CarpOrigen As Scripting.folder, CarpDest As Scripting.folder
    tERR.Anotar "001-0183"
    For Each CarpOrigen In origen.SubFolders
        ' copiar esta subcarpeta en la carpeta destino
        tERR.Anotar "001-0184"
            Set CarpDest = destino.SubFolders.add(CarpOrigen.name)
        ' repetir el proceso recursivamente para todas las
        ' subcarpetas de la carpeta considerada
        tERR.Anotar "001-0185"
        DuplicarDirArbolSub CarpOrigen, CarpDest
    Next
End Sub

' Busca una cadena en todos los archivos TXT contenidos en un directorio.

' Por cada archivo localizado devuelve un elemento Variant que contiene
' un array de tres elementos: el nombre del archivo, la línea
' y el número de columna.
' NOTA: las búsquedas no distinguen el uso de mayúsculas y minúsculas

Function BuscarArchTexto(ruta As String, buscar As String) As Variant()
    tERR.Anotar "001-0186"
    Dim fil As Scripting.File, ts As Scripting.TextStream
    Dim pos As Long, ContadorArch As Long
    tERR.Anotar "001-0187"
    ReDim Resultado(50) As Variant
    ' buscar for all the TXT files in the directory
    tERR.Anotar "001-0188"
    For Each fil In fso.GetFolder(ruta).Files
        tERR.Anotar "001-0189"
        If UCase$(fso.GetExtensionName(fil.path)) = "TXT" Then
            ' obtener el objeto TextStream correspondiente
            tERR.Anotar "001-0190"
            Set ts = fil.OpenAsTextStream(ForReading)
            ' leer su contenido, buscar la cadena y cerrarlo
            tERR.Anotar "001-0191"
            pos = InStr(1, ts.ReadAll, buscar, vbTextCompare)
            tERR.Anotar "001-0192"
            ts.Close
            tERR.Anotar "001-0193"
            If pos > 0 Then
                ' si se encuentra la cadena, reabre el archivo
                ' para determinar su posición en forma de (línea, columna)
                tERR.Anotar "001-0194"
                Set ts = fil.OpenAsTextStream(ForReading)
                ' salta todos los caracteres precedentes para saber dónde se
                ' encuentra la cadena
                tERR.Anotar "001-0194"
                ts.Skip pos - 1
                ' llena el array resultado, hace sitio en caso necesario
                ContadorArch = ContadorArch + 1
                tERR.Anotar "001-0195"
                If ContadorArch > UBound(Resultado) Then
                    tERR.Anotar "001-0196"
                    ReDim Preserve Resultado(UBound(Resultado) + 50) As Variant
                End If
                ' cada array resultado tiene tres elementos
                tERR.Anotar "001-0197"
                Resultado(ContadorArch) = Array(fil.path, ts.Line, ts.Column)
                ' ahora podemos cerrar el TextStrem
                tERR.Anotar "001-0198"
                ts.Close
            End If
        End If
    Next
    ' cambia el tamaño del array resultado para indicar el número de
    ' coincidencas
    tERR.Anotar "001-0199"
    ReDim Preserve Resultado(0 To ContadorArch) As Variant
    tERR.Anotar "001-0200"
    BuscarArchTexto = Resultado
End Function

' espera un número de milisegundos y devuelve el estado de ejecución de un
' proceso; si se omite el argumento, espera hasta que el proceso finalice.

Function EsperarPorProceso(taskId As Long, Optional msecs As Long = -1) _
    As Boolean
        tERR.Anotar "001-0201"
        Dim procHandle As Long
        ' obtiene el manejador del proceso
        tERR.Anotar "001-0202"
        procHandle = OpenProcess(&H100000, True, taskId)
        ' verifica su estado señalado, lo devuelve al que hizo la llamada
        tERR.Anotar "001-0203"
        EsperarPorProceso = EsperarUnicoObjeto(procHandle, msecs) <> -1
        ' cierra el gestor
        tERR.Anotar "001-0204"
        CloseHandle procHandle
End Function

Public Function LeerArch1Linea(Arch As String) As String
    On Error GoTo MiErr
    
    If fso.FileExists(Arch) = False Then
        tERR.Anotar "001-0206", Arch
        LeerArch1Linea = "No existe archivo"
        Exit Function
    End If
    tERR.Anotar "001-0208", Arch
    Set TE = fso.OpenTextFile(Arch, ForReading, False)
        Dim Tmp66 As String
        Tmp66 = TE.ReadLine
        LeerArch1Linea = Tmp66
        tERR.Anotar "001-0210", Tmp66
    TE.Close
    
    Exit Function
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), "Archivos.bas" + ".acpk"
    Resume Next
End Function

Public Sub EscribirArch1Linea(Arch As String, TXT As String)
    tERR.Anotar "001-0211"
    Set TE = fso.CreateTextFile(Arch, True)
    tERR.Anotar "001-0212"
    TE.WriteLine TXT
    tERR.Anotar "001-0213"
    TE.Close
End Sub

Public Function ObtenerRankComoMM(Optional MaxTop As Long = 15) As String()

    Dim TMP() As String
    If fso.FileExists(GPF("rd3_444")) = False Then
        'no hay nada !!!!
        fso.CreateTextFile GPF("rd3_444"), True
        ReDim TMP(0)
        ObtenerRankComoMM = TMP
        Exit Function
    End If
    
    Set TE = fso.OpenTextFile(GPF("rd3_444"), ForReading, False)
    Dim TT As String
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    Dim C As Long: C = 0
    ReDim TMP(0)
    Do While Not TE.AtEndOfStream
        TT = TE.ReadLine
        ThisPTS = Val(txtInLista(TT, 0, ","))
        ThisArch = txtInLista(TT, 1, ",")
        ThisTEMA = txtInLista(TT, 2, ",")
        ThisTEMA = QuitarNumeroDeTema(ThisTEMA)
        ThisDISCO = txtInLista(TT, 3, ",")
            
        If C = MaxTop Then Exit Do
        'si elarchivo no existe no se debe cargar
        If fso.FileExists(ThisArch) Then
        
            C = C + 1 'la otra matriz obtenerMM empieza en 1 entonces esta tambien y la hacemos compatible
            
            ReDim Preserve TMP(C)
            TMP(C) = ThisArch + "#" + CStr(C) + "º (" + CStr(ThisPTS) + ") -" + ThisTEMA
            
        End If
    Loop
    
    ObtenerRankComoMM = TMP
    
End Function

Public Function ObtenerArchMM(Carpeta As String, _
    Optional ordenarABC As Boolean = False) As String()
    
    
    'devuelve "Carpeta + NombreArchivo + "#" + NombreArchivo"
    'devuelve PathFull,SoloNombre

    'ADEMÁS DEBO ASEGURARME QUE NO HAYA COMAS EN LOS NOMBRES
    On Error GoTo ErrObtMM
    tERR.Anotar "001-0214"
    If Right(Carpeta, 1) <> "\" Then Carpeta = Carpeta + "\"
    tERR.Anotar "001-0215", Carpeta
    Dim TMPmatriz() As String
    ReDim Preserve TMPmatriz(0)
    'mp3
    Dim NombreArchivo As String, ContadorArch As Long, NewName As String
    
    Dim EEXX(9) As String
    EEXX(0) = "mp3": EEXX(1) = "wma": EEXX(2) = "mpg": EEXX(3) = "mpeg": EEXX(4) = "avi"
    EEXX(5) = "vob": EEXX(6) = "mn0": EEXX(7) = "mn1": EEXX(8) = "dat": EEXX(9) = "wmv"
    
    Dim H As Long
    For H = 0 To UBound(EEXX)
        NombreArchivo = Dir$(Carpeta + "*." + EEXX(H))
        Do While Len(NombreArchivo)
            tERR.Anotar "001-0217", NombreArchivo, EEXX(H)
            'corregir el nombre del tema
            NewName = Replace(NombreArchivo, ",", "")
            NewName = Replace(NewName, "#", "")
            If NombreArchivo <> NewName Then
                'no se puede corregir si es un CD. Solo corrige si es disco duro
                'esta funcion se usa para leer CDs debo prevenir
                tERR.Anotar "001-0220", NewName
                If fso.Drives(Left(Carpeta, 1)).DriveType = Fixed Then
                    tERR.Anotar "001-0221"
                    'ver si existe lo que se esta por escribir
                    'si es asi elimino el actual
                    If fso.FileExists(Carpeta + NewName) Then
                        fso.DeleteFile Carpeta + NombreArchivo, True
                    Else
                        fso.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
                    End If
                    tERR.Anotar "001-0222"
                    NombreArchivo = NewName
                End If
            End If
            ContadorArch = ContadorArch + 1
            ReDim Preserve TMPmatriz(ContadorArch)
            tERR.Anotar "001-0225", ContadorArch
            TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "#" + NombreArchivo
            NombreArchivo = Dir$
        Loop
    
    Next H

    tERR.Anotar "001-0256"
    
    If ordenarABC Then
        Dim TMP2Matriz() As String
        ReDim TMP2Matriz(0)
        Dim OKs As Long
        
        Dim K As Long, L As Long
        Dim CPR As String 'comparador
        CPR = "ZZZ"
        Dim Min As Long 'indice del minimo
        Min = 0
        For L = 1 To UBound(TMPmatriz) 'esta en base 1
            CPR = "ZZZZZZZZZ"
            For K = 1 To UBound(TMPmatriz) 'esta en base 1
                If TMPmatriz(K) < CPR Then
                    CPR = TMPmatriz(K)
                    Min = K
                End If
            Next K
            
            OKs = OKs + 1
            ReDim Preserve TMP2Matriz(OKs)
            TMP2Matriz(OKs) = CPR
            TMPmatriz(Min) = "ZZZZZZZZZ"
        Next L
        
        ObtenerArchMM = TMP2Matriz
        
    Else
        ObtenerArchMM = TMPmatriz
    End If
    
    
    
    Exit Function
ErrObtMM:
    tERR.AppendLog tERR.ErrToTXT(Err), "Archivos.bas" + ".acpk4"
    Resume Next
    
End Function
