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
        'upmanu
        'posibilidad de limpiarlos NO TERMINADO!!!
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
        
        Case "acumsg0": TMP = "jumal.los" 'venta de musica
        Case "acumsg1": TMP = "guen.w" 'venta de musica
        Case "acumsg2": TMP = "japi.lon" 'venta de musica
        Case "acumsg3": TMP = "buca.rest" 'venta de musica
        Case "acumsg4": TMP = "buda.pest" 'venta de musica
        
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
    If Left(LCase(AP), 10) = "d:\dev\3pm" Then Exit Sub
    
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
    Dim newName As String 'nuevo nombre si hay que corregir puntos metidos en el nombre de la carpeta

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
            If RankToPeople = False And NombreDir = "_" + TopListen Then
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
            newName = Replace(NombreDir, ".", "")
            newName = Replace(newName, "#", "")
            newName = Replace(NombreDir, ",", "")
            
            tERR.Anotar "001-0054"
            If NombreDir <> newName Then
            
                tERR.Anotar "001-0055", Ruta2 + NombreDir, Ruta2 + newName
                'si la carpeta de destino ya exista da un error!!!
                If fso.FolderExists(Ruta2 + newName) Then
                    Dim BB As Long, tmpNewName As String
                    'busco un numero que al ponerlo al final no este duplicado
                    For BB = 2 To 1000
                        tmpNewName = newName + CStr(BB)
                        If fso.FolderExists(Ruta2 + tmpNewName) = False Then
                            newName = tmpNewName
                            Exit For
                        End If
                    Next BB
                    newName = tmpNewName
                End If
            
                fso.MoveFolder Ruta2 + NombreDir, Ruta2 + newName

                tERR.Anotar "001-0057", newName
                NombreDir = newName
            End If
            tERR.Anotar "001-0058"
            ParaMatriz = Ruta2 & NombreDir + "," + NombreDir
            tERR.Anotar "001-0059"
            Resultado(ContadorArch) = ParaMatriz
            
            frmINI.lblINI.Caption = ParaMatriz
            frmINI.lblINI.Refresh
            frmINI.PBar.Width = (frmINI.lblINI.Width * ContadorArch / 100) Mod frmINI.lblINI.Width
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
    Dim MinSTR As String 'comparacion de cadenas. Empiezo con el máximo
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
        frmINI.PBar.Width = (frmINI.lblINI.Width * mtx / 100) Mod frmINI.lblINI.Width
        
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
        Dim CarpFull As String, NameCarp As String

        CarpFull = txtInLista(Resultado(Ordenados(mtx)), 1, ",")
        tERR.Anotar "001-0093", CarpFull
        NameCarp = txtInLista(Resultado(Ordenados(mtx)), 2, ",")
        tERR.Anotar "001-0094", NameCarp
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
    TXT.tExt = Texto
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
            Set CarpDest = destino.SubFolders.Add(CarpOrigen.Name)
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
    Dim Fil As Scripting.File, ts As Scripting.TextStream
    Dim pos As Long, ContadorArch As Long
    tERR.Anotar "001-0187"
    ReDim Resultado(50) As Variant
    ' buscar for all the TXT files in the directory
    tERR.Anotar "001-0188"
    For Each Fil In fso.GetFolder(ruta).Files
        tERR.Anotar "001-0189"
        If UCase$(fso.GetExtensionName(Fil.path)) = "TXT" Then
            ' obtener el objeto TextStream correspondiente
            tERR.Anotar "001-0190"
            Set ts = Fil.OpenAsTextStream(ForReading)
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
                Set ts = Fil.OpenAsTextStream(ForReading)
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
                Resultado(ContadorArch) = Array(Fil.path, ts.Line, ts.Column)
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

Public Function GetFullStringArch(Arch As String) As String
    On Error GoTo MiErr
    
    If fso.FileExists(Arch) = False Then
        tERR.Anotar "001-0206", Arch
        GetFullStringArch = "No existe archivo"
        Exit Function
    End If
    tERR.Anotar "001-0208", Arch
    Set TE = fso.OpenTextFile(Arch, ForReading, False)
        Dim Tmp66 As String
        Tmp66 = TE.ReadAll
    TE.Close
    
    GetFullStringArch = Tmp66
    tERR.Anotar "001-0210", Tmp66
    
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

Public Sub EscribirArch1Linea2(Arch As String, TXT As String)
    tERR.Anotar "001-0211v", Arch
    Set TE = fso.CreateTextFile(Arch, True)
        TE.Write TXT
    TE.Close
    tERR.Anotar "001-0213v"
End Sub

'devuelve una matriz con las canciones mas escuchadas
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
    Dim tt As String
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    Dim C As Long: C = 0
    ReDim TMP(0)
    Do While Not TE.AtEndOfStream
        tt = TE.ReadLine
        ThisPTS = Val(txtInLista(tt, 0, ","))
        ThisArch = txtInLista(tt, 1, ",")
        ThisTEMA = txtInLista(tt, 2, ",")
        ThisTEMA = QuitarNumeroDeTema(ThisTEMA)
        ThisDISCO = txtInLista(tt, 3, ",")
            
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
    Optional ordenarABC As Boolean = False, Optional buscarPerfil As Long = 0) As String()
    
    'devuelve "Carpeta + NombreArchivo + "#" + NombreArchivo"
    'devuelve PathFull,SoloNombre
    
    'buscarPerfil se agrega 14/08/2008 para tener en cuenta discos de ringtones, wallpapers y java
    'es un valor de etrada y de salida tambien
    'cuando entra puede ser 0 (cero) para trabajar normal como siempre, solo discos de musica y videos
    'cuando entra en 1 analiza automáticamente perfiles
    'si entra en 100 + X interpreto que esta forzando el disco a ser del perfil X (no programado aun xxxx)

    ' = 1 basico de multimedia
    ' = 2 disco de ringtones
    ' = 3 disco de wallpapers
    ' = 4 discos de java
    ' = 5 disco de imagenes iso / nrg / etc 'mm91
    ' = 6 disco de videos 3gp
    ' = 7 disco de temas para celular

    'ADEMÁS DEBO ASEGURARME QUE NO HAYA COMAS EN LOS NOMBRES
    On Error GoTo ErrObtMM
    tERR.Anotar "001-0214", ordenarABC, buscarPerfil
    If Right(Carpeta, 1) <> "\" Then Carpeta = Carpeta + "\"
    tERR.Anotar "001-0215", Carpeta
    Dim TMPmatriz() As String
    ReDim Preserve TMPmatriz(0)
    'mp3
    Dim NombreArchivo As String, ContadorArch As Long, newName As String
    
    Dim EEXX() As String
    
    Select Case buscarPerfil
        Case 0, 101 'no deberia llamarse como 101 pero por prolijidad esta. esta eleccion es fonola base
            ReDim EEXX(9) 'solo lo de siemrpe basico del 3PM
            EEXX(0) = "mp3"
            EEXX(1) = "wma"
            EEXX(2) = "mpg"
            EEXX(3) = "mpeg"
            EEXX(4) = "avi"
            EEXX(5) = "vob"
            EEXX(6) = "mn0"
            EEXX(7) = "mn1"
            EEXX(8) = "dat"
            EEXX(9) = "wmv"
        Case 2, 102 'ringtones solo en MP3
            ReDim EEXX(0)
            EEXX(0) = "mp3"
        Case 3, 103 'wallpapers
            ReDim EEXX(3)
            EEXX(0) = "jpg" 'NO CONFUNDIR con tapas de discos!
            EEXX(1) = "jpeg"
            EEXX(2) = "bmp"
            EEXX(3) = "gif"
        Case 4, 104 'java
            ReDim EEXX(1)
            EEXX(0) = "jar"
            EEXX(1) = "jad"
        Case 5, 105 'mm91 imagenes iso/nero
            ReDim EEXX(13)
            EEXX(0) = "iso"
            EEXX(1) = "nrg"
            EEXX(2) = "nr3"
            EEXX(3) = "nra"
            EEXX(4) = "nrb"
            EEXX(5) = "nrc"
            EEXX(6) = "nrd"
            EEXX(7) = "nre"
            EEXX(8) = "nrh"
            EEXX(9) = "nri"
            EEXX(10) = "nrm"
            EEXX(11) = "nru"
            EEXX(12) = "nrv"
            EEXX(13) = "nrw"
        Case 6, 106 'mm91 videos 3gp
            ReDim EEXX(0)
            EEXX(0) = "3gp"
        Case 7, 107
            ReDim EEXX(1)
            EEXX(0) = "thm"
            EEXX(1) = "nth"
        Case 1 'deteccio automática
            ReDim EEXX(35) 'ampliado a cualquiera de los perfiles mm91
            EEXX(0) = "mp3"
            EEXX(1) = "wma"
            EEXX(2) = "mpg"
            EEXX(3) = "mpeg"
            EEXX(4) = "avi"
            EEXX(5) = "vob"
            EEXX(6) = "mn0"
            EEXX(7) = "mn1"
            EEXX(8) = "dat"
            EEXX(9) = "wmv"
            EEXX(10) = "" 'reservado para futuros archivos multimedia wav, mp4, midi
            EEXX(11) = ""
            EEXX(12) = ""
            'rigtones "mp3" que ya se leen de todas formas
            EEXX(13) = "" 'deberia identificarlos por el largo NO SE USA CON ESTA EXTENCION, ES SOLO PARA DEMOSTRAR QUE LOS BUSCO
            'wallpapers "jpg", "jpeg", "bmp", "gif"
            EEXX(14) = "jpg" 'NO CONFUNDIR con tapas de discos!
            EEXX(15) = "jpeg"
            EEXX(16) = "bmp"
            EEXX(17) = "gif"
            EEXX(18) = "jar" 'aplicaciones o juegos java
            EEXX(19) = "iso" 'imaganes iso/nero     'mm91
            EEXX(20) = "nrg"
            EEXX(21) = "nr3"
            EEXX(22) = "nra"
            EEXX(23) = "nrb"
            EEXX(24) = "nrc"
            EEXX(25) = "nrd"
            EEXX(26) = "nre"
            EEXX(27) = "nrh"
            EEXX(28) = "nri"
            EEXX(29) = "nrm"
            EEXX(30) = "nru"
            EEXX(31) = "nrv"
            EEXX(32) = "nrw"
            EEXX(33) = "3gp" 'videos para movil  'mm91
            EEXX(34) = "thm" 'temas para movil 'mm91
            EEXX(35) = "nth"
    End Select
    
    Dim ArchMMBase As Long 'cantidad de archivos de musica y videos
    Dim ArchJava As Long 'cantidad de archivos de musica y videos
    Dim ArchImagen As Long 'cantidad de imagenes
    Dim ArchMMRingtone As Long 'cantidad de mp3s o wmas de menos de 1,5 mb (hacer configurable)
    Dim ArchKaraoke As Long 'cantidad de mp3s o wmas de menos de 1,5 mb (hacer configurable)
    Dim ArchTotales As Long 'para saber proporciones de cada uno
    Dim ArchISO As Long 'imagenes ISO o de nero 'mm91
    Dim Arch3GP As Long  'videos para movil 'mm91
    Dim ArchThemes As Long 'temas para movil
    
    'una vez cargado esto por fuera se define un perfil del disco para ver de que tipo es
    'tambie estaría bueno definir un origen ya con características de tipo de disco
    'por que no solo son disco multimedia si no que hay
    'discos de aplicaciones JAVA
    'discos de ringtones (ideal para confundirse con de mp3)
    'discos de wallapers
    
    'deberia dar un perfil a cada disco, si por ejemplo hay muchos mp3 de _
        mas de 2 minutos y aparece un mp3 de 30 segundos y 2 o 3 imagenes _
        deberia darse cuenta que es un disco comun _
        Si la mayoria son imagenes es una carpeta de wallpapers _
        Si la mayoria son mp3 de menos de 2 minutos es una carpeta de ringtones _
        Si hay 10 jar y 10 jpg debo interpretar que es una carpeta de java con _
        sus screenshots correspondientes
        
    'para esto hay un dato que sirve, una cancion MP3 con calidad comun ocupa 1 MB por cada minuto
    'POR ejemplo en el disco "En vivo en Cemento (10-10-1998) - A morir !!! de catupecu _
        hay 4 canciones de menos de 1.5 MB pero hay 18 MP3s de mas de 1.5 MB _
        el sistema debería identificar este disco como de musica !
    
    ArchMMBase = 0
    ArchJava = 0
    ArchImagen = 0
    ArchMMRingtone = 0
    ArchKaraoke = 0
    ArchTotales = 0
    ArchISO = 0 'mm91
    Arch3GP = 0 'mm91
    ArchThemes = 0
    
    Dim H As Long
    For H = 0 To UBound(EEXX)
        If EEXX(H) = "" Then GoTo sigEstaVacio
        
        NombreArchivo = Dir$(Carpeta + "*." + EEXX(H))
        Do While Len(NombreArchivo)
            tERR.Anotar "001-0217", NombreArchivo, EEXX(H)
            'corregir el nombre del tema
            newName = Replace(NombreArchivo, ",", "")
            newName = Replace(newName, "#", "")
            If NombreArchivo <> newName Then
                'no se puede corregir si es un CD. Solo corrige si es disco duro
                'esta funcion se usa para leer CDs debo prevenir
                tERR.Anotar "001-0220", newName
                If fso.Drives(Left(Carpeta, 1)).DriveType = Fixed Then
                    tERR.Anotar "001-0221"
                    'ver si existe lo que se esta por escribir
                    'si es asi elimino el actual
                    If fso.FileExists(Carpeta + newName) Then
                        fso.DeleteFile Carpeta + NombreArchivo, True
                    Else
                        fso.MoveFile Carpeta + NombreArchivo, Carpeta + newName
                    End If
                    tERR.Anotar "001-0222"
                    NombreArchivo = newName
                End If
            End If
            ContadorArch = ContadorArch + 1
            
            'el hecho de diferenciar los tipos de archivos de una carpeta es exclusivo
            'de cuado se le pide a esta funcion que defina el perfil automáticamente
            If buscarPerfil = 1 Then
                Select Case H 'segun este indice es un tipo de archivo diferente
                    Case 0 'mp3
                        'ver si es arch grande o chico!
                        Dim SzeFil As Long
                        SzeFil = FileLen(Carpeta + NombreArchivo)
                        If SzeFil > CLng(1572864) Then ' <<1.5 * (1024 * 1024)>> 1,5 MB es mi base
                            ArchMMBase = ArchMMBase + 1
                        Else
                            ArchMMRingtone = ArchMMRingtone + 1
                        End If
                    Case 1, 2, 3, 4, 5, 8, 9 'wma, mpg, mpeg, avi, vob, dat, wmv
                        ArchMMBase = ArchMMBase + 1
                    Case 6, 7 'mn0,mn1
                        ArchKaraoke = ArchKaraoke + 1
                    Case 14, 15, 16, 17
                        ArchImagen = ArchImagen + 1
                    Case 18
                        ArchJava = ArchJava + 1
                    Case 19 To 32 'mm91
                        ArchISO = ArchISO + 1
                    Case 33 'mm91
                        Arch3GP = Arch3GP + 1
                    Case 34, 35 'mm91
                        ArchThemes = ArchThemes + 1
                End Select
                
                ArchTotales = ArchTotales + 1
            End If
            
            ReDim Preserve TMPmatriz(ContadorArch)
            tERR.Anotar "001-0225", ContadorArch
            TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "#" + NombreArchivo
            NombreArchivo = Dir$
        Loop
sigEstaVacio:
    Next H

    tERR.Anotar "001-0256", buscarPerfil, ArchMMBase, ArchKaraoke, ArchTotales
    tERR.Anotar "001-0256-b", ArchMMRingtone, ArchImagen, ArchJava, ArchISO
    '//////////////////////////////////////////////
    'definir el perfil del disco
    Dim PerfilFINAL As Long 'identificador del perfil
    ' = 1 basico de multimedia
    ' = 2 disco de ringtones
    ' = 3 disco de wallpapers
    ' = 4 discos de java
    
    PerfilFINAL = 0 'quiere decir que es 100% basico (por ahora)
    'pero si pide otra cosa ...
    If buscarPerfil > 0 Then
        'si puso 1 es que quiere que hagamos esto de buscarlo nosotros
        If buscarPerfil = 1 And ArchTotales > 0 Then
            'PERFIL BASICO DE 3PM, MULTIMEDIA PARA ESCUCHAR
            Dim PROP As Single
            PROP = (ArchMMBase + ArchKaraoke) / ArchTotales
            If PROP > 0.6 Then
                PerfilFINAL = 1
            End If
            'perfil de ringtones
            PROP = ArchMMRingtone / ArchTotales
            If PROP > 0.6 Then
                PerfilFINAL = 2
            End If
            'perfil de wallpapers
            PROP = ArchImagen / ArchTotales
            If PROP > 0.6 Then
                PerfilFINAL = 3
            End If
            'perfil de java
            PROP = ArchJava / ArchTotales
            If PROP > 0.4 Then 'aqui es mas sensible porque deberia haber al menos una imagen por cada JAVA o como maximo eso + 1 de la tapa.jpg
                PerfilFINAL = 4
            End If
            PROP = ArchISO / ArchTotales 'mm91
            If PROP > 0.3 Then 'aqui es mas sensible porque deberia haber al menos una imagen por cada JAVA o como maximo eso + 1 de la tapa.jpg
                PerfilFINAL = 5
            End If
            PROP = Arch3GP / ArchTotales 'mm91
            If PROP > 0.3 Then 'aqui es mas sensible porque deberia haber al menos una imagen por cada JAVA o como maximo eso + 1 de la tapa.jpg
                PerfilFINAL = 6
            End If
            PROP = ArchThemes / ArchTotales 'mm91
            If PROP > 0.3 Then 'aqui es mas sensible porque deberia haber al menos una imagen por cada JAVA o como maximo eso + 1 de la tapa.jpg
                PerfilFINAL = 7
            End If
            
            'si hay pocos archivos no me pongo a renegar, es un multimedia
            If ArchTotales < 3 Then PerfilFINAL = 1
            
            '//////////////////////////////////////////////
            'si no entro a ninguno dejo el predeterminado
            If PerfilFINAL = 0 Then PerfilFINAL = 1
            '//////////////////////////////////////////////
            
            
            'ahora que defini el perfil quitar todos los archivo que no correspondan estar en esta lista
            'segun el perfil que se ha determinado correcto.
            'por ejemplo si determinamos que el perfil es java saco de esta lista los archivo de karaokes
            Dim mm As Long, Cancion As String, totEliminados As Long, cadaExtencion As String
            totEliminados = 0
            For mm = 1 To UBound(TMPmatriz) 'esta en base 1
                Cancion = txtInLista(TMPmatriz(mm), 1, "#")
                cadaExtencion = LCase(fso.GetExtensionName(Cancion)) 'minuscula por si quiero comparar con EEXX
                Select Case PerfilFINAL
                    Case 1 'base comun
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If H >= 13 Then
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
                    Case 2 'rigtones
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If H > 0 Then 'solo los mp3 y naaaada mas va aqui
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
                    Case 3 'es de wallpapers
                        'la tapa de los wallpapers no se debe mostrar
                        If LCase(Cancion) = "tapa.jpg" Then TMPmatriz(mm) = ""
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If (H > 17) And (H < 14) Then
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
                    Case 4 'es de java
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If (H <> 18) Then
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
                    Case 5 'es una imagen de disco 'mm91
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If (H < 19) Or (H > 32) Then
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
                    Case 6 'es video para movil 'mm91
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If (H <> 33) Then
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
                    Case 7 'temas para movil 'mm91
                        For H = 0 To UBound(EEXX)
                            If cadaExtencion = EEXX(H) Then
                                If (H < 34) And (H > 35) Then
                                    TMPmatriz(mm) = "" 'si no es de la base lo marco para eliminar!
                                    totEliminados = totEliminados + 1
                                End If
                                Exit For
                            End If
                        Next H
finFOR:
                End Select
            Next mm 'fin de poner en "" todos los que no van
            
            'quitar de la matriz los que no van
            limpiarMtxVacios TMPmatriz
            'XXXXXXXXXXXXXXXXXXXXXXXXXX
            'XXXX queda revisar tooooooodas las llamadas al obtenerArchMM
            'y probar esta funcion que no esta probada !!!
            'ver que detecte los perfiles joiaaaaaaa
            'hacer que muestre cada perfil como corresponde en 3PM
            'XXXXXXXXXXXXXXXXXXXXXXXXXX
        End If
        
        'los otros casos son mas directos
        If buscarPerfil > 100 Then
            PerfilFINAL = buscarPerfil - 100
        End If
        
        
    End If
    '//////////////////////////////////////////////
    
    
    '//////////////////////////////////////////////
    'devuelvo el resultado
    buscarPerfil = PerfilFINAL
    '//////////////////////////////////////////////
    
    'XXXX
    'este ordenar lee las matrices desde 1 hasta ubound
    'y mi matriz usa hasta el cero
    
    'por otra parte no lee todos los discos de wallapers al iniciar
    'igual no puse perfil automático de 3pm al iniciar
    'ya que alli hay una revision de que discos se incluyen y cuales
    'no y se base en que tengan archivos mp3s
    
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

Private Function limpiarMtxVacios(ByRef mtx() As String)
    Dim H As Long, Listo As Boolean
    Listo = False
    Do While Listo = False
        Listo = True
        For H = 1 To UBound(mtx)
            If mtx(H) = "" Then
                quitarElemMatriz mtx, H
                Listo = False 'lo hace quedarse una vuelta mas por las dudas
                Exit For
            End If
        Next H
    Loop
End Function

Private Function quitarElemMatriz(ByRef mtx() As String, Index As Long) As Long
    If Index > UBound(mtx) Then
        quitarElemMatriz = -1
        Exit Function
    End If
    
    If (Index = UBound(mtx)) Then
        If (Index > 0) Then
            ReDim Preserve mtx(Index - 1)
        Else
            ReDim mtx(0) 'mejor la dejo asi, suena a menos errores
            'Erase mtx 'la deja en situacion de error !
        End If
    Else
        Dim H As Long
        For H = Index To (UBound(mtx) - 1)
            mtx(H) = mtx(H + 1)
        Next H
        ReDim Preserve mtx(UBound(mtx) - 1)
    End If
    
    
End Function

Public Function GetTpPred() As String
    Dim iMf2 As String
    If K.sabseee(dcr("1Vx0YVGhEoIisHPLAZMHXw==")) >= Supsabseee Then
        If fso.FileExists(GPF("tddp322")) Then
            iMf2 = GPF("tddp322")
        Else
            iMf2 = ExtraData.getDef.getImagePath("tapapredeterminada")
        End If
    Else
        iMf2 = ExtraData.getDef.getImagePath("tapapredeterminada")
    End If
    
    GetTpPred = iMf2
End Function

Public Sub EsperarSec(Sec As Single)
    Dim T As Single
    T = Timer
    Do While Timer < T + Sec
        DoEvents
    Loop
End Sub

'despues de copiar algo puede quedar turuleco
'asi que voy separando las funciones para reinicar el bluetooth

Public Sub reiniblUtu()
    tERR.Anotar "eaar2f"
    TengoBluetooth = False
    
    tERR.Anotar "eaar3f"
    downblUtu True
    
    tERR.Anotar "eaar4f"
    EsperarSec 3
    
    tERR.Anotar "eaar5f"
    upblUtu
    
    tERR.Anotar "eaar6f"
    TengoBluetooth = True
End Sub

Public Sub downblUtu(killBTM As Boolean)
    tERR.Anotar "eaar7"
    BTM.unInitialize
    
    tERR.Anotar "eaar8"
    tbrBtActivex.ResetWindowMsg
    
    tERR.Anotar "eaar9"
    If killBTM Then Set BTM = Nothing
End Sub

Public Sub upblUtu()
    'indica en el modulo que se usa la referencia al objeto BTManager
    'tbrBtActivex.UsarBluetooth
    tERR.Anotar "eaar"
    Set BTM = tbrBtActivex.btManager
    tERR.Anotar "eaar10"
    tbrBtActivex.SetWindowMsg frmIndex.HWND
    tERR.Anotar "eaar11"
    BTM.Initialize
End Sub
