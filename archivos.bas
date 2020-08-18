Attribute VB_Name = "Archivos"
Option Explicit
Option Compare Text

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


' API declaración (utilizada por la rutina EsperarPorProceso)
Private Declare Function EsperarUnicoObjeto Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilisegundos As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwAccess As _
    Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Public Sub AbrirArchivo(Arch As String, FrmSolicita As Form)
    tERR.Anotar "001-0001"
    ShellExecute FrmSolicita.hwnd, vbNullString, Arch, vbNullString, vbNullString, vbMaximizedFocus
End Sub

Public Sub ExportarCFG(Optional DestArch As String = "")
    
    Dim F As String
    If DestArch <> "" Then
        'para modo automático
        If FSO.FileExists(DestArch) Then FSO.DeleteFile DestArch, True
        F = DestArch
    Else
        Dim CmdLg As New CommonDialog
        CmdLg.DialogTitle = "Exportar Archivo de configuración de 3PM"
        CmdLg.ShowSave
    
        F = CmdLg.FileName
        If F = "" Then Exit Sub
    
        If FSO.FileExists(F) Then
            If MsgBox("Ya existe el archivo " + vbCrLf + F + _
                vbCrLf + "¿Desea reemplazarlo?", vbQuestion + vbYesNo) = _
                vbNo Then Exit Sub
        End If
    End If
    
    
    FSO.CopyFile SYSfolder + "3pmcfg.tbr", F, True
    'solo mostrar el mensaje si lo habia abierto el usuario
    If DestArch = "" Then MsgBox "El archivo se exporto correctamente"

End Sub

Public Sub MostrarCursor(Mostrar As Boolean)
    
    'si estoy en el IDE NOLO HAGO!
    'necesito el mouse para depurar!
    If LCase(AP) = "d:\dev\3pm\" Then Exit Sub
    If LCase(AP) = "d:\dev\3pm kundera 68200\" Then Exit Sub

    tERR.Anotar "001-0002"
    Dim A As Long
    If Mostrar Then
        A = 0
        Do While A < 1
            tERR.Anotar "001-0003"
            A = ShowCursor(1) 'suma uno
        Loop
    Else
        tERR.Anotar "001-0004"
        A = 1
        Do While A >= 0
            tERR.Anotar "001-0005"
            A = ShowCursor(0) 'suma uno
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

'tiene metido el mostrador de avance de proceso

Function ObtenerDir(ruta As String) As String()
    tERR.Anotar "001-0029"
    Dim NewName As String 'nuevo nombre si hay que corregir puntos metidos en el nombre de la carpeta
    Dim MaxPBAR As Long
    tERR.Anotar "001-0030"
    MaxPBAR = frmINI.pBar.Width
    tERR.Anotar "001-0031"
    frmINI.pBar.Width = 0
    tERR.Anotar "001-0032"
    frmINI.lblPROCES.Clear
    frmINI.lblPROCES.AddItem "Iniciando busqueda", 0
    tERR.Anotar "001-0033"
    frmINI.lblPROCES.Refresh
    tERR.Anotar "001-0034"
    Dim ParaMatriz As String 'para generar cada elemento de la matriz
    Dim ContadorCarp As Long, CantMP3 As Long
    Dim Resultado() As String
    Dim NombreDir As String, ContadorArch As Long, Ruta2 As String
    tERR.Anotar "001-0035"
    Const ALLOC_CHUNK = 50
    tERR.Anotar "001-0036"
    ReDim Resultado(ALLOC_CHUNK) As String
    tERR.Anotar "001-0037"
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
            If RankToPeople = False And NombreDir = "01- Los mas escuchados" Then
                'pasar al que sigue
                tERR.Anotar "001-0045"
                GoTo NextCarp
            End If
            tERR.Anotar "001-0046"
            ContadorArch = ContadorArch + 1
            
            frmINI.lblINI = "Contando Discos: " + Trim(Str(ContadorArch))
            tERR.Anotar "001-0047"
            frmINI.lblINI.Refresh
            tERR.Anotar "001-0048"
            If ContadorArch > UBound(Resultado) Then
                ' cambia el tamaño del array resultante, si
                ' en necesario
                tERR.Anotar "001-0049"
                ReDim Preserve Resultado(ContadorArch + ALLOC_CHUNK) As String
            End If
            tERR.Anotar "001-0050"
            frmINI.pBar.Width = frmINI.pBar.Width + 100
            'si me hacerco al max de pbar lo hago inalcanzable
            tERR.Anotar "001-0051"
            frmINI.lblPROCES.AddItem NombreDir, 0
            tERR.Anotar "001-0052"
            frmINI.lblPROCES.Refresh
            tERR.Anotar "001-0053"
            ContadorCarp = ContadorCarp + 1
            'corregir el nombre del tema
            NewName = QuitarCaracter(NombreDir, ".")
            tERR.Anotar "001-0054"
            If NombreDir <> NewName Then
            
                tERR.Anotar "001-0055", Ruta2 + NombreDir, Ruta2 + NewName
                'si la carpeta de destino ya exista da un error!!!
                If FSO.FolderExists(Ruta2 + NewName) Then
                    Dim BB As Long, tmpNewName As String
                    'busco un numero que al ponerlo al final no este duplicado
                    For BB = 2 To 1000
                        tmpNewName = NewName + CStr(BB)
                        If FSO.FolderExists(Ruta2 + tmpNewName) = False Then
                            NewName = tmpNewName
                            Exit For
                        End If
                    Next BB
                    NewName = tmpNewName
                End If
            
                FSO.MoveFolder Ruta2 + NombreDir, Ruta2 + NewName

                tERR.Anotar "001-0057", NewName
                NombreDir = NewName
            End If
            tERR.Anotar "001-0058"
            ParaMatriz = Ruta2 & NombreDir + "," + NombreDir
            tERR.Anotar "001-0059"
            Resultado(ContadorArch) = ParaMatriz
NextCarp:
            
        End If
        tERR.Anotar "001-0066"
        NombreDir = Dir$
        
    Loop
solo12: 'solo los 12 primeros
    tERR.Anotar "001-0067"
    frmINI.pBar.Width = MaxPBAR
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
    Dim c As Long, mtx As Long, ValComp As String
    c = 0 'cantidad de minimos encontrados
    tERR.Anotar "001-0073"
    Dim Ordenados() As Long 'matriz con los indices ordenados
    Do
        tERR.Anotar "001-0074"
        For mtx = LBound(Resultado) + 1 To UBound(Resultado)
            tERR.Anotar "001-0075"
            ValComp = txtInLista(Resultado(mtx), 0, ",")
            tERR.Anotar "001-0076"
            If ValComp < MinSTR Then
                tERR.Anotar "001-0077"
                MinSTR = ValComp
                tERR.Anotar "001-0078"
                ubicMIN = mtx
            End If
        Next
        tERR.Anotar "001-0079"
        Resultado(ubicMIN) = "zzzzzzzzzz," + Resultado(ubicMIN)
        c = c + 1
        ReDim Preserve Ordenados(c)
        tERR.Anotar "001-0080"
        Ordenados(c) = ubicMIN
        tERR.Anotar "001-0081"
        If c >= UBound(Resultado) Then Exit Do
        tERR.Anotar "001-0082"
        MinSTR = "zzzzzzzzzz"
    Loop
    'cargar todos y sacar la primera columna de las zetas
    tERR.Anotar "001-0083"
    Dim MTXsort() As String
    
EntreAlPedo:
    tERR.Anotar "001-0087"
    frmINI.pBar.Width = 0
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
        MsgBox "NO HAY DISCOS PARA MOSTRAR." + vbCrLf + _
        "Una vez iniciado el sistema presione la tecla " + _
        "'C' para ingresar a la configuracion y utilize el " + _
        "asistente para cargar multimedia al sistema."
    End If
EntreAlPedo:
    Dim MaxPBAR As Long
    tERR.Anotar "001-0030"
    MaxPBAR = frmINI.pBar.Width
    
    Dim AY As Long
    Dim nTAPAcd As Integer
    nTAPAcd = 0
    For AY = 0 To UBound(MATRIZ_DISCOS)
        
        tERR.Anotar "001-0095"
        UbicDiscoActual = txtInLista(MATRIZ_DISCOS(AY), 0, ",")
        
        Dim CarpFull As String, NameCarp As String
        CarpFull = txtInLista(MATRIZ_DISCOS(AY), 0, ",")
        NameCarp = txtInLista(MATRIZ_DISCOS(AY), 1, ",")
        
        If CargarIMGinicio Then
            
            'caragar las imágenes en diferentes IMGs para que no se cargen despues
            Dim ArchTapa As String
            tERR.Anotar "001-0096"
            ArchTapa = UbicDiscoActual + "\tapa.jpg"
            'arranca con 5 ya cargados
            
            tERR.Anotar "001-0105"
            frmINI.lblINI = "Ordenando Discos: " + Trim(Str(AY))
            tERR.Anotar "001-0106"
            frmINI.lblINI.Refresh
            tERR.Anotar "001-0107"
            frmINI.pBar.Width = frmINI.pBar.Width + 100
            tERR.Anotar "001-0108"
            frmINI.pBar.Refresh
            tERR.Anotar "001-0109"
            'si paso una pàgina....
            If nTAPAcd > ((TapasMostradasH * TapasMostradasV) - 1) Then
                tERR.Anotar "001-0110"
                Load frmIndex.TapaCD(nTAPAcd)
                tERR.Anotar "001-0111"
                '.... se ubica en la misma ubicacion que el mismo dsico de la página anterior
                frmIndex.TapaCD(nTAPAcd).Left = frmIndex.TapaCD(nTAPAcd - ((TapasMostradasH * TapasMostradasV))).Left
                tERR.Anotar "001-0112"
                frmIndex.TapaCD(nTAPAcd).Top = frmIndex.TapaCD(nTAPAcd - ((TapasMostradasH * TapasMostradasV))).Top
                tERR.Anotar "001-0113"
            End If
            tERR.Anotar "001-0114"
            If FSO.FileExists(ArchTapa) Then
                tERR.Anotar "001-0115", ArchTapa
                'XXXX
                'VERIFICAR EL ERROR 481 DE IMAGEN NO VALIDA!!!!
                If FileLen(ArchTapa) > 50000 Then
                    tERR.Anotar "acgf3", ArchTapa, CStr(FileLen(ArchTapa))
                    GoTo TAPADEF
                End If
                frmIndex.TapaCD(nTAPAcd).Picture = LoadPicture(ArchTapa)
            Else
TAPADEF:
                tERR.Anotar "001-0116"
                'ver si hay SuperLicencia!!!
                If FSO.FileExists(WINfolder + "SL\indexCHI.tbr") Then
                    frmIndex.TapaCD(nTAPAcd).Picture = LoadPicture(WINfolder + "SL\indexCHI.tbr")
                Else
                    frmIndex.TapaCD(nTAPAcd).Picture = LoadPicture(SYSfolder + "f61.dlw")
                End If
            End If
            tERR.Anotar "001-0117"
        End If
        tERR.Anotar "001-0097"
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
        frmINI.lblPROCES.AddItem NameCarp, 0
        tERR.Anotar "001-0103"
        frmINI.lblPROCES.Refresh
        nTAPAcd = nTAPAcd + 1
    Next AY
    
solo12:
    tERR.Anotar "001-0118"
    frmINI.lblINI = "Proceso terminado, cargando 3PM..."
    tERR.Anotar "001-0119"
    frmINI.lblINI.Refresh
    tERR.Anotar "001-0120"
    frmINI.pBar.Width = MaxPBAR
    
End Sub

'cuenta los archivos de determinada extension contenidos en una carpeta
Public Function ContarArch(ByVal ruta As String, EXT As String, VerSubDIRs As Boolean) As Long
    Dim nombres() As String, i As Long, CONT As Long
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
        For i = 1 To UBound(nombres)
            tERR.Anotar "001-0128"
            ContarArch ruta & nombres(i), EXT, True
        Next
        tERR.Anotar "001-0129"
    End If
End Function

' carga un archivo de texto en un control TextBox

Sub cargarArchivoEnTextBox(NombreArchivo As String, TXT As TextBox)
    tERR.Anotar "001-0130"
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
        Dim ArchTexto As String, elementos() As String, i As Long
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
            For i = 0 To UBound(elementos)
                tERR.Anotar "001-0162"
                If Len(elementos(i)) = 0 Then elementos(i) = vbCrLf
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
        Dim lineas() As String, i As Long
        ' obtiene todas las lineas contenidas en el archivo,
        ' ignorando las líneas en blanco
        tERR.Anotar "001-0166"
        lineas() = ObtenLineasArchTexto(NombreArchivo, True)
        ' crea un array de cadena por cada línea de texto
        ' y lo almacena en un elemento Variant
        tERR.Anotar "001-0167"
        ReDim valores(0 To UBound(lineas)) As Variant
        tERR.Anotar "001-0168"
        For i = 0 To UBound(lineas)
            tERR.Anotar "001-0169"
            valores(i) = Split(lineas(i), delimitador, -1)
        Next
        tERR.Anotar "001-0170"
        ImportarArchDelimitado = valores()
End Function

' escribir el contenido de un array de arrays de cadenas a un
' archivo de texto deliminado.
' NOTA: necesita la rutina EscribirContenidoArchTexto

Sub ExportarArchDelimitado(valores() As Variant, NombreArchivo As String, _
    Optional delimitador As String = vbTab)
        tERR.Anotar "001-0171"
        Dim i As Long, J As Long, ArchTexto As String
        ' reconstruye las líneas individuales de texto del archivo
        tERR.Anotar "001-0172"
        ReDim lineas(0 To UBound(valores)) As String
        tERR.Anotar "001-0173"
        For i = 0 To UBound(valores)
            tERR.Anotar "001-0174"
            lineas(i) = Join(valores(i), delimitador)
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
    Dim CarpOrigen As Scripting.Folder, CarpDest As Scripting.Folder
    ' la carpeta origen debe existir
    tERR.Anotar "001-0178"
    Set CarpOrigen = FSO.GetFolder(rutaOrigen)
    ' la carpeta destino se creará en caso necesario
    tERR.Anotar "001-0179"
    If FSO.FolderExists(rutaDest) Then
        tERR.Anotar "001-0180"
        Set CarpDest = FSO.GetFolder(rutaDest)
    Else
        tERR.Anotar "001-0181"
        Set CarpDest = FSO.CreateFolder(rutaDest)
    End If
    ' saltar a la rutina recursiva para realizar el trabajo real
    tERR.Anotar "001-0181"
    DuplicarDirArbolSub CarpOrigen, CarpDest
End Sub

' Procedimiento recursivo privado utilizado por DuplicarDirArbol

Private Sub DuplicarDirArbolSub(origen As Folder, destino As Folder)
    tERR.Anotar "001-0182"
    Dim CarpOrigen As Scripting.Folder, CarpDest As Scripting.Folder
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
    Dim fil As Scripting.File, ts As Scripting.TextStream
    Dim pos As Long, ContadorArch As Long
    tERR.Anotar "001-0187"
    ReDim Resultado(50) As Variant
    ' buscar for all the TXT files in the directory
    tERR.Anotar "001-0188"
    For Each fil In FSO.GetFolder(ruta).Files
        tERR.Anotar "001-0189"
        If UCase$(FSO.GetExtensionName(fil.path)) = "TXT" Then
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
    
    tERR.Anotar "001-0205", Arch
    If FSO.FileExists(Arch) = False Then
        tERR.Anotar "001-0206"
        LeerArch1Linea = "No existe archivo"
        tERR.Anotar "001-0207"
        Exit Function
    End If
    tERR.Anotar "001-0208"
    Set TE = FSO.OpenTextFile(Arch, ForReading, False)
    tERR.Anotar "001-0209"
    LeerArch1Linea = TE.ReadLine
    tERR.Anotar "001-0210", LeerArch1Linea
    TE.Close
    
    Exit Function
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), "Archivos.bas" + ".acpk"
    Resume Next
End Function

Public Sub EscribirArch1Linea(Arch As String, TXT As String)
    tERR.Anotar "001-0211"
    Set TE = FSO.CreateTextFile(Arch, True)
    tERR.Anotar "001-0212"
    TE.WriteLine TXT
    tERR.Anotar "001-0213"
    TE.Close
End Sub

Public Function ObtenerArchMM(Carpeta As String) As String()
    'devuelve "Carpeta + NombreArchivo + "#" + NombreArchivo"
    'devuelve PathFull,SoloNombre

    'ADEMÁS DEBO ASEGURARME QUE NO HAYA COMAS EN LOS NOMBRES
    On Error GoTo ErrObtMM
    tERR.Anotar "001-0214"
    If Right(Carpeta, 1) <> "\" Then Carpeta = Carpeta + "\"
    tERR.Anotar "001-0215"
    Dim TMPmatriz() As String
    ReDim Preserve TMPmatriz(0)
    'mp3
    tERR.Anotar "001-0216"
    Dim NombreArchivo As String, ContadorArch As Long, NewName As String
    NombreArchivo = Dir$(Carpeta + "*.mp3")
    tERR.Anotar "001-0217"
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        tERR.Anotar "001-0218"
        NewName = QuitarCaracter(NombreArchivo, ",")
        tERR.Anotar "001-0219"
        If NombreArchivo <> NewName Then
            'no se puede corregir si es un CD. Solo corrige si es disco duro
            'esta funcion se usa para leer CDs debo prevenir
            tERR.Anotar "001-0220"
            If FSO.Drives(Left(Carpeta, 1)).DriveType = Fixed Then
                tERR.Anotar "001-0221"
                FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
                tERR.Anotar "001-0222"
                NombreArchivo = NewName
            End If
        End If
        tERR.Anotar "001-0224"
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        tERR.Anotar "001-0225"
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "#" + NombreArchivo
        tERR.Anotar "001-0226"
        NombreArchivo = Dir$
    Loop
    
'    '''MP4
'    tERR.Anotar "001-0216b"
'    NombreArchivo = Dir$(Carpeta + "*.mp4")
'    tERR.Anotar "001-0217b"
'    Do While Len(NombreArchivo)
'        'corregir el nombre del tema
'        tERR.Anotar "001-0218b", NombreArchivo
'        NewName = QuitarCaracter(NombreArchivo, ",")
'        tERR.Anotar "001-0219b"
'        If NombreArchivo <> NewName Then
'            'no se puede corregir si es un CD. Solo corrige si es disco duro
'            'esta funcion se usa para leer CDs debo prevenir
'            tERR.Anotar "001-0220b"
'            If FSO.Drives(Left(Carpeta, 1)).DriveType = Fixed Then
'                tERR.Anotar "001-0221b"
'                FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
'                tERR.Anotar "001-0222b"
'                NombreArchivo = NewName
'            End If
'        End If
'        tERR.Anotar "001-0224b"
'        ContadorArch = ContadorArch + 1
'        ReDim Preserve TMPmatriz(ContadorArch)
'        tERR.Anotar "001-0225b"
'        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "#" + NombreArchivo
'        tERR.Anotar "001-0226b"
'        NombreArchivo = Dir$
'    Loop
    
    'WMA
    tERR.Anotar "001-0216b"
    NombreArchivo = Dir$(Carpeta + "*.wma")
    tERR.Anotar "001-0217"
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        tERR.Anotar "001-0218"
        NewName = QuitarCaracter(NombreArchivo, ",")
        tERR.Anotar "001-0219"
        If NombreArchivo <> NewName Then
            'no se puede corregir si es un CD. Solo corrige si es disco duro
            'esta funcion se usa para leer CDs debo prevenir
            tERR.Anotar "001-0220"
            If FSO.Drives(Left(Carpeta, 1)).DriveType = Fixed Then
                tERR.Anotar "001-0221"
                FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
                tERR.Anotar "001-0222"
                NombreArchivo = NewName
            End If
        End If
        tERR.Anotar "001-0224"
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        tERR.Anotar "001-0225"
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "#" + NombreArchivo
        tERR.Anotar "001-0226"
        NombreArchivo = Dir$
    Loop
    
    'mpg
    tERR.Anotar "001-0227"
    NombreArchivo = Dir$(Carpeta + "\*.mpg")
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        tERR.Anotar "001-0228"
        NewName = QuitarCaracter(NombreArchivo, ",")
        tERR.Anotar "001-0229"
        If NombreArchivo <> NewName Then
            tERR.Anotar "001-0230"
            FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
            tERR.Anotar "001-0231"
            'WriteTBRLog "Se corrigio el nombre de tema " + NombreArchivo + _
                " por " + NewName + " en la carpeta " + Carpeta, True
            tERR.Anotar "001-0232"
            NombreArchivo = NewName
        End If
        tERR.Anotar "001-0233"
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        tERR.Anotar "001-0234"
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "#" + NombreArchivo
        tERR.Anotar "001-0235"
        NombreArchivo = Dir$
    Loop
    
    'mpeg
    tERR.Anotar "001-0236"
    NombreArchivo = Dir$(Carpeta + "\*.mpeg")
    tERR.Anotar "001-0237"
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        tERR.Anotar "001-0238"
        NewName = QuitarCaracter(NombreArchivo, ",")
        tERR.Anotar "001-0239"
        If NombreArchivo <> NewName Then
            tERR.Anotar "001-0240"
            FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
            tERR.Anotar "001-0241"
            'WriteTBRLog "Se corrigio el nombre de tema " + NombreArchivo + _
                " por " + NewName + " en la carpeta " + Carpeta, True
            tERR.Anotar "001-0242"
            NombreArchivo = NewName
        End If
        tERR.Anotar "001-0243"
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        tERR.Anotar "001-0244"
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "#" + NombreArchivo
        tERR.Anotar "001-0245"
        NombreArchivo = Dir$
    Loop
    
    'avi
    tERR.Anotar "001-0246"
    NombreArchivo = Dir$(Carpeta + "\*.avi")
    tERR.Anotar "001-0247"
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        tERR.Anotar "001-0248"
        NewName = QuitarCaracter(NombreArchivo, ",")
        tERR.Anotar "001-0249"
        If NombreArchivo <> NewName Then
            tERR.Anotar "001-0250"
            FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
            tERR.Anotar "001-0251"
            'WriteTBRLog "Se corrigio el nombre de tema " + NombreArchivo + _
                " por " + NewName + " en la carpeta " + Carpeta, True
            tERR.Anotar "001-0252"
            NombreArchivo = NewName
        End If
        tERR.Anotar "001-0253"
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        tERR.Anotar "001-0254"
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "#" + NombreArchivo
        tERR.Anotar "001-0255"
        NombreArchivo = Dir$
    Loop
    
    'dat
    tERR.Anotar "001-0246"
    NombreArchivo = Dir$(Carpeta + "\*.dat")
    tERR.Anotar "001-0247"
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        tERR.Anotar "001-0248"
        NewName = QuitarCaracter(NombreArchivo, ",")
        tERR.Anotar "001-0249"
        If NombreArchivo <> NewName Then
            tERR.Anotar "001-0250"
            FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
            tERR.Anotar "001-0251"
            'WriteTBRLog "Se corrigio el nombre de tema " + NombreArchivo + _
                " por " + NewName + " en la carpeta " + Carpeta, True
            tERR.Anotar "001-0252"
            NombreArchivo = NewName
        End If
        tERR.Anotar "001-0253"
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        tERR.Anotar "001-0254"
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "#" + NombreArchivo
        tERR.Anotar "001-0255"
        NombreArchivo = Dir$
    Loop
    
    'wmv
    tERR.Anotar "001-0246"
    NombreArchivo = Dir$(Carpeta + "\*.wmv")
    tERR.Anotar "001-0247"
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        tERR.Anotar "001-0248"
        NewName = QuitarCaracter(NombreArchivo, ",")
        tERR.Anotar "001-0249"
        If NombreArchivo <> NewName Then
            tERR.Anotar "001-0250"
            FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
            tERR.Anotar "001-0251"
            'WriteTBRLog "Se corrigio el nombre de tema " + NombreArchivo + _
                " por " + NewName + " en la carpeta " + Carpeta, True
            tERR.Anotar "001-0252"
            NombreArchivo = NewName
        End If
        tERR.Anotar "001-0253"
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        tERR.Anotar "001-0254"
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "#" + NombreArchivo
        tERR.Anotar "001-0255"
        NombreArchivo = Dir$
    Loop
    
    tERR.Anotar "001-0256"
    ObtenerArchMM = TMPmatriz
    
    Exit Function
ErrObtMM:
    tERR.AppendLog tERR.ErrToTXT(Err), "Archivos.bas" + ".acpk4"
    Resume Next
    
End Function

Public Function QuitarCaracter(FileOrFolder As String, _
    CaracterToKill As String) As String
    'sacar en caracter de una cadena
    'lo uso para sacar las comas de los archivos mp3
    'o los puntos de los nombre de los discos
    tERR.Anotar "001-0257"
    Dim SeCambio As Boolean
    Dim TMPfOf 'temporario de file or folder
    tERR.Anotar "001-0258"
    TMPfOf = FileOrFolder
    Dim FindIn As Long
    Dim Parte1 As String, Parte2 As String
    tERR.Anotar "001-0259"
    SeCambio = False
    Do
        tERR.Anotar "001-0260"
        FindIn = InStr(TMPfOf, CaracterToKill)
        If FindIn > 0 Then
            tERR.Anotar "001-0261"
            SeCambio = True
            tERR.Anotar "001-0262"
            Parte1 = Mid(TMPfOf, 1, FindIn - 1)
            tERR.Anotar "001-0263"
            Parte2 = Mid(TMPfOf, FindIn + 1, Len(TMPfOf) - FindIn)
            tERR.Anotar "001-0264"
            TMPfOf = Parte1 + Parte2
        Else
            tERR.Anotar "001-0265"
            Exit Do
        End If
    Loop
    tERR.Anotar "001-0266"
    QuitarCaracter = TMPfOf
    
End Function
