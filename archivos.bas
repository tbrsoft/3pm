Attribute VB_Name = "Archivos"
Option Explicit
Option Compare Text

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


' API declaración (utilizada por la rutina EsperarPorProceso)
Private Declare Function EsperarUnicoObjeto Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilisegundos As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwAccess As _
    Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Public Sub AbrirArchivo(Arch As String, FrmSolicita As Form)
    ShellExecute FrmSolicita.hWnd, vbNullString, Arch, vbNullString, vbNullString, vbMaximizedFocus
End Sub


Public Sub MostrarCursor(Mostrar As Boolean)
    Dim A As Long
    If Mostrar Then
        A = 0
        Do While A < 1
            A = ShowCursor(1) 'suma uno
        Loop
    Else
        A = 1
        Do While A >= 0
            A = ShowCursor(0) 'suma uno
        Loop
    End If
End Sub

' devuelve los atributos de un archivo en un formato legible
' esta rutina también funciona con archivos abiertos
' provoca un error si el archivo no existe

Function ObtAtribDescrip(nombrearch As String) As String
    Dim Resultado As String, attr As Long
    attr = GetAttr(nombrearch)
    ' GetAttr también funciona con directorios
    If attr And vbDirectory Then Resultado = Resultado & " Directorio"
    If attr And vbReadOnly Then Resultado = Resultado & " Sólo lectura"
    If attr And vbHidden Then Resultado = Resultado & " Oculto"
    If attr And vbSystem Then Resultado = Resultado & " Sistema"
    If attr And vbArchive Then Resultado = Resultado & " Archivo"
    ' descarta el primer espacio
    ObtAtribDescrip = Mid$(Resultado, 2)
End Function



Function ObtenerArchivos(path As String, EXT As String) As String()
        ' proporciona un array de cadenas que almacenan todos los nombres de archivo que
        ' coinciden con una especificación de archivo dada y unos atributos de búsqueda.
        'devuelve path,nombrearchivo
        If Right(path, 1) <> "\" Then path = path + "\"
        Dim Resultado() As String
        Dim NombreArchivo As String, ContadorArch As Long, Ruta2 As String
        Const ALLOC_CHUNK = 50
        ReDim Resultado(0 To ALLOC_CHUNK) As String
        NombreArchivo = Dir$(path + EXT)
        Do While Len(NombreArchivo)
            ContadorArch = ContadorArch + 1
            If ContadorArch > UBound(Resultado) Then
                ' cambia el tamaño del array resultado, si es necesario
                ReDim Preserve Resultado(0 To ContadorArch + ALLOC_CHUNK) As String
            End If
            Resultado(ContadorArch) = path + NombreArchivo + "," + NombreArchivo
            ' queda preparado para la siguiente iteración
            NombreArchivo = Dir$
        Loop
        'devuelve el array resultado
        ReDim Preserve Resultado(0 To ContadorArch) As String
        ObtenerArchivos = Resultado
End Function

' analiza la existencia de un archivo

Function ExisteArch(NombreArchivo As String) As Boolean
    On Error Resume Next
    ExisteArch = (Dir$(NombreArchivo) <> "")
End Function

' verificar si existe un directorio

Function ExisteDir(ruta As String) As Boolean
    On Error Resume Next
    ExisteDir = (Dir$(ruta & "\nul") <> "")
End Function

' Devuelve un array de cadenas que incluye todos los subdirectorios
' contenidos en una ruta

'corrige ademas los puntos que pueda tener, los saca

'tiene metido el mostrador de avance de proceso

Function ObtenerDir(ruta As String) As String()
        Dim NewName As String 'nuevo nombre si hay que corregir puntos metidos en el nombre de la carpeta
        Dim MaxPBAR As Long
        MaxPBAR = frmINI.PBar.Width
        frmINI.PBar.Width = 0
        frmINI.lblProces = "Iniciando busqueda"
        frmINI.lblProces.Refresh
        Dim ParaMatriz As String 'para generar cada elemento de la matriz
        Dim ContadorCarp As Long, CantMP3 As Long
        Dim Resultado() As String
        Dim NombreDir As String, ContadorArch As Long, Ruta2 As String
        Const ALLOC_CHUNK = 50
        ReDim Resultado(ALLOC_CHUNK) As String
        ' genera el nombre de ruta + barra invertida
        Ruta2 = ruta
        If Right$(Ruta2, 1) <> "\" Then Ruta2 = Ruta2 & "\"
        NombreDir = Dir$(Ruta2 & "*.*", vbDirectory)
        Do While Len(NombreDir)
            If NombreDir = "." Or NombreDir = ".." Then
                ' excluir las entradas "." y ".."
            ElseIf (GetAttr(Ruta2 & NombreDir) And vbDirectory) = 0 Then
                ' este es un archivo normal
            Else
                ' es un directorio
                If RankToPeople = False And NombreDir = "01- Los mas escuchados" Then
                    'pasar al que sigue
                    GoTo NextCarp
                End If
                ContadorArch = ContadorArch + 1
                                
                frmINI.lblINI = "Contando Discos: " + Trim(Str(ContadorArch))
                frmINI.lblINI.Refresh
                If ContadorArch > UBound(Resultado) Then
                    ' cambia el tamaño del array resultante, si
                    ' en necesario
                    ReDim Preserve Resultado(ContadorArch + ALLOC_CHUNK) As String
                End If
                
                frmINI.PBar.Width = frmINI.PBar.Width + 100
                'si me hacerco al max de pbar lo hago inalcanzable
                frmINI.lblProces = NombreDir
                frmINI.lblProces.Refresh
                ContadorCarp = ContadorCarp + 1
                'corregir el nombre del tema
                NewName = QuitarCaracter(NombreDir, ".")
                If NombreDir <> NewName Then
                    FSO.MoveFolder Ruta2 + NombreDir, Ruta2 + NewName
                    WriteTBRLog "Se corrigio el nombre de la carpeta " + NombreDir + _
                        " por " + NewName, True
                    NombreDir = NewName
                End If
                ParaMatriz = Ruta2 & NombreDir + "," + NombreDir
                
                Resultado(ContadorArch) = ParaMatriz
NextCarp:
                '=============================================================================
                '=============================================================================
                Dim MD
                MD = 12
                If TypeVersion = "DEMO" And ContadorArch > MD Then
                    'limite de discos
                    MsgBox "Esta es una version demo y no se pueden cargar más " + _
                    "de " + Trim(Str(MD)) + " discos." + vbCrLf + _
                    "Para conseguir la versión sin límite de discos y con el manual " + _
                    "completo envie un e-mail a tbrsoft@hotmail.com o a " + _
                    "avazquez@cpcipc.org. Solo se cargaran los " + Trim(Str(MD)) + " primeros discos"
                    GoTo Solo12
                End If
                If TypeVersion = "DEMO2" And ContadorArch > MD Then
                    'limite de discos
                    MsgBox "Esta es una version demo y no se pueden cargar más " + _
                    "de " + Trim(Str(MD)) + " discos." + vbCrLf + _
                    "Para conseguir la versión sin límite de discos y con el manual " + _
                    "completo envie un e-mail a tbrsoft@hotmail.com o a " + _
                    "avazquez@cpcipc.org. Solo se cargaran los " + Trim(Str(MD)) + " primeros discos"
                    GoTo Solo12
                End If
                '=============================================================================
                '=============================================================================
            End If
            NombreDir = Dir$
            
        Loop
Solo12: 'solo los 12 primeros
        frmINI.PBar.Width = MaxPBAR
        TOTAL_DISCOS = ContadorCarp
        ' proporciona el array resultante
        ReDim Preserve Resultado(ContadorArch) As String
        
        'tomar la matriz (con valores separador) y ordenala en base a la columna indicada.
        'en este caso el separador es "," y la columna es 0.
        'seria los mismo que tomara 1 ya que todos tienen el mismo path
        Dim MinSTR As String 'comparacoin de cadenas. Empiezo con el máximo
        Dim ubicMIN As Long 'indice en la matriz del menor encontrado cada vuelta
        MinSTR = "zzzzzzzzzzzzzzzz"
        Dim c As Long, mtx As Long, ValComp As String
        c = 0 'cantidad de minimos encontrados
        Dim Ordenados() As Long 'matriz con los indices ordenados
        Do
            For mtx = LBound(Resultado) + 1 To UBound(Resultado)
                ValComp = txtInLista(Resultado(mtx), 0, ",")
                If ValComp < MinSTR Then
                    MinSTR = ValComp
                    ubicMIN = mtx
                End If
            Next
            Resultado(ubicMIN) = "zzzzzzzzzz," + Resultado(ubicMIN)
            c = c + 1
            ReDim Preserve Ordenados(c)
            Ordenados(c) = ubicMIN
            If c >= UBound(Resultado) Then Exit Do
            MinSTR = "zzzzzzzzzz"
        Loop
        'cargar todos y sacar la primera columna de las zetas
        Dim MTXsort() As String
        'ver que haya alguna carpeta
        If ContadorCarp < 2 Then
            'VER SI HAY UN DISCO Y NO ES EL RANKING
            If RankToPeople = False And ContadorCarp = 1 Then GoTo EntreAlPedo
            MsgBox "NO HAY DISCOS PARA MOSTRAR." + vbCrLf + _
            "Una vez iniciado el sistema presione la tecla " + _
            "'C' para ingresar a la configuracion y utilize el " + _
            "asistente para cargar multimedia al sistema."
        End If
EntreAlPedo:
        Dim nTAPAcd As Integer
        nTAPAcd = 0
        frmINI.PBar.Width = 0
        
        For mtx = LBound(Resultado) + 1 To UBound(Resultado)
            ReDim Preserve MTXsort(mtx)
            Dim CarpFull As String, NameCarp As String
            CarpFull = txtInLista(Resultado(Ordenados(mtx)), 1, ",")
            NameCarp = txtInLista(Resultado(Ordenados(mtx)), 2, ",")
            MTXsort(mtx) = CarpFull + "," + NameCarp
            'cargar todas las imagenes si asi esta configurado
            If CargarIMGinicio Then
                UbicDiscoActual = txtInLista(Resultado(Ordenados(mtx)), 1, ",")
                
                'caragar las imágenes en diferentes IMGs para que no se cargen despues
                Dim ArchTapa As String
                ArchTapa = UbicDiscoActual + "\tapa.jpg"
                'arranca con 5 ya cargados
                If nTAPAcd > 0 Then
                    Load frmIndex.L(nTAPAcd)
                    frmIndex.L(nTAPAcd).Top = frmIndex.L(nTAPAcd - 1).Top + frmIndex.L(nTAPAcd - 1).Height
                    frmIndex.L(nTAPAcd).Visible = True
                End If
                frmIndex.L(nTAPAcd) = NameCarp
                frmINI.lblProces = NameCarp
                frmINI.lblProces.Refresh
                'INICIO RAPIDO fastini
                'si hay, mostrar la tapa
                If FASTini = False And FSO.FileExists(ArchTapa) Then frmINI.TapaCD.Picture = LoadPicture(ArchTapa)
                
                frmINI.lblINI = "Ordenando Discos: " + Trim(Str(mtx))
                frmINI.lblINI.Refresh
                frmINI.PBar.Width = frmINI.PBar.Width + 100
                frmINI.PBar.Refresh
                
                If nTAPAcd > ((TapasMostradasH * TapasMostradasV) - 1) Then
                    Load frmIndex.TapaCD(nTAPAcd)
                    frmIndex.TapaCD(nTAPAcd).Left = frmIndex.TapaCD(nTAPAcd - ((TapasMostradasH * TapasMostradasV))).Left
                    frmIndex.TapaCD(nTAPAcd).Top = frmIndex.TapaCD(nTAPAcd - ((TapasMostradasH * TapasMostradasV))).Top
                End If
            
                If FSO.FileExists(ArchTapa) Then
                    frmIndex.TapaCD(nTAPAcd).Picture = LoadPicture(ArchTapa)
                Else
                    frmIndex.TapaCD(nTAPAcd).Picture = LoadPicture(AP + "tapa.jpg")
                End If
                nTAPAcd = nTAPAcd + 1
            End If
        Next
        frmINI.lblINI = "Proceso terminado, cargando 3PM..."
        frmINI.lblINI.Refresh
        frmINI.PBar.Width = MaxPBAR
        ObtenerDir = MTXsort
End Function

'cuenta los archivos de determinada extension contenidos en una carpeta
Public Function ContarArch(ByVal ruta As String, EXT As String, VerSubDIRs As Boolean) As Long
    Dim nombres() As String, i As Long, CONT As Long
    ' asegurarse de que existe una barra invertida inicial
    If Right(ruta, 1) <> "\" Then ruta = ruta & "\"
    ' obtener la lista de archivos ejecutables
    
    nombres() = ObtenerArchivos(ruta, EXT)
    'aqui esta la lista, por ahora no la uso
    'For i = 1 To UBound(nombres)
    '    lst.AddItem Ruta & nombres(i)
    'Next
    CONT = CONT + UBound(nombres)
    
    If VerSubDIRs Then
        ' obtener la lista de subdirectorios, incluyendo los ocultos
        ' y ejecutar recursivamente esta rutina en todos ellos.
        nombres() = ObtenerDir(ruta)
        For i = 1 To UBound(nombres)
            ContarArch ruta & nombres(i), EXT, True
        Next
    End If
End Function

' carga un archivo de texto en un control TextBox

Sub cargarArchivoEnTextBox(NombreArchivo As String, TXT As TextBox)
    Dim numlib As Integer, isOpen As Boolean
    Dim lineatexto As String, Texto As String
    On Error GoTo Manejador_Error
    ' obtiene el siguiente número libre de archivo
    numlib = FreeFile()
    Open NombreArchivo For Input As #numlib
    ' si el flujo llega hasta aquí, se habrán abierto los archivos
    ' sin que se produzca ningún error
    isOpen = True
    Do Until EOF(numlib)
        Line Input #numlib, lineatexto
        Texto = Texto & lineatexto & vbCrLf
    Loop
    ' cargar en el cuadro de texto
    TXT.Text = Texto
    ' se cae intencionadamente en el manejador de error para
    ' cerrar el archivo
Manejador_Error:
    ' se provoca un error(si es que hay alguno), pero primero
    ' se cierra el archivo
    If isOpen Then Close #numlib
    If Err Then Err.Raise Err.Number, , Err.Description
End Sub

' proporciona el contenido completo de un archivo de texto en una cadena

Function LeerContenidoArchTexto(NombreArchivo As String) As String
    Dim numlib As Integer, isOpen As Boolean
    On Error GoTo Manejador_Error
    ' obtiene el siguiente número libre de archivo
    numlib = FreeFile()
    Open NombreArchivo For Input As #numlib
    ' si el flujo llega hasta aquí, se habrán abierto los archivos
    ' sin que se produzca ningún error
    isOpen = True
    ' leer todo el contenido en una única operación
    LeerContenidoArchTexto = Input(LOF(numlib), numlib)
    ' se cae intencionadamente en el manejador de error para
    ' cerrar el archivo
Manejador_Error:
    ' se provoca un error(si es que hay alguno), pero primero
    ' se cierra el archivo
    If isOpen Then Close #numlib
    If Err Then Err.Raise Err.Number, , Err.Description
End Function

' escribe el contenido de una cadena en un archivo, opcionalmente
' en modo Append

Sub EscribirContenidoArchTexto(Texto As String, NombreArchivo As String, _
    Optional ModoAppend As Boolean)
        Dim numlib As Integer, isOpen As Boolean
        On Error GoTo Manejador_Error
        ' obtiene el siguiente número libre de archivo
        numlib = FreeFile()
        If ModoAppend Then
            Open NombreArchivo For Append As #numlib
        Else
            Open NombreArchivo For Output As #numlib
        End If
        ' si el flujo de ejecución llega hasta aquí es que el archivo
        ' se ha abierto correctamente
        isOpen = True
        ' imprime al archivo en una única operación
        Print #numlib, Texto
    ' se cae intencionadamente en el manejador de error para
    ' cerrar el archivo
Manejador_Error:
    ' se provoca un error(si es que hay alguno), pero primero
    ' se cierra el archivo
        If isOpen Then Close #numlib
        If Err Then Err.Raise Err.Number, , Err.Description
End Sub

' proporciona el contenido de un archivo de texto
' como un array de cadenas.
' NOTA: requiere el empleo de la rutina LeerContenidoArchTexto

Function ObtenLineasArchTexto(NombreArchivo As String, Optional DropEmpty As Boolean, _
    Optional Limite As Variant) As String()
        Dim ArchTexto As String, elementos() As String, i As Long
        ' lee el contenido del archivo, salir si hay un error
        ArchTexto = LeerContenidoArchTexto(NombreArchivo)
        ' esto es necesario porque Split() sólo acepta delimitadores de 1 carácter
        ArchTexto = Replace(ArchTexto, vbCrLf, vbCr)
        ' divide al archivo en líneas individuales de texto
        elementos() = Split(ArchTexto, vbCr, Limite)
        ' proporciona líneas sencillas, si se solicita
        If DropEmpty Then
            ' llena las líneas vacías con algo que otros elementos
            ' no contienen
            For i = 0 To UBound(elementos)
                If Len(elementos(i)) = 0 Then elementos(i) = vbCrLf
            Next
            ' utiliza la función Filter() para soltar rápidamente
            ' las líneas vacías
            elementos() = Filter(elementos(), vbCrLf, False)
        End If
        ObtenLineasArchTexto = elementos()
End Function

' proporciona el contenido de un archivo de texto delimitado como un
' array de arrays de cadenas.
' NOTA: requiere el empleo de las rutinas LeerContenidoArchTexto y
' ObtenLineasArchTexto

Function ImportarArchDelimitado(NombreArchivo As String, _
    Optional delimitador As String = vbTab) As Variant()
        Dim lineas() As String, i As Long
        ' obtiene todas las lineas contenidas en el archivo,
        ' ignorando las líneas en blanco
        lineas() = ObtenLineasArchTexto(NombreArchivo, True)
        ' crea un array de cadena por cada línea de texto
        ' y lo almacena en un elemento Variant
        ReDim valores(0 To UBound(lineas)) As Variant
        For i = 0 To UBound(lineas)
            valores(i) = Split(lineas(i), delimitador, -1)
        Next
        ImportarArchDelimitado = valores()
End Function

' escribir el contenido de un array de arrays de cadenas a un
' archivo de texto deliminado.
' NOTA: necesita la rutina EscribirContenidoArchTexto

Sub ExportarArchDelimitado(valores() As Variant, NombreArchivo As String, _
    Optional delimitador As String = vbTab)
        Dim i As Long, j As Long, ArchTexto As String
        ' reconstruye las líneas individuales de texto del archivo
        ReDim lineas(0 To UBound(valores)) As String
        For i = 0 To UBound(valores)
            lineas(i) = Join(valores(i), delimitador)
        Next
        ' introduce CRLFs entre registros
        ArchTexto = Replace(Join(lineas, vbCr), vbCr, vbCrLf)
        EscribirContenidoArchTexto ArchTexto, NombreArchivo
End Sub

' duplica el árbol de directorios sin copiar los archivos

' llamar a esta rutina para iniciar el proceso de copia
' NOTA: la carpeta destino se creará en caso necesario
'       utiliza el procedimiento Private Sub DuplicarDirArbolSub

Sub DuplicarDirArbol(rutaOrigen As String, rutaDest As String)
    Dim CarpOrigen As Scripting.Folder, CarpDest As Scripting.Folder
    ' la carpeta origen debe existir
    Set CarpOrigen = FSO.GetFolder(rutaOrigen)
    ' la carpeta destino se creará en caso necesario
    If FSO.FolderExists(rutaDest) Then
        Set CarpDest = FSO.GetFolder(rutaDest)
    Else
        Set CarpDest = FSO.CreateFolder(rutaDest)
    End If
    ' saltar a la rutina recursiva para realizar el trabajo real
    DuplicarDirArbolSub CarpOrigen, CarpDest
End Sub

' Procedimiento recursivo privado utilizado por DuplicarDirArbol

Private Sub DuplicarDirArbolSub(origen As Folder, destino As Folder)
    Dim CarpOrigen As Scripting.Folder, CarpDest As Scripting.Folder
    For Each CarpOrigen In origen.SubFolders
        ' copiar esta subcarpeta en la carpeta destino
        Set CarpDest = destino.SubFolders.Add(CarpOrigen.Name)
        ' repetir el proceso recursivamente para todas las
        ' subcarpetas de la carpeta considerada
        DuplicarDirArbolSub CarpOrigen, CarpDest
    Next
End Sub

' Busca una cadena en todos los archivos TXT contenidos en un directorio.

' Por cada archivo localizado devuelve un elemento Variant que contiene
' un array de tres elementos: el nombre del archivo, la línea
' y el número de columna.
' NOTA: las búsquedas no distinguen el uso de mayúsculas y minúsculas

Function BuscarArchTexto(ruta As String, buscar As String) As Variant()
    Dim fil As Scripting.File, ts As Scripting.TextStream
    Dim pos As Long, ContadorArch As Long
    ReDim Resultado(50) As Variant
    ' buscar for all the TXT files in the directory
    For Each fil In FSO.GetFolder(ruta).Files
        If UCase$(FSO.GetExtensionName(fil.path)) = "TXT" Then
            ' obtener el objeto TextStream correspondiente
            Set ts = fil.OpenAsTextStream(ForReading)
            ' leer su contenido, buscar la cadena y cerrarlo
            pos = InStr(1, ts.ReadAll, buscar, vbTextCompare)
            ts.Close
            If pos > 0 Then
                ' si se encuentra la cadena, reabre el archivo
                ' para determinar su posición en forma de (línea, columna)
                Set ts = fil.OpenAsTextStream(ForReading)
                ' salta todos los caracteres precedentes para saber dónde se
                ' encuentra la cadena
                ts.Skip pos - 1
                ' llena el array resultado, hace sitio en caso necesario
                ContadorArch = ContadorArch + 1
                If ContadorArch > UBound(Resultado) Then
                    ReDim Preserve Resultado(UBound(Resultado) + 50) As Variant
                End If
                ' cada array resultado tiene tres elementos
                Resultado(ContadorArch) = Array(fil.path, ts.Line, ts.Column)
                ' ahora podemos cerrar el TextStrem
                ts.Close
            End If
        End If
    Next
    ' cambia el tamaño del array resultado para indicar el número de
    ' coincidencas
    ReDim Preserve Resultado(0 To ContadorArch) As Variant
    BuscarArchTexto = Resultado
End Function

' espera un número de milisegundos y devuelve el estado de ejecución de un
' proceso; si se omite el argumento, espera hasta que el proceso finalice.

Function EsperarPorProceso(taskId As Long, Optional msecs As Long = -1) _
    As Boolean
        Dim procHandle As Long
        ' obtiene el manejador del proceso
        procHandle = OpenProcess(&H100000, True, taskId)
        ' verifica su estado señalado, lo devuelve al que hizo la llamada
        EsperarPorProceso = EsperarUnicoObjeto(procHandle, msecs) <> -1
        ' cierra el gestor
        CloseHandle procHandle
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

Public Sub EscribirArch1Linea(Arch As String, TXT As String)
    Set TE = FSO.CreateTextFile(Arch, True)
    TE.WriteLine TXT
    TE.Close
End Sub

Public Function ObtenerArchMM(Carpeta As String) As String()
    'devuelve "Carpeta + NombreArchivo + "," + NombreArchivo"
    'devuelve PathFull,SoloNombre

    'ADEMÁS DEBO ASEGURARME QUE NO HAYA COMAS EN LOS NOMBRES
    
    If Right(Carpeta, 1) <> "\" Then Carpeta = Carpeta + "\"
    Dim TMPmatriz() As String
    ReDim Preserve TMPmatriz(0)
    'mp3
    Dim NombreArchivo As String, ContadorArch As Long, NewName As String
    NombreArchivo = Dir$(Carpeta + "*.mp3")
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        NewName = QuitarCaracter(NombreArchivo, ",")
        If NombreArchivo <> NewName Then
            'no se puede corregir si es un CD. Solo corrige si es disco duro
            'esta funcion se usa para leer CDs debo prevenir
            If FSO.Drives(Left(Carpeta, 1)).DriveType = Fixed Then
                FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
                WriteTBRLog "Se corrigio el nombre de tema " + NombreArchivo + _
                    " por " + NewName + " en la carpeta " + Carpeta, True
                NombreArchivo = NewName
            End If
        End If
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "," + NombreArchivo
        NombreArchivo = Dir$
    Loop
    
    'mpg
    NombreArchivo = Dir$(Carpeta + "\*.mpg")
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        NewName = QuitarCaracter(NombreArchivo, ",")
        If NombreArchivo <> NewName Then
            FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
            WriteTBRLog "Se corrigio el nombre de tema " + NombreArchivo + _
                " por " + NewName + " en la carpeta " + Carpeta, True
            NombreArchivo = NewName
        End If
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "," + NombreArchivo
        NombreArchivo = Dir$
    Loop
    
    'mpeg
    NombreArchivo = Dir$(Carpeta + "\*.mpeg")
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        NewName = QuitarCaracter(NombreArchivo, ",")
        If NombreArchivo <> NewName Then
            FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
            WriteTBRLog "Se corrigio el nombre de tema " + NombreArchivo + _
                " por " + NewName + " en la carpeta " + Carpeta, True
            NombreArchivo = NewName
        End If
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "," + NombreArchivo
        NombreArchivo = Dir$
    Loop
    
    'avi
    NombreArchivo = Dir$(Carpeta + "\*.avi")
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        NewName = QuitarCaracter(NombreArchivo, ",")
        If NombreArchivo <> NewName Then
            FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
            WriteTBRLog "Se corrigio el nombre de tema " + NombreArchivo + _
                " por " + NewName + " en la carpeta " + Carpeta, True
            NombreArchivo = NewName
        End If
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "," + NombreArchivo
        NombreArchivo = Dir$
    Loop
    ObtenerArchMM = TMPmatriz
End Function

Public Function QuitarCaracter(FileOrFolder As String, _
    CaracterToKill As String) As String
    'sacar en caracter de una cadena
    'lo uso para sacar las comas de los archivos mp3
    'o los puntos de los nombre de los discos
    Dim SeCambio As Boolean
    Dim TMPfOf 'temporario de file or folder
    TMPfOf = FileOrFolder
    Dim FindIn As Long
    Dim Parte1 As String, Parte2 As String
    SeCambio = False
    Do
        FindIn = InStr(TMPfOf, CaracterToKill)
        If FindIn > 0 Then
            SeCambio = True
            Parte1 = Mid(TMPfOf, 1, FindIn - 1)
            Parte2 = Mid(TMPfOf, FindIn + 1, Len(TMPfOf) - FindIn)
            TMPfOf = Parte1 + Parte2
        Else
            Exit Do
        End If
    Loop
    QuitarCaracter = TMPfOf
    
End Function
