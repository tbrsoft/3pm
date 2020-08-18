Attribute VB_Name = "Archivos"
'---------------------------------------------------------
' BIBLIOTECA DE RUTINAS DE PROGRAMACIÓN VB6 - (C) Francesco Balena
'
' (36+10) Rutinas del capítulo 05
'---------------------------------------------------------

Option Explicit

' API declaración (utilizada por la rutina EsperarPorProceso)
Private Declare Function EsperarUnicoObjeto Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilisegundos As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwAccess As _
    Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

' devuelve los atributos de un archivo en un formato legible
' esta rutina también funciona con archivos abiertos
' provoca un error si el archivo no existe

Function ObtAtribDescrip(nombrearch As String) As String
    Dim resultado As String, attr As Long
    attr = GetAttr(nombrearch)
    ' GetAttr también funciona con directorios
    If attr And vbDirectory Then resultado = resultado & " Directorio"
    If attr And vbReadOnly Then resultado = resultado & " Sólo lectura"
    If attr And vbHidden Then resultado = resultado & " Oculto"
    If attr And vbSystem Then resultado = resultado & " Sistema"
    If attr And vbArchive Then resultado = resultado & " Archivo"
    ' descarta el primer espacio
    ObtAtribDescrip = Mid$(resultado, 2)
End Function



Function ObtenerArchivos(Path As String, EXT As String) As String()
        ' proporciona un array de cadenas que almacenan todos los nombres de archivo que
        ' coinciden con una especificación de archivo dada y unos atributos de búsqueda.
        'devuelve path,nombrearchivo
        If Right(Path, 1) <> "\" Then Path = Path + "\"
        Dim resultado() As String
        Dim NombreArchivo As String, Contador As Long, ruta2 As String
        Const ALLOC_CHUNK = 50
        ReDim resultado(0 To ALLOC_CHUNK) As String
        NombreArchivo = Dir$(Path + EXT)
        Do While Len(NombreArchivo)
            Contador = Contador + 1
            If Contador > UBound(resultado) Then
                ' cambia el tamaño del array resultado, si es necesario
                ReDim Preserve resultado(0 To Contador + ALLOC_CHUNK) As String
            End If
            resultado(Contador) = Path + NombreArchivo + "," + NombreArchivo
            ' queda preparado para la siguiente iteración
            NombreArchivo = Dir$
        Loop
        
        ' devuelve el array resultado
        ReDim Preserve resultado(0 To Contador) As String
        ObtenerArchivos = resultado
End Function

' analiza la existencia de un archivo

Function ExisteArch(NombreArchivo As String) As Boolean
    On Error Resume Next
    ExisteArch = (Dir$(NombreArchivo) <> "")
End Function

' verificar si existe un directorio

Function ExisteDir(Ruta As String) As Boolean
    On Error Resume Next
    ExisteDir = (Dir$(Ruta & "\nul") <> "")
End Function

' Devuelve un array de cadenas que incluye todos los subdirectorios
' contenidos en una ruta

'tiene metido el mostrador de avance de proceso

Function ObtenerDir(Ruta As String) As String()
        frmProces.Show
        frmProces.pBar.Max = 150 'calculo que mas de estos discos no hay
        frmProces.pBar = 0
        frmProces.lblProces = "Iniciando busqueda"
        frmProces.lblProces.Refresh
        Dim ParaMatriz As String 'para generar cada elemento de la matriz
        Dim ContadorCarp As Long, CantMP3 As Long
        Dim resultado() As String
        Dim NombreDir As String, Contador As Long, ruta2 As String
        Const ALLOC_CHUNK = 50
        ReDim resultado(ALLOC_CHUNK) As String
        ' genera el nombre de ruta + barra invertida
        ruta2 = Ruta
        If Right$(ruta2, 1) <> "\" Then ruta2 = ruta2 & "\"
        NombreDir = Dir$(ruta2 & "*.*", vbDirectory)
        Do While Len(NombreDir)
            If NombreDir = "." Or NombreDir = ".." Then
                ' excluir las entradas "." y ".."
            ElseIf (GetAttr(ruta2 & NombreDir) And vbDirectory) = 0 Then
                ' este es un archivo normal
            Else
                ' es un directorio
                Contador = Contador + 1
                If Contador > UBound(resultado) Then
                    ' cambia el tamaño del array resultante, si
                    ' en necesario
                    ReDim Preserve resultado(Contador + ALLOC_CHUNK) As String
                End If
                
                frmProces.pBar = frmProces.pBar + 1
                'si me hacerco al max de pbar lo hago inalcanzable
                If frmProces.pBar > 140 Then frmProces.pBar.Max = frmProces.pBar.Max + 1
                frmProces.lblProces = NombreDir
                frmProces.lblProces.Refresh
                ContadorCarp = ContadorCarp + 1
                ParaMatriz = ruta2 & NombreDir + "," + NombreDir
                
                resultado(Contador) = ParaMatriz
            End If
            NombreDir = Dir$
            
        Loop
        TOTAL_DISCOS = ContadorCarp
        ' proporciona el array resultante
        ReDim Preserve resultado(Contador) As String
        ObtenerDir = resultado
End Function

'cuenta los archivos de determinada extension contenidos en una carpeta
Public Function ContarArch(ByVal Ruta As String, EXT As String, VerSubDIRs As Boolean) As Long
    Dim nombres() As String, i As Long, CONT As Long
    ' asegurarse de que existe una barra invertida inicial
    If Right(Ruta, 1) <> "\" Then Ruta = Ruta & "\"
    ' obtener la lista de archivos ejecutables
    
    nombres() = ObtenerArchivos(Ruta, EXT)
    'aqui esta la lista, por ahora no la uso
    'For i = 1 To UBound(nombres)
    '    lst.AddItem Ruta & nombres(i)
    'Next
    CONT = CONT + UBound(nombres)
    
    If VerSubDIRs Then
        ' obtener la lista de subdirectorios, incluyendo los ocultos
        ' y ejecutar recursivamente esta rutina en todos ellos.
        nombres() = ObtenerDir(Ruta)
        For i = 1 To UBound(nombres)
            ContarArch Ruta & nombres(i), EXT, True
        Next
    End If
End Function

' carga un archivo de texto en un control TextBox

Sub cargarArchivoEnTextBox(NombreArchivo As String, txt As TextBox)
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
    txt.Text = Texto
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
    Dim FSO As New Scripting.FileSystemObject
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

Function BuscarArchTexto(Ruta As String, buscar As String) As Variant()
    Dim FSO As New Scripting.FileSystemObject
    Dim fil As Scripting.File, ts As Scripting.TextStream
    Dim pos As Long, Contador As Long
    ReDim resultado(50) As Variant
    ' buscar for all the TXT files in the directory
    For Each fil In FSO.GetFolder(Ruta).Files
        If UCase$(FSO.GetExtensionName(fil.Path)) = "TXT" Then
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
                Contador = Contador + 1
                If Contador > UBound(resultado) Then
                    ReDim Preserve resultado(UBound(resultado) + 50) As Variant
                End If
                ' cada array resultado tiene tres elementos
                resultado(Contador) = Array(fil.Path, ts.Line, ts.Column)
                ' ahora podemos cerrar el TextStrem
                ts.Close
            End If
        End If
    Next
    ' cambia el tamaño del array resultado para indicar el número de
    ' coincidencas
    ReDim Preserve resultado(0 To Contador) As Variant
    BuscarArchTexto = resultado
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


