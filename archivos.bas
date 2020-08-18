Attribute VB_Name = "Archivos"
Option Explicit
Option Compare Text

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


' API declaraci�n (utilizada por la rutina EsperarPorProceso)
Private Declare Function EsperarUnicoObjeto Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilisegundos As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwAccess As _
    Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Public Sub AbrirArchivo(Arch As String, FrmSolicita As Form)
    LineaError = "001-0001"
    ShellExecute FrmSolicita.hWnd, vbNullString, Arch, vbNullString, vbNullString, vbMaximizedFocus
End Sub


Public Sub MostrarCursor(Mostrar As Boolean)
    LineaError = "001-0002"
    Dim A As Long
    If Mostrar Then
        A = 0
        Do While A < 1
            LineaError = "001-0003"
            A = ShowCursor(1) 'suma uno
        Loop
    Else
        LineaError = "001-0004"
        A = 1
        Do While A >= 0
            LineaError = "001-0005"
            A = ShowCursor(0) 'suma uno
        Loop
    End If
End Sub

' devuelve los atributos de un archivo en un formato legible
' esta rutina tambi�n funciona con archivos abiertos
' provoca un error si el archivo no existe

Function ObtAtribDescrip(nombrearch As String) As String
    LineaError = "001-0006"
    Dim Resultado As String, attr As Long
    LineaError = "001-0007"
    attr = GetAttr(nombrearch)
    ' GetAttr tambi�n funciona con directorios
    LineaError = "001-0008"
    If attr And vbDirectory Then Resultado = Resultado & " Directorio"
    LineaError = "001-0009"
    If attr And vbReadOnly Then Resultado = Resultado & " S�lo lectura"
    LineaError = "001-0010"
    If attr And vbHidden Then Resultado = Resultado & " Oculto"
    LineaError = "001-0011"
    If attr And vbSystem Then Resultado = Resultado & " Sistema"
    LineaError = "001-0012"
    If attr And vbArchive Then Resultado = Resultado & " Archivo"
    ' descarta el primer espacio
    LineaError = "001-0013"
    ObtAtribDescrip = Mid$(Resultado, 2)
End Function



Function ObtenerArchivos(path As String, EXT As String) As String()
        ' proporciona un array de cadenas que almacenan todos los nombres de archivo que
        ' coinciden con una especificaci�n de archivo dada y unos atributos de b�squeda.
        'devuelve path,nombrearchivo
        LineaError = "001-0014"
        If Right(path, 1) <> "\" Then path = path + "\"
        LineaError = "001-0015"
        Dim Resultado() As String
        Dim NombreArchivo As String, ContadorArch As Long, Ruta2 As String
        Const ALLOC_CHUNK = 50
        LineaError = "001-0016"
        ReDim Resultado(0 To ALLOC_CHUNK) As String
        LineaError = "001-0017"
        NombreArchivo = Dir$(path + EXT)
        LineaError = "001-0018"
        Do While Len(NombreArchivo)
            LineaError = "001-0019"
            ContadorArch = ContadorArch + 1
            LineaError = "001-0020"
            If ContadorArch > UBound(Resultado) Then
                ' cambia el tama�o del array resultado, si es necesario
                LineaError = "001-0021"
                ReDim Preserve Resultado(0 To ContadorArch + ALLOC_CHUNK) As String
            End If
            LineaError = "001-0022"
            Resultado(ContadorArch) = path + NombreArchivo + "," + NombreArchivo
            ' queda preparado para la siguiente iteraci�n
            LineaError = "001-0023"
            NombreArchivo = Dir$
        Loop
        'devuelve el array resultado
        LineaError = "001-0024"
        ReDim Preserve Resultado(0 To ContadorArch) As String
        LineaError = "001-0025"
        ObtenerArchivos = Resultado
End Function

' analiza la existencia de un archivo
Function ExisteArch(NombreArchivo As String) As Boolean
    LineaError = "001-0026"
    On Error Resume Next
    LineaError = "001-0027"
    ExisteArch = (Dir$(NombreArchivo) <> "")
End Function

' verificar si existe un directorio

Function ExisteDir(ruta As String) As Boolean
    LineaError = "001-0028"
    On Error Resume Next
    ExisteDir = (Dir$(ruta & "\nul") <> "")
End Function

' Devuelve un array de cadenas que incluye todos los subdirectorios
' contenidos en una ruta

'corrige ademas los puntos que pueda tener, los saca

'tiene metido el mostrador de avance de proceso

Function ObtenerDir(ruta As String) As String()
    LineaError = "001-0029"
        Dim NewName As String 'nuevo nombre si hay que corregir puntos metidos en el nombre de la carpeta
        Dim MaxPBAR As Long
        LineaError = "001-0030"
        MaxPBAR = frmINI.PBar.Width
        LineaError = "001-0031"
        frmINI.PBar.Width = 0
        LineaError = "001-0032"
        frmINI.lblProces = "Iniciando busqueda"
        LineaError = "001-0033"
        frmINI.lblProces.Refresh
        LineaError = "001-0034"
        Dim ParaMatriz As String 'para generar cada elemento de la matriz
        Dim ContadorCarp As Long, CantMP3 As Long
        Dim Resultado() As String
        Dim NombreDir As String, ContadorArch As Long, Ruta2 As String
        LineaError = "001-0035"
        Const ALLOC_CHUNK = 50
        LineaError = "001-0036"
        ReDim Resultado(ALLOC_CHUNK) As String
        LineaError = "001-0037"
        ' genera el nombre de ruta + barra invertida
        Ruta2 = ruta
        LineaError = "001-0038"
        If Right$(Ruta2, 1) <> "\" Then Ruta2 = Ruta2 & "\"
        LineaError = "001-0039"
        NombreDir = Dir$(Ruta2 & "*.*", vbDirectory)
        LineaError = "001-0040"
        Do While Len(NombreDir)
            LineaError = "001-0041"
            If NombreDir = "." Or NombreDir = ".." Then
                ' excluir las entradas "." y ".."
                LineaError = "001-0042"
            ElseIf (GetAttr(Ruta2 & NombreDir) And vbDirectory) = 0 Then
                ' este es un archivo normal
                LineaError = "001-0043"
            Else
                ' es un directorio
                LineaError = "001-0044"
                If RankToPeople = False And NombreDir = "01- Los mas escuchados" Then
                    'pasar al que sigue
                    LineaError = "001-0045"
                    GoTo NextCarp
                End If
                LineaError = "001-0046"
                ContadorArch = ContadorArch + 1
                
                frmINI.lblINI = "Contando Discos: " + Trim(Str(ContadorArch))
                LineaError = "001-0047"
                frmINI.lblINI.Refresh
                LineaError = "001-0048"
                If ContadorArch > UBound(Resultado) Then
                    ' cambia el tama�o del array resultante, si
                    ' en necesario
                    LineaError = "001-0049"
                    ReDim Preserve Resultado(ContadorArch + ALLOC_CHUNK) As String
                End If
                LineaError = "001-0050"
                frmINI.PBar.Width = frmINI.PBar.Width + 100
                'si me hacerco al max de pbar lo hago inalcanzable
                LineaError = "001-0051"
                frmINI.lblProces = NombreDir
                LineaError = "001-0052"
                frmINI.lblProces.Refresh
                LineaError = "001-0053"
                ContadorCarp = ContadorCarp + 1
                'corregir el nombre del tema
                NewName = QuitarCaracter(NombreDir, ".")
                LineaError = "001-0054"
                If NombreDir <> NewName Then
                    LineaError = "001-0055"
                    FSO.MoveFolder Ruta2 + NombreDir, Ruta2 + NewName
                    LineaError = "001-0056"
                    WriteTBRLog "Se corrigio el nombre de la carpeta " + NombreDir + _
                        " por " + NewName, True
                    LineaError = "001-0057"
                    NombreDir = NewName
                End If
                LineaError = "001-0058"
                ParaMatriz = Ruta2 & NombreDir + "," + NombreDir
                LineaError = "001-0059"
                Resultado(ContadorArch) = ParaMatriz
NextCarp:
                '=============================================================================
                '=============================================================================
                Dim MD
                MD = 12
                LineaError = "001-0060"
                If K.LICENCIA = aSinCargar And ContadorArch > MD Then
                    'limite de discos
                    LineaError = "001-0061"
                    MsgBox "Esta es una version demo y no se pueden cargar m�s " + _
                    "de " + Trim(Str(MD)) + " discos." + vbCrLf + _
                    "Para conseguir la versi�n sin l�mite de discos y con el manual " + _
                    "completo envie un e-mail a tbrsoft@hotmail.com o a " + _
                    "tbrsoft@cpcipc.org. Solo se cargaran los " + Trim(Str(MD)) + " primeros discos"
                    LineaError = "001-0062"
                    GoTo Solo12
                End If
                LineaError = "001-0063"
                If K.LICENCIA = CGratuita And ContadorArch > MD Then
                    'limite de discos
                    LineaError = "001-0064"
                    MsgBox "Esta es una version demo y no se pueden cargar m�s " + _
                    "de " + Trim(Str(MD)) + " discos." + vbCrLf + _
                    "Para conseguir la versi�n sin l�mite de discos y con el manual " + _
                    "completo envie un e-mail a tbrsoft@hotmail.com o a " + _
                    "tbrsoft@cpcipc.org. Solo se cargaran los " + Trim(Str(MD)) + " primeros discos"
                    LineaError = "001-0065"
                    GoTo Solo12
                End If
                '=============================================================================
                '=============================================================================
            End If
            LineaError = "001-0066"
            NombreDir = Dir$
            
        Loop
Solo12: 'solo los 12 primeros
        LineaError = "001-0067"
        frmINI.PBar.Width = MaxPBAR
        LineaError = "001-0068"
        TOTAL_DISCOS = ContadorCarp
        ' proporciona el array resultante
        LineaError = "001-0069"
        ReDim Preserve Resultado(ContadorArch) As String
        
        'tomar la matriz (con valores separador) y ordenala en base a la columna indicada.
        'en este caso el separador es "," y la columna es 0.
        'seria los mismo que tomara 1 ya que todos tienen el mismo path
        LineaError = "001-0070"
        Dim MinSTR As String 'comparacoin de cadenas. Empiezo con el m�ximo
        Dim ubicMIN As Long 'indice en la matriz del menor encontrado cada vuelta
        LineaError = "001-0071"
        MinSTR = "zzzzzzzzzzzzzzzz"
        LineaError = "001-0072"
        Dim C As Long, mtx As Long, ValComp As String
        C = 0 'cantidad de minimos encontrados
        LineaError = "001-0073"
        Dim Ordenados() As Long 'matriz con los indices ordenados
        Do
            LineaError = "001-0074"
            For mtx = LBound(Resultado) + 1 To UBound(Resultado)
                LineaError = "001-0075"
                ValComp = txtInLista(Resultado(mtx), 0, ",")
                LineaError = "001-0076"
                If ValComp < MinSTR Then
                    LineaError = "001-0077"
                    MinSTR = ValComp
                    LineaError = "001-0078"
                    ubicMIN = mtx
                End If
            Next
            LineaError = "001-0079"
            Resultado(ubicMIN) = "zzzzzzzzzz," + Resultado(ubicMIN)
            C = C + 1
            ReDim Preserve Ordenados(C)
            LineaError = "001-0080"
            Ordenados(C) = ubicMIN
            LineaError = "001-0081"
            If C >= UBound(Resultado) Then Exit Do
            LineaError = "001-0082"
            MinSTR = "zzzzzzzzzz"
        Loop
        'cargar todos y sacar la primera columna de las zetas
        LineaError = "001-0083"
        Dim MTXsort() As String
        'ver que haya alguna carpeta
        LineaError = "001-0084"
        If ContadorCarp < 2 Then
            'VER SI HAY UN DISCO Y NO ES EL RANKING
            LineaError = "001-0085"
            If RankToPeople = False And ContadorCarp = 1 Then GoTo EntreAlPedo
            LineaError = "001-0086"
            MsgBox "NO HAY DISCOS PARA MOSTRAR." + vbCrLf + _
            "Una vez iniciado el sistema presione la tecla " + _
            "'C' para ingresar a la configuracion y utilize el " + _
            "asistente para cargar multimedia al sistema."
        End If
EntreAlPedo:
        LineaError = "001-0087"
        Dim nTAPAcd As Integer
        nTAPAcd = 0
        LineaError = "001-0088"
        frmINI.PBar.Width = 0
        LineaError = "001-0089"
        For mtx = LBound(Resultado) + 1 To UBound(Resultado)
            LineaError = "001-0090"
            ReDim Preserve MTXsort(mtx)
            LineaError = "001-0091"
            Dim CarpFull As String, NameCarp As String
            LineaError = "001-0092"
            CarpFull = txtInLista(Resultado(Ordenados(mtx)), 1, ",")
            LineaError = "001-0093"
            NameCarp = txtInLista(Resultado(Ordenados(mtx)), 2, ",")
            LineaError = "001-0094"
            MTXsort(mtx) = CarpFull + "," + NameCarp
            'cargar todas las imagenes si asi esta configurado
            If CargarIMGinicio Then
                LineaError = "001-0095"
                UbicDiscoActual = txtInLista(Resultado(Ordenados(mtx)), 1, ",")
                'caragar las im�genes en diferentes IMGs para que no se cargen despues
                Dim ArchTapa As String
                LineaError = "001-0096"
                ArchTapa = UbicDiscoActual + "\tapa.jpg"
                'arranca con 5 ya cargados
                
                'INICIO RAPIDO fastini
                'si hay, mostrar la tapa
                LineaError = "001-0104"
                If FASTini = False And FSO.FileExists(ArchTapa) Then frmINI.TapaCD.Picture = LoadPicture(ArchTapa)
                LineaError = "001-0105"
                frmINI.lblINI = "Ordenando Discos: " + Trim(Str(mtx))
                LineaError = "001-0106"
                frmINI.lblINI.Refresh
                LineaError = "001-0107"
                frmINI.PBar.Width = frmINI.PBar.Width + 100
                LineaError = "001-0108"
                frmINI.PBar.Refresh
                LineaError = "001-0109"
                If nTAPAcd > ((TapasMostradasH * TapasMostradasV) - 1) Then
                    LineaError = "001-0110"
                    Load frmIndex.TapaCD(nTAPAcd)
                    LineaError = "001-0111"
                    frmIndex.TapaCD(nTAPAcd).Left = frmIndex.TapaCD(nTAPAcd - ((TapasMostradasH * TapasMostradasV))).Left
                    LineaError = "001-0112"
                    frmIndex.TapaCD(nTAPAcd).Top = frmIndex.TapaCD(nTAPAcd - ((TapasMostradasH * TapasMostradasV))).Top
                    LineaError = "001-0113"
                End If
                LineaError = "001-0114"
                If FSO.FileExists(ArchTapa) Then
                    LineaError = "001-0115"
                    frmIndex.TapaCD(nTAPAcd).Picture = LoadPicture(ArchTapa)
                Else
                    LineaError = "001-0116"
                    'ver si hay SuperLicencia!!!
                    If FSO.FileExists(WINfolder + "SL\indexCHI.tbr") Then
                        frmIndex.TapaCD(nTAPAcd).Picture = LoadPicture(WINfolder + "SL\indexCHI.tbr")
                    Else
                        frmIndex.TapaCD(nTAPAcd).Picture = LoadPicture(SYSfolder + "f8ya.nam")
                    End If
                End If
                LineaError = "001-0117"
            End If
            LineaError = "001-0097"
            If nTAPAcd > 0 Then
                LineaError = "001-0098"
                Load frmIndex.L(nTAPAcd)
                LineaError = "001-0099"
                frmIndex.L(nTAPAcd).Top = frmIndex.L(nTAPAcd - 1).Top + frmIndex.L(nTAPAcd - 1).Height
                LineaError = "001-0100"
                frmIndex.L(nTAPAcd).Visible = True
            End If
            LineaError = "001-0101"
            frmIndex.L(nTAPAcd) = NameCarp
            LineaError = "001-0102"
            frmINI.lblProces = NameCarp
            LineaError = "001-0103"
            frmINI.lblProces.Refresh
            nTAPAcd = nTAPAcd + 1
        Next
        LineaError = "001-0118"
        frmINI.lblINI = "Proceso terminado, cargando 3PM..."
        LineaError = "001-0119"
        frmINI.lblINI.Refresh
        LineaError = "001-0120"
        frmINI.PBar.Width = MaxPBAR
        LineaError = "001-0121"
        ObtenerDir = MTXsort
End Function

'cuenta los archivos de determinada extension contenidos en una carpeta
Public Function ContarArch(ByVal ruta As String, EXT As String, VerSubDIRs As Boolean) As Long
    Dim nombres() As String, i As Long, CONT As Long
    LineaError = "001-0122"
    ' asegurarse de que existe una barra invertida inicial
    If Right(ruta, 1) <> "\" Then ruta = ruta & "\"
    ' obtener la lista de archivos ejecutables
    LineaError = "001-0123"
    nombres() = ObtenerArchivos(ruta, EXT)
    'aqui esta la lista, por ahora no la uso
    'For i = 1 To UBound(nombres)
    '    lst.AddItem Ruta & nombres(i)
    'Next
    LineaError = "001-0124"
    CONT = CONT + UBound(nombres)
    LineaError = "001-0125"
    If VerSubDIRs Then
        ' obtener la lista de subdirectorios, incluyendo los ocultos
        ' y ejecutar recursivamente esta rutina en todos ellos.
        LineaError = "001-0126"
        nombres() = ObtenerDir(ruta)
        LineaError = "001-0127"
        For i = 1 To UBound(nombres)
            LineaError = "001-0128"
            ContarArch ruta & nombres(i), EXT, True
        Next
        LineaError = "001-0129"
    End If
End Function

' carga un archivo de texto en un control TextBox

Sub cargarArchivoEnTextBox(NombreArchivo As String, TXT As TextBox)
    LineaError = "001-0130"
    Dim numlib As Integer, isOpen As Boolean
    Dim lineatexto As String, Texto As String
    LineaError = "001-0131"
    On Error GoTo Manejador_Error
    ' obtiene el siguiente n�mero libre de archivo
    LineaError = "001-0132"
    numlib = FreeFile()
    LineaError = "001-0133"
    Open NombreArchivo For Input As #numlib
    ' si el flujo llega hasta aqu�, se habr�n abierto los archivos
    ' sin que se produzca ning�n error
    LineaError = "001-0134"
    isOpen = True
    LineaError = "001-0135"
    Do Until EOF(numlib)
        LineaError = "001-0136"
        Line Input #numlib, lineatexto
        LineaError = "001-0137"
        Texto = Texto & lineatexto & vbCrLf
    Loop
    LineaError = "001-0138"
    ' cargar en el cuadro de texto
    TXT.Text = Texto
    ' se cae intencionadamente en el manejador de error para
    ' cerrar el archivo
Manejador_Error:
    LineaError = "001-0139"
    ' se provoca un error(si es que hay alguno), pero primero
    ' se cierra el archivo
    If isOpen Then Close #numlib
    If Err Then Err.Raise Err.Number, , Err.Description
End Sub

' proporciona el contenido completo de un archivo de texto en una cadena

Function LeerContenidoArchTexto(NombreArchivo As String) As String
    LineaError = "001-0140"
    Dim numlib As Integer, isOpen As Boolean
    On Error GoTo Manejador_Error
    ' obtiene el siguiente n�mero libre de archivo
    LineaError = "001-0141"
    numlib = FreeFile()
    LineaError = "001-0142"
    Open NombreArchivo For Input As #numlib
    ' si el flujo llega hasta aqu�, se habr�n abierto los archivos
    ' sin que se produzca ning�n error
    LineaError = "001-0143"
    isOpen = True
    ' leer todo el contenido en una �nica operaci�n
    LineaError = "001-0144"
    LeerContenidoArchTexto = Input(LOF(numlib), numlib)
    ' se cae intencionadamente en el manejador de error para
    ' cerrar el archivo
Manejador_Error:
    ' se provoca un error(si es que hay alguno), pero primero
    ' se cierra el archivo
    LineaError = "001-0145"
    If isOpen Then Close #numlib
    LineaError = "001-0146"
    If Err Then Err.Raise Err.Number, , Err.Description
End Function

' escribe el contenido de una cadena en un archivo, opcionalmente
' en modo Append

Sub EscribirContenidoArchTexto(Texto As String, NombreArchivo As String, _
    Optional ModoAppend As Boolean)
    LineaError = "001-0147"
        Dim numlib As Integer, isOpen As Boolean
        On Error GoTo Manejador_Error
        ' obtiene el siguiente n�mero libre de archivo
        LineaError = "001-0148"
        numlib = FreeFile()
        LineaError = "001-0149"
        If ModoAppend Then
            LineaError = "001-0150"
            Open NombreArchivo For Append As #numlib
        Else
            LineaError = "001-0151"
            Open NombreArchivo For Output As #numlib
        End If
        ' si el flujo de ejecuci�n llega hasta aqu� es que el archivo
        ' se ha abierto correctamente
        LineaError = "001-0152"
        isOpen = True
        ' imprime al archivo en una �nica operaci�n
        LineaError = "001-0153"
        Print #numlib, Texto
    ' se cae intencionadamente en el manejador de error para
    ' cerrar el archivo
Manejador_Error:
    ' se provoca un error(si es que hay alguno), pero primero
    ' se cierra el archivo
        LineaError = "001-0154"
        If isOpen Then Close #numlib
        LineaError = "001-0155"
        If Err Then Err.Raise Err.Number, , Err.Description
End Sub

' proporciona el contenido de un archivo de texto
' como un array de cadenas.
' NOTA: requiere el empleo de la rutina LeerContenidoArchTexto

Function ObtenLineasArchTexto(NombreArchivo As String, Optional DropEmpty As Boolean, _
    Optional Limite As Variant) As String()
        LineaError = "001-0156"
        Dim ArchTexto As String, elementos() As String, i As Long
        ' lee el contenido del archivo, salir si hay un error
        LineaError = "001-0157"
        ArchTexto = LeerContenidoArchTexto(NombreArchivo)
        ' esto es necesario porque Split() s�lo acepta delimitadores de 1 car�cter
        LineaError = "001-0158"
        ArchTexto = Replace(ArchTexto, vbCrLf, vbCr)
        ' divide al archivo en l�neas individuales de texto
        LineaError = "001-0159"
        elementos() = Split(ArchTexto, vbCr, Limite)
        ' proporciona l�neas sencillas, si se solicita
        LineaError = "001-0160"
        If DropEmpty Then
            ' llena las l�neas vac�as con algo que otros elementos
            ' no contienen
            LineaError = "001-0161"
            For i = 0 To UBound(elementos)
                LineaError = "001-0162"
                If Len(elementos(i)) = 0 Then elementos(i) = vbCrLf
            Next
            ' utiliza la funci�n Filter() para soltar r�pidamente
            ' las l�neas vac�as
            LineaError = "001-0163"
            elementos() = Filter(elementos(), vbCrLf, False)
        End If
        LineaError = "001-0164"
        ObtenLineasArchTexto = elementos()
End Function

' proporciona el contenido de un archivo de texto delimitado como un
' array de arrays de cadenas.
' NOTA: requiere el empleo de las rutinas LeerContenidoArchTexto y
' ObtenLineasArchTexto

Function ImportarArchDelimitado(NombreArchivo As String, _
    Optional delimitador As String = vbTab) As Variant()
        LineaError = "001-0165"
        Dim lineas() As String, i As Long
        ' obtiene todas las lineas contenidas en el archivo,
        ' ignorando las l�neas en blanco
        LineaError = "001-0166"
        lineas() = ObtenLineasArchTexto(NombreArchivo, True)
        ' crea un array de cadena por cada l�nea de texto
        ' y lo almacena en un elemento Variant
        LineaError = "001-0167"
        ReDim valores(0 To UBound(lineas)) As Variant
        LineaError = "001-0168"
        For i = 0 To UBound(lineas)
            LineaError = "001-0169"
            valores(i) = Split(lineas(i), delimitador, -1)
        Next
        LineaError = "001-0170"
        ImportarArchDelimitado = valores()
End Function

' escribir el contenido de un array de arrays de cadenas a un
' archivo de texto deliminado.
' NOTA: necesita la rutina EscribirContenidoArchTexto

Sub ExportarArchDelimitado(valores() As Variant, NombreArchivo As String, _
    Optional delimitador As String = vbTab)
        LineaError = "001-0171"
        Dim i As Long, j As Long, ArchTexto As String
        ' reconstruye las l�neas individuales de texto del archivo
        LineaError = "001-0172"
        ReDim lineas(0 To UBound(valores)) As String
        LineaError = "001-0173"
        For i = 0 To UBound(valores)
            LineaError = "001-0174"
            lineas(i) = Join(valores(i), delimitador)
        Next
        ' introduce CRLFs entre registros
        LineaError = "001-0175"
        ArchTexto = Replace(Join(lineas, vbCr), vbCr, vbCrLf)
        LineaError = "001-0176"
        EscribirContenidoArchTexto ArchTexto, NombreArchivo
End Sub

' duplica el �rbol de directorios sin copiar los archivos

' llamar a esta rutina para iniciar el proceso de copia
' NOTA: la carpeta destino se crear� en caso necesario
'       utiliza el procedimiento Private Sub DuplicarDirArbolSub

Sub DuplicarDirArbol(rutaOrigen As String, rutaDest As String)
    LineaError = "001-0177"
    Dim CarpOrigen As Scripting.Folder, CarpDest As Scripting.Folder
    ' la carpeta origen debe existir
    LineaError = "001-0178"
    Set CarpOrigen = FSO.GetFolder(rutaOrigen)
    ' la carpeta destino se crear� en caso necesario
    LineaError = "001-0179"
    If FSO.FolderExists(rutaDest) Then
        LineaError = "001-0180"
        Set CarpDest = FSO.GetFolder(rutaDest)
    Else
        LineaError = "001-0181"
        Set CarpDest = FSO.CreateFolder(rutaDest)
    End If
    ' saltar a la rutina recursiva para realizar el trabajo real
    LineaError = "001-0181"
    DuplicarDirArbolSub CarpOrigen, CarpDest
End Sub

' Procedimiento recursivo privado utilizado por DuplicarDirArbol

Private Sub DuplicarDirArbolSub(origen As Folder, destino As Folder)
    LineaError = "001-0182"
    Dim CarpOrigen As Scripting.Folder, CarpDest As Scripting.Folder
    LineaError = "001-0183"
    For Each CarpOrigen In origen.SubFolders
        ' copiar esta subcarpeta en la carpeta destino
        LineaError = "001-0184"
        Set CarpDest = destino.SubFolders.Add(CarpOrigen.Name)
        ' repetir el proceso recursivamente para todas las
        ' subcarpetas de la carpeta considerada
        LineaError = "001-0185"
        DuplicarDirArbolSub CarpOrigen, CarpDest
    Next
End Sub

' Busca una cadena en todos los archivos TXT contenidos en un directorio.

' Por cada archivo localizado devuelve un elemento Variant que contiene
' un array de tres elementos: el nombre del archivo, la l�nea
' y el n�mero de columna.
' NOTA: las b�squedas no distinguen el uso de may�sculas y min�sculas

Function BuscarArchTexto(ruta As String, buscar As String) As Variant()
    LineaError = "001-0186"
    Dim fil As Scripting.File, ts As Scripting.TextStream
    Dim pos As Long, ContadorArch As Long
    LineaError = "001-0187"
    ReDim Resultado(50) As Variant
    ' buscar for all the TXT files in the directory
    LineaError = "001-0188"
    For Each fil In FSO.GetFolder(ruta).Files
        LineaError = "001-0189"
        If UCase$(FSO.GetExtensionName(fil.path)) = "TXT" Then
            ' obtener el objeto TextStream correspondiente
            LineaError = "001-0190"
            Set ts = fil.OpenAsTextStream(ForReading)
            ' leer su contenido, buscar la cadena y cerrarlo
            LineaError = "001-0191"
            pos = InStr(1, ts.ReadAll, buscar, vbTextCompare)
            LineaError = "001-0192"
            ts.Close
            LineaError = "001-0193"
            If pos > 0 Then
                ' si se encuentra la cadena, reabre el archivo
                ' para determinar su posici�n en forma de (l�nea, columna)
                LineaError = "001-0194"
                Set ts = fil.OpenAsTextStream(ForReading)
                ' salta todos los caracteres precedentes para saber d�nde se
                ' encuentra la cadena
                LineaError = "001-0194"
                ts.Skip pos - 1
                ' llena el array resultado, hace sitio en caso necesario
                ContadorArch = ContadorArch + 1
                LineaError = "001-0195"
                If ContadorArch > UBound(Resultado) Then
                    LineaError = "001-0196"
                    ReDim Preserve Resultado(UBound(Resultado) + 50) As Variant
                End If
                ' cada array resultado tiene tres elementos
                LineaError = "001-0197"
                Resultado(ContadorArch) = Array(fil.path, ts.Line, ts.Column)
                ' ahora podemos cerrar el TextStrem
                LineaError = "001-0198"
                ts.Close
            End If
        End If
    Next
    ' cambia el tama�o del array resultado para indicar el n�mero de
    ' coincidencas
    LineaError = "001-0199"
    ReDim Preserve Resultado(0 To ContadorArch) As Variant
    LineaError = "001-0200"
    BuscarArchTexto = Resultado
End Function

' espera un n�mero de milisegundos y devuelve el estado de ejecuci�n de un
' proceso; si se omite el argumento, espera hasta que el proceso finalice.

Function EsperarPorProceso(taskId As Long, Optional msecs As Long = -1) _
    As Boolean
        LineaError = "001-0201"
        Dim procHandle As Long
        ' obtiene el manejador del proceso
        LineaError = "001-0202"
        procHandle = OpenProcess(&H100000, True, taskId)
        ' verifica su estado se�alado, lo devuelve al que hizo la llamada
        LineaError = "001-0203"
        EsperarPorProceso = EsperarUnicoObjeto(procHandle, msecs) <> -1
        ' cierra el gestor
        LineaError = "001-0204"
        CloseHandle procHandle
End Function

Public Function LeerArch1Linea(Arch As String) As String
    LineaError = "001-0205"
    If FSO.FileExists(Arch) = False Then
        LineaError = "001-0206"
        LeerArch1Linea = "No existe archivo"
        LineaError = "001-0207"
        Exit Function
    End If
    LineaError = "001-0208"
    Set TE = FSO.OpenTextFile(Arch, ForReading, False)
    LineaError = "001-0209"
    LeerArch1Linea = TE.ReadLine
    LineaError = "001-0210"
    TE.Close
End Function

Public Sub EscribirArch1Linea(Arch As String, TXT As String)
    LineaError = "001-0211"
    Set TE = FSO.CreateTextFile(Arch, True)
    LineaError = "001-0212"
    TE.WriteLine TXT
    LineaError = "001-0213"
    TE.Close
End Sub

Public Function ObtenerArchMM(Carpeta As String) As String()
    'devuelve "Carpeta + NombreArchivo + "," + NombreArchivo"
    'devuelve PathFull,SoloNombre

    'ADEM�S DEBO ASEGURARME QUE NO HAYA COMAS EN LOS NOMBRES
    LineaError = "001-0214"
    If Right(Carpeta, 1) <> "\" Then Carpeta = Carpeta + "\"
    LineaError = "001-0215"
    Dim TMPmatriz() As String
    ReDim Preserve TMPmatriz(0)
    'mp3
    LineaError = "001-0216"
    Dim NombreArchivo As String, ContadorArch As Long, NewName As String
    NombreArchivo = Dir$(Carpeta + "*.mp3")
    LineaError = "001-0217"
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        LineaError = "001-0218"
        NewName = QuitarCaracter(NombreArchivo, ",")
        LineaError = "001-0219"
        If NombreArchivo <> NewName Then
            'no se puede corregir si es un CD. Solo corrige si es disco duro
            'esta funcion se usa para leer CDs debo prevenir
            LineaError = "001-0220"
            If FSO.Drives(Left(Carpeta, 1)).DriveType = Fixed Then
                LineaError = "001-0221"
                FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
                LineaError = "001-0222"
                WriteTBRLog "Se corrigio el nombre de tema " + NombreArchivo + _
                    " por " + NewName + " en la carpeta " + Carpeta, True
                    LineaError = "001-0223"
                NombreArchivo = NewName
            End If
        End If
        LineaError = "001-0224"
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        LineaError = "001-0225"
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "," + NombreArchivo
        LineaError = "001-0226"
        NombreArchivo = Dir$
    Loop
    
    'mpg
    LineaError = "001-0227"
    NombreArchivo = Dir$(Carpeta + "\*.mpg")
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        LineaError = "001-0228"
        NewName = QuitarCaracter(NombreArchivo, ",")
        LineaError = "001-0229"
        If NombreArchivo <> NewName Then
            LineaError = "001-0230"
            FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
            LineaError = "001-0231"
            WriteTBRLog "Se corrigio el nombre de tema " + NombreArchivo + _
                " por " + NewName + " en la carpeta " + Carpeta, True
            LineaError = "001-0232"
            NombreArchivo = NewName
        End If
        LineaError = "001-0233"
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        LineaError = "001-0234"
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "," + NombreArchivo
        LineaError = "001-0235"
        NombreArchivo = Dir$
    Loop
    
    'mpeg
    LineaError = "001-0236"
    NombreArchivo = Dir$(Carpeta + "\*.mpeg")
    LineaError = "001-0237"
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        LineaError = "001-0238"
        NewName = QuitarCaracter(NombreArchivo, ",")
        LineaError = "001-0239"
        If NombreArchivo <> NewName Then
            LineaError = "001-0240"
            FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
            LineaError = "001-0241"
            WriteTBRLog "Se corrigio el nombre de tema " + NombreArchivo + _
                " por " + NewName + " en la carpeta " + Carpeta, True
            LineaError = "001-0242"
            NombreArchivo = NewName
        End If
        LineaError = "001-0243"
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        LineaError = "001-0244"
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "," + NombreArchivo
        LineaError = "001-0245"
        NombreArchivo = Dir$
    Loop
    
    'avi
    LineaError = "001-0246"
    NombreArchivo = Dir$(Carpeta + "\*.avi")
    LineaError = "001-0247"
    Do While Len(NombreArchivo)
        'corregir el nombre del tema
        LineaError = "001-0248"
        NewName = QuitarCaracter(NombreArchivo, ",")
        LineaError = "001-0249"
        If NombreArchivo <> NewName Then
            LineaError = "001-0250"
            FSO.MoveFile Carpeta + NombreArchivo, Carpeta + NewName
            LineaError = "001-0251"
            WriteTBRLog "Se corrigio el nombre de tema " + NombreArchivo + _
                " por " + NewName + " en la carpeta " + Carpeta, True
            LineaError = "001-0252"
            NombreArchivo = NewName
        End If
        LineaError = "001-0253"
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        LineaError = "001-0254"
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo + "," + NombreArchivo
        LineaError = "001-0255"
        NombreArchivo = Dir$
    Loop
    LineaError = "001-0256"
    ObtenerArchMM = TMPmatriz
End Function

Public Function QuitarCaracter(FileOrFolder As String, _
    CaracterToKill As String) As String
    'sacar en caracter de una cadena
    'lo uso para sacar las comas de los archivos mp3
    'o los puntos de los nombre de los discos
    LineaError = "001-0257"
    Dim SeCambio As Boolean
    Dim TMPfOf 'temporario de file or folder
    LineaError = "001-0258"
    TMPfOf = FileOrFolder
    Dim FindIn As Long
    Dim Parte1 As String, Parte2 As String
    LineaError = "001-0259"
    SeCambio = False
    Do
        LineaError = "001-0260"
        FindIn = InStr(TMPfOf, CaracterToKill)
        If FindIn > 0 Then
            LineaError = "001-0261"
            SeCambio = True
            LineaError = "001-0262"
            Parte1 = Mid(TMPfOf, 1, FindIn - 1)
            LineaError = "001-0263"
            Parte2 = Mid(TMPfOf, FindIn + 1, Len(TMPfOf) - FindIn)
            LineaError = "001-0264"
            TMPfOf = Parte1 + Parte2
        Else
            LineaError = "001-0265"
            Exit Do
        End If
    Loop
    LineaError = "001-0266"
    QuitarCaracter = TMPfOf
    
End Function
