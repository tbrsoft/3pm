VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' cComDlg Clase para simular el control de Diálogos Comunes
'
' Primera tentativa:                                        (04:57 25/Ago/1997)
'
' Versión reducida, sólo Abrir                              (10:20 18/Sep/1997)
' Versión reducida Diálogos de Abrir y Guardar                    (21/Oct/1997)
' Con BrowseForFolder                                       (01:20 04/Dic/2000)
'
' Revisada para Windows 2000                                (15:12 27/Dic/2000)
' Este fichero no funcionaba, pero si la de sólo Abrir
'
' ©Guillermo 'guille' Som, 1997-2000 <guille@costasol.net>
'------------------------------------------------------------------------------
Option Explicit

Private sFilter As String

'Esta propiedad hará referencia al hWnd de un Form
Public hWnd As Long

'Propiedades genéricas de los diálogos comunes
Public DialogTitle As String
Public CancelError As Boolean
Public Flags As eOFN 'Long

'Propiedades para Abrir y Guardar como
Public DefaultExt As String
Public FileName As String
Public FileTitle As String
'Public Filter As String
Public FilterIndex As Long
Public InitDir As String
'Public MaxFileSize As Long (será 260)

'----------------------------------------------------------------------------
'Estructura de datos para Abrir y Guardar como...
'----------------------------------------------------------------------------
Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
    (pOpenFilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" _
    (pOpenFilename As OPENFILENAME) As Long

'------------------------------------------------------------------------------
' Constantes para las funciones de archivos
Public Enum eOFN
    'Tamaño máximo de un nombre de archivo (incluyendo el path)
    MAX_PATH = 260
    'Constantes para el diálogo de archivos
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
    OFN_EXPLORER = &H80000                         '  new look commdlg
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
    '
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
End Enum

'------------------------------------------------------------------------------
' Para la función BrowseForFolders                                  (04/Dic/00)
'------------------------------------------------------------------------------
Public Enum eBIF
    BIF_RETURNONLYFSDIRS = &H1          ' Sólo directorios del sistema
    BIF_DONTGOBELOWDOMAIN = &H2         ' No incluir carpetas de red
    BIF_STATUSTEXT = &H4                '
    BIF_RETURNFSANCESTORS = &H8         '
    BIF_BROWSEFORCOMPUTER = &H1000      ' Buscar PCs
    BIF_BROWSEFORPRINTER = &H2000       ' Buscar impresoras
    BIF_BROWSEINCLUDEFILES = &H4000&    ' Incluir los ficheros      (04/Abr/01)
                                        ' (esta constante no estaba asignada)
End Enum

'Valores para usar con pIDLRoot
'Public Enum ShellSpecialFolderConstants
'    ssfDESKTOP = &H0
'    ssfPROGRAMS = &H2
'    ssfCONTROLS = &H3
'    ssfPRINTERS = &H4
'    ssfPERSONAL = &H5
'    ssfFAVORITES = &H6
'    ssfSTARTUP = &H7
'    ssfRECENT = &H8
'    ssfSENDTO = &H9
'    ssfBITBUCKET = &HA
'    ssfSTARTMENU = &HB
'    ssfDESKTOPDIRECTORY = &H10
'    ssfDRIVES = &H11
'    ssfNETWORK = &H12
'    ssfNETHOOD = &H13
'    ssfFONTS = &H14
'    ssfTEMPLATES = &H15
'End Enum

'Estructuras
Private Type SHFILEOPSTRUCT
    hWnd                    As Long
    wFunc                   As Long
    pFrom                   As String
    pTo                     As String
    fFlags                  As Integer
    fAnyOperationsAborted   As Boolean
    hNameMappings           As Long
    lpszProgressTitle       As String
End Type

'Declaración de SHFILEOPSTRUCT
'typedef WORD FILEOP_FLAGS;
'
'typedef struct _SHFILEOPSTRUCTA
'{
'        HWND            hwnd;
'        UINT            wFunc;
'        LPCSTR          pFrom;
'        LPCSTR          pTo;
'        FILEOP_FLAGS    fFlags;
'        BOOL            fAnyOperationsAborted;
'        LPVOID          hNameMappings;
'        LPCSTR           lpszProgressTitle; // only used if FOF_SIMPLEPROGRESS
'} SHFILEOPSTRUCTA, FAR *LPSHFILEOPSTRUCTA;

'también me he encontrado con esta declaración:
'(pero después de comprobar cómo se declara en ShellApi.h...)
'Private Type SHFILEOPSTRUCT2
'    hWnd                    As Long
'    wFunc                   As Long
'    pFrom                   As String
'    pTo                     As String
'    fFlags                  As Long
'    fAnyOperationsAborted   As Long
'    hNameMappings           As Long
'    lpszProgressTitle       As String
'End Type

Private Type BrowseInfo
    hwndOwner               As Long
    pIDLRoot                As Long         ' Especifica dónde se empezará a mostrar
    'pszDisplayName          As String 'Long
    pszDisplayName          As Long         ' El nombre del directorio, sin el Path completo
    lpszTitle               As String 'Long
    ulFlags                 As Long
    lpfnCallback            As Long
    lParam                  As Long
    iImage                  As Long
End Type

' Funciones del API
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
        (lpFileOp As SHFILEOPSTRUCT) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
        (lpbi As BrowseInfo) As Long

Private Declare Sub CoTaskMemFree Lib "OLE32.DLL" _
        (ByVal hMem As Long)

Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" _
        (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
        (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function BrowseForFolder(ByVal hwndOwner As Long, _
                                ByVal sPrompt As String, _
                                Optional ByVal lFlags As eBIF = BIF_RETURNONLYFSDIRS) As String
    '
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    'Dim lFlags As Long

    'If IsMissing(vFlags) Then
    '    lFlags = BIF_RETURNONLYFSDIRS
    'Else
'    If Not IsMissing(vFlags) Then
'        lFlags = CInt(vFlags)
'        'If lFlags = 0 Then
'        '    lFlags = BIF_RETURNONLYFSDIRS
'        'End If
'    End If

    With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = sPrompt 'lstrcat(sPrompt, "")
        .ulFlags = lFlags 'Or BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    Else
        ' Se ha pulsado en cancelar
        sPath = ""
        If CancelError Then
            With Err
                .Source = "cComDialog.BrowseForFolder"
                .Number = 32755
                .Description = "Cancelada la operación de BrowseForFolder"
            End With
        End If
    End If

    BrowseForFolder = sPath
End Function

Public Function ShowOpen(Optional ByVal vFileName As String = "", _
                         Optional ByVal vTitle As String = "", _
                         Optional ByVal vFilter As String = "", _
                         Optional ByVal vFlags As Long = 0, _
                         Optional ByVal vhWnd As Long = 0) As Boolean
    '----------------------------------------------------------
    'Método para mostrar el cuadro de diálogo de Abrir
    '
    '(c) Guillermo Som Cerezo                  24/Oct/93
    '
    'Convertido en objeto (clase)                   (25/Ago/97)
    '
    'Los parámetros opcionales especificarán:
    '   vFileName   El nombre del archivo
    '   vTitle      Título del cuadro de diálogo
    '   vFilter     Extensiones
    '   vFlags      Los flags
    '   vhWnd       El hWnd del Form
    '----------------------------------------------------------
    Dim resultado As Long
    Dim ofn As OPENFILENAME

    Err.Clear
    Err.Number = 0
    '
    If Len(vFileName) Then _
        FileName = CStr(vFileName)
    If vhWnd <> 0 Then _
        hWnd = CLng(vhWnd)
    If Len(vFilter) Then _
        Me.Filter = CStr(vFilter)
    If Len(vTitle) Then _
        DialogTitle = CStr(vTitle)
    If vFlags <> 0 Then _
        Flags = CLng(vFlags)
    '
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = hWnd
        .hInstance = 0
        If Len(sFilter) = 0 Then _
            sFilter = "Todos los Archivos (*.*)" & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
        '
        .lpstrFilter = sFilter
        .nFilterIndex = FilterIndex
        .lpstrFile = Left$(FileName & String$(MAX_PATH, 0), MAX_PATH)
        .nMaxFile = MAX_PATH
        .nFileOffset = 0
        .nFileExtension = 0
        .lpstrDefExt = DefaultExt
        .lpstrFileTitle = Left$(FileTitle & String$(260, 0), MAX_PATH)
        .nMaxFileTitle = MAX_PATH
        .lpstrInitialDir = Left$(InitDir & String$(260, 0), MAX_PATH)
        '
        'Nombres largos y estilo explorer           (21/Oct/97)
        'y otros valore "obvios"
        'Flags = Flags Or OFN_LONGNAMES Or OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
        .Flags = Flags
        If Len(DialogTitle) = 0 Then    'Si no se especifica el título
            DialogTitle = "Abrir"
        End If
        .lpstrTitle = DialogTitle
        '
        .nFileOffset = 0
        .lpstrDefExt = 0
        .lCustData = 0
        .lpfnHook = 0
        .lpTemplateName = 0
    End With

    resultado = GetOpenFileName(ofn)
    If resultado <> 0 Then
        If Flags And OFN_ALLOWMULTISELECT Then
            'Si está multiselect, se separan los nombres con Chr$(0)
            FileName = Chr$(34) & Trim$(QuitarCaracterEx(ofn.lpstrFile, Chr$(0), Chr$(34) & " " & Chr$(34))) & Chr$(34)
            FileTitle = Chr$(34) & Trim$(QuitarCaracterEx(ofn.lpstrFileTitle, Chr$(0), Chr$(34) & " " & Chr$(34))) & Chr$(34)
            InitDir = Chr$(34) & Trim$(QuitarCaracterEx(ofn.lpstrInitialDir, Chr$(0), Chr$(34) & " " & Chr$(34))) & Chr$(34)
        Else
            FileName = Left$(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr$(0)) - 1)
            FileTitle = Left$(ofn.lpstrFileTitle, InStr(ofn.lpstrFileTitle, Chr$(0)) - 1)
            InitDir = Left$(ofn.lpstrInitialDir, InStr(ofn.lpstrInitialDir, Chr$(0)) - 1)
        End If
    Else
        If CancelError Then
            'Err.Raise 32755, "cComDlgR.ShowOpen", "Error en Abrir (clase cComDlg)"
            With Err
                .Source = "cComDlgR2.ShowOpen"
                .Number = 32755
                .Description = "Error en Abrir (clase cComDlgR)"
            End With
        End If
    End If
    'Devuelve True si se puede abrir
    ShowOpen = (resultado <> 0)
End Function

Public Sub AgregarAText(ByVal sArchivos As String, queControl As Control, Optional ByVal vSeparador As Variant)
    'agregar los archivos indicados al control indicado
    'Los archivos estarán separados por espacios y dentro de comillas
    'o simplemente será un archivo
    Dim i&, j&
    Dim sTmp$, sDir$
    Dim colArchivos As New Collection
    Dim sSeparador As String

    sArchivos = Trim$(sArchivos)
    If Len(sArchivos) = 0 Then Exit Sub

    If IsMissing(vSeparador) Then
        sSeparador = " "
    Else
        sSeparador = CStr(vSeparador)
    End If

    If InStr(sArchivos, Chr$(34)) Then
        'hay comillas, es que hay varios archivos
        j = 0
        Do While Len(sArchivos)
            Do While Left$(sArchivos, 1) = Chr$(34)
                sArchivos = Trim$(Mid$(sArchivos, 2))
            Loop
            i = InStr(sArchivos, Chr$(34))
            If i Then
                sTmp = Left$(sArchivos, i - 1)
                sArchivos = Trim$(Mid$(sArchivos, i + 1))
                If j Then
                    colArchivos.Add sDir & sTmp
                Else
                    'El primer parámetro es el directorio
                    j = j + 1
                    sDir = sTmp
                    'Si no tiene la barra ponersela
                    If Right$(sDir, 1) <> "\" Then
                        sDir = sDir & "\"
                    End If
                End If
            Else
                Exit Do
            End If
        Loop
        If Len(sArchivos) Then
            colArchivos.Add sDir & sArchivos
        End If
        'Por si sólo se selecciona un archivo
        If colArchivos.Count = 0 Then
            colArchivos.Add sTmp
        End If
    Else
        'no hay comillas, es sólo un archivos
        colArchivos.Add sArchivos
    End If

    'Asignar los datos anteriores
    sTmp = Trim$(queControl)
    If Len(sTmp) Then
        If Right$(RTrim$(sTmp), 1) <> sSeparador Then
            sTmp = sTmp & sSeparador
        End If
    End If
    For i = colArchivos.Count To 1 Step -1
        If i > 1 Then
            sTmp = sTmp & colArchivos(i) & sSeparador
        Else
            sTmp = sTmp & colArchivos(i)
        End If
    Next
    queControl = sTmp

    Set colArchivos = Nothing
End Sub

Public Sub AgregarALista(ByVal sArchivos As String, queControl As Control, Optional ByVal vAlPrincipio As Variant)
    'agregar los archivos indicados a la lista
    'Los archivos estarán separados por espacios y dentro de comillas
    'o simplemente será un archivo
    Dim i&, j&
    Dim sTmp$, sDir$
    Dim bAlPrincipio As Boolean
    Dim colArchivos As New Collection

    sArchivos = Trim$(sArchivos)
    If Len(sArchivos) = 0 Then Exit Sub

    If IsMissing(vAlPrincipio) Then
        bAlPrincipio = False
    Else
        bAlPrincipio = CBool(vAlPrincipio)
    End If

    If InStr(sArchivos, Chr$(34)) Then
        'hay comillas, es que hay varios archivos
        j = 0
        Do While Len(sArchivos)
            Do While Left$(sArchivos, 1) = Chr$(34)
                sArchivos = Trim$(Mid$(sArchivos, 2))
            Loop
            i = InStr(sArchivos, Chr$(34))
            If i Then
                sTmp = Left$(sArchivos, i - 1)
                sArchivos = Trim$(Mid$(sArchivos, i + 1))
                If j Then
                    colArchivos.Add sDir & sTmp
                Else
                    'El primer parámetro es el directorio
                    j = j + 1
                    sDir = sTmp
                    'Si no tiene la barra ponersela
                    If Right$(sDir, 1) <> "\" Then
                        sDir = sDir & "\"
                    End If
                End If
            Else
                Exit Do
            End If
        Loop
        If Len(sArchivos) Then
            colArchivos.Add sDir & sArchivos
        End If
        'Por si sólo se selecciona un archivo
        If colArchivos.Count = 0 Then
            colArchivos.Add sTmp
        End If
    Else
        'no hay comillas, es sólo un archivos
        colArchivos.Add sArchivos
    End If

    For i = colArchivos.Count To 1 Step -1
        If bAlPrincipio Then
            queControl.AddItem colArchivos(i), 0
        Else
            queControl.AddItem colArchivos(i)
        End If
    Next

    Set colArchivos = Nothing
End Sub

Public Function QuitarCaracterEx(ByVal sValor As String, ByVal sCaracter As String, Optional ByVal sPoner) As String
    '----------------------------------------------------------
    ' Cambiar/Quitar caracteres                     (17/Sep/97)
    ' Si se especifica sPoner, se cambiará por ese carácter
    '
    'Esta versión permite cambiar los caracteres    (17/Sep/97)
    'y sustituirlos por el/los indicados
    'a diferencia de QuitarCaracter, no se buscan uno a uno,
    'sino todos juntos
    '----------------------------------------------------------
    Dim i As Long
    Dim sCh As String
    Dim bPoner As Boolean
    Dim iLen As Long

    bPoner = False
    If Not IsMissing(sPoner) Then
        sCh = sPoner
        bPoner = True
    End If
    iLen = Len(sCaracter)
    If iLen = 0 Then
        QuitarCaracterEx = sValor
        Exit Function
    End If

    'Si el caracter a quitar/cambiar es Chr$(0), usar otro método
    If Asc(sCaracter) = 0 Then
        'Quitar todos los chr$(0) del final
        Do While Right$(sValor, 1) = Chr$(0)
            sValor = Left$(sValor, Len(sValor) - 1)
            If Len(sValor) = 0 Then Exit Do
        Loop
        iLen = 1
        Do
            i = InStr(iLen, sValor, sCaracter)
            If i Then
                If bPoner Then
                    sValor = Left$(sValor, i - 1) & sCh & Mid$(sValor, i + 1)
                Else
                    sValor = Left$(sValor, i - 1) & Mid$(sValor, i + 1)
                End If
                iLen = i
            Else
                'ya no hay más, salir del bucle
                Exit Do
            End If
        Loop
    Else
        i = 1
        Do While i <= Len(sValor)
            'Debug.Print Mid$(sValor, i, 1); Asc(Mid$(sValor, i, 1));
            If Mid$(sValor, i, iLen) = sCaracter Then
                If bPoner Then
                    sValor = Left$(sValor, i - 1) & sCh & Mid$(sValor, i + iLen)
                    i = i - 1
                Else
                    sValor = Left$(sValor, i - 1) & Mid$(sValor, i + iLen)
                End If
            End If

            i = i + 1
        Loop
    End If

    QuitarCaracterEx = sValor
End Function

Public Function ShowSave(Optional ByVal vFileName As String = "", _
                         Optional ByVal vTitle As String = "", _
                         Optional ByVal vFilter As String = "", _
                         Optional ByVal vFlags As Long = 0, _
                         Optional ByVal vhWnd As Long = 0) As Boolean
    '----------------------------------------------------------
    'Método para mostrar el cuadro de diálogo de Guardar como...
    '
    '(c) Guillermo Som Cerezo                  24/Oct/93
    '
    'Convertido en objeto (clase)                   (25/Ago/97)
    '
    'Los parámetros opcionales especificarán:
    '   vFileName   El nombre del archivo
    '   vTitle      Título del cuadro de diálogo
    '   vFilter     Extensiones
    '   vFlags      Los flags
    '   vhWnd       El hWnd del Form
    '----------------------------------------------------------
    Dim resultado As Long
    Dim ofn As OPENFILENAME

    Err.Clear
    Err.Number = 0

    If Len(vFileName) Then _
        FileName = CStr(vFileName)
    If vhWnd <> 0 Then _
        hWnd = CLng(vhWnd)
    If Len(vFilter) Then _
        Me.Filter = CStr(vFilter)
    If Len(vTitle) Then _
        DialogTitle = CStr(vTitle)
    If vFlags <> 0 Then _
        Flags = CLng(vFlags)
    '
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = hWnd
        .hInstance = 0
        If Len(sFilter) = 0 Then _
            sFilter = "Todos los Archivos (*.*)" & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
        .lpstrFilter = sFilter
        '.lpstrCustomFilter = ""
        '.nMaxCustFilter = 0
        .nFilterIndex = FilterIndex
        .lpstrFile = Left$(FileName & String$(260, 0), 260)
        .nMaxFile = 260
        .lpstrFileTitle = Left$(FileTitle & String$(260, 0), 260)
        .nMaxFileTitle = 260
        .lpstrDefExt = DefaultExt
        .lpstrInitialDir = Left$(InitDir & String$(260, 0), 260)
        '
        'Nombres largos y estilo explorer           (21/Oct/97)
        'Flags = Flags Or OFN_LONGNAMES Or OFN_EXPLORER Or OFN_HIDEREADONLY
        '
        .Flags = Flags
        If Len(DialogTitle) = 0 Then
            DialogTitle = "Guardar como..."
        End If
        .lpstrTitle = DialogTitle
        '
        .nFileOffset = 0
        .lpstrDefExt = 0
        .lCustData = 0
        .lpfnHook = 0
        .lpTemplateName = 0
        '
    End With
    '
    resultado = GetSaveFileName(ofn)
    If resultado <> 0 Then
        FileName = Left$(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr$(0)) - 1)
        FileTitle = Left$(ofn.lpstrFileTitle, InStr(ofn.lpstrFileTitle, Chr$(0)) - 1)
        InitDir = Left$(ofn.lpstrInitialDir, InStr(ofn.lpstrInitialDir, Chr$(0)) - 1)
    Else
'        If CancelError Then
'            Err.Raise 32755, "cComDlg.ShowSave", "Error en Guardar como... (clase cComDlg)"
'        End If
        If CancelError Then
            'Err.Raise 32755, "cComDlgR2.ShowOpen", "Error en Abrir (clase cComDlg)"
            With Err
                .Source = "cComDlgR2.ShowSave"
                .Number = 32755
                .Description = "Error en Guardar como... (clase cComDlgR2)"
            End With
        End If
    End If
    'Devuelve True si se puede abrir
    ShowSave = (resultado <> 0)

'    Err = 0
'    On Error GoTo 0
End Function

Public Property Let Action(vNewValue As Integer)
    '0   Ninguna acción.
    '1   Muestra el cuadro de diálogo Abrir.
    '2   Muestra el cuadro de diálogo Guardar como.
    '3   Muestra el cuadro de diálogo Color.
    '4   Muestra el cuadro de diálogo Fuente.
    '5   Muestra el cuadro de diálogo Impresora.
    '6   Ejecuta WINHELP.EXE.
    '
    Select Case vNewValue
    Case 1: ShowOpen
    Case 2: ShowSave
    Case 3: 'ShowColor
    Case 4: 'ShowFont
    Case 5: 'ShowPrinter
    Case 6: 'ShowHelp
    Case Else
        'nada que mostrar
    End Select
End Property

Public Property Let Filter(ByVal sNewFilter As String)
    'Procesar el parámetro para convertirlo a formato C,
    'Se usará | como separador.
    Dim i As Integer, j As Integer
    Dim sTmp As String

    sTmp = ""
    If InStr(sNewFilter, "|") Then
        sNewFilter = Trim$(sNewFilter)
        If Right$(sNewFilter, 1) <> "|" Then
            sNewFilter = sNewFilter & "|"
        End If
        Do
            i = InStr(sNewFilter, "|")
            If i Then
                sTmp = sTmp & Left$(sNewFilter, i - 1) & Chr$(0)
                sNewFilter = Mid$(sNewFilter, i + 1)
            Else
                Exit Do
            End If
        Loop While i
        If Right$(sTmp, 1) = Chr$(0) Then
            sNewFilter = sTmp & Chr$(0)
        Else
            sNewFilter = sTmp & Chr$(0) & Chr$(0)
        End If
    ElseIf InStr(sNewFilter, Chr$(0)) = 0 Then
        sNewFilter = ""
    End If
    sFilter = sNewFilter
End Property

Public Function OpenFile(ByVal sFile As String, sCadena As String) As Boolean
    '----------------------------------------------------------
    'Abrir el archivo y asignarlo a una cadena      (27/Ago/97)
    '
    'Entrada:
    '   sFile       Archivo dónde se guardará
    '   sCadena     Cadena a guardar
    'Salida:
    '   True        Si NO se pudo abrir
    '----------------------------------------------------------
    Dim nF As Integer

    On Local Error GoTo OpenFileErr

    If Len(Dir$(sFile)) Then
        'Abrir y guardar el contenido de sCadena
        nF = FreeFile
        Open sFile For Input As nF
        sCadena = Input$(LOF(nF), nF)
        Close nF
    Else
        Error.Raise 76
    End If

SalirOpenFile:
    Err = 0
    On Local Error GoTo 0
    Exit Function
OpenFileErr:
    OpenFile = True
    GoTo SalirOpenFile
End Function

Public Function SaveFile(ByVal sFile As String, sCadena As String, Optional vOverWrite) As Boolean
    '----------------------------------------------------------
    'Guarda una cadena en un archivo                (27/Ago/97)
    '
    'Entrada:
    '   sFile       Archivo dónde se guardará
    '   sCadena     Cadena a guardar
    '   vOverWrite  Si se pide confirmación antes de sobreescribir
    'Salida:
    '   True        Si NO se pudo guardar
    '----------------------------------------------------------
    Dim nF As Integer
    Dim bOverWrite As Boolean

    'Si no se especifica, se sobreescribe sin avisar
    If IsMissing(vOverWrite) Then
        bOverWrite = True
    Else
        bOverWrite = CBool(vOverWrite)
    End If

    On Local Error GoTo SaveFileErr

    If Len(Dir$(sFile)) Then
        If bOverWrite Then
            'Preguntar si se sobreescribe
            If MsgBox("Ya existe el archivo, ¿quieres sobreescribirlo?", vbYesNo + vbQuestion, "Guardar Archivo") = vbNo Then
                Error.Raise 76
            End If
        End If
        Kill sFile
    End If
    Err = 0
    'Guardar el contenido de sCadena
    nF = FreeFile
    Open sFile For Output As nF
    Print #nF, sCadena
    Close nF

SalirSaveFile:
    Err = 0
    On Local Error GoTo 0
    Exit Function
SaveFileErr:
    SaveFile = True
    GoTo SalirSaveFile
End Function

