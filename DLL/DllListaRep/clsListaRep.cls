VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListaRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'************************************************************************************
'maneja listas de canciones a reproducir
'debería incluir enganches de canciones y los indice del reproductor que se estan usando
'************************************************************************************

Private ListaFalta() As String 'solo los paths completos de los que falta
'el indice cero es solo para que no este vacia, es siempre ""

Private ListaTag() As String
'algunos elementos se agregan como publicidades o como gratis, etc.
'cada programa carga lo que quiere aqui !

Private FSO As New Scripting.FileSystemObject
Private mGrabaAuto As Boolean
Private mArchivo As String

Public Enum MMType 'tipo de archivo multimedia
    'solo musica
    mmtypeSinDEFINIR = 0
    mmtypeMUSICA = 1
    mmtypeVIDEO = 2
    mmtypeKARAOKEMN1 = 3
    mmtypeKARAOKEMPG = 4
End Enum

Public Function ListaAdd(ArchToAdd As String, Optional sTAG As String = "", _
    Optional nOrden As Long = -1)
    
    '-1 significa al final
    '0 es ahora mismo cortando cualquier cosa (no programado aun)
    '1 es el proximo
    '2 es el segundo etc.
    
    'DEVUELVE
    '0 el archivo existe y se agrego, todo ok
    '1 el archivo no existe !!!. no se agrego
    
    'ver si existe
    If FSO.FileExists(ArchToAdd) = False Then
        ListaAdd = 1
        Exit Function
    End If

    'si pide un orden especial y esta fuera del rango va al ultimo co mo un comun
    If (nOrden > 0) And (nOrden > UBound(ListaFalta)) Then nOrden = -1

    'para uso en diferentes lugares
    Dim NewInd As Long
    NewInd = UBound(ListaFalta) + 1

    If nOrden = -1 Then 'lo paso al ultimo
        ReDim Preserve ListaFalta(NewInd): ListaFalta(NewInd) = ArchToAdd
        ReDim Preserve ListaTag(NewInd): ListaTag(NewInd) = sTAG
    End If
    
    If nOrden > 0 Then
        'ponerlo en el orden que corresponde
        
        ReDim Preserve ListaFalta(NewInd) 'agregar uno mas !
        ReDim Preserve ListaTag(NewInd)
        
        Dim A As Long
        For A = NewInd - 1 To nOrden Step -1
            'paso a todos uno mas atras
            ListaFalta(A + 1) = ListaFalta(A)
            ListaTag(A + 1) = ListaTag(A)
        Next A
        
        ListaFalta(nOrden) = ArchToAdd
        ListaTag(nOrden) = sTAG
        
    End If
    
    If nOrden = 0 Then
        'ejecutarlo yaaaa
        '*****
    End If
    
    'ver si graba automaticamente
    If mGrabaAuto And mArchivo <> "" Then ListaGuardarADisco mArchivo
    
End Function

Public Function ListaKillElement(Optional nOrden As Long = 1) As Long
    'saca uno, se usa para quitar el primero cuando se preprodujo
    'ele elemto 1 es el que sigue siempre
    
    'devuelve la cantidad nueva de la lista
    
    If nOrden > UBound(ListaFalta) Then Exit Function
    
    If nOrden = UBound(ListaFalta) Then
        'solo voy a correr el final
        GoTo FIN
    End If
    
    Dim A As Long
    For A = nOrden To UBound(ListaFalta) - 1
        'paso a todos uno mas adelante !!!
        ListaFalta(A) = ListaFalta(A + 1)
        ListaTag(A) = ListaTag(A + 1)
    Next A

FIN:
    Dim NewInd As Long
    NewInd = UBound(ListaFalta) - 1
    
    ReDim Preserve ListaFalta(NewInd)
    ReDim Preserve ListaTag(NewInd)
        
    ListaKillElement = UBound(ListaFalta)
        
    'ver si graba automaticamente
    If mGrabaAuto And mArchivo <> "" Then ListaGuardarADisco mArchivo
    
End Function

Public Function ListaGuardarADisco(sPath As String) As Long
    'guardarla para recuperar en el futuro
    
    'devuelve 0 en ok
    '1 algun error
    
    On Error GoTo errgraba
    
    Dim Te As TextStream
    Dim A As Long
    Set Te = FSO.CreateTextFile(sPath, True) 'lo borra si existe
        For A = 1 To UBound(ListaFalta)
            'con un seprador gueno, gueno
            Te.WriteLine ListaFalta(A) + Chr(5) + ListaTag(A)
        Next A
    Te.Close
    
    ListaGuardarADisco = 0
    
    Exit Function
errgraba:
    ListaGuardarADisco = 1
End Function

Public Function ListaAbrirDeDisco(sPath As String) As Long
    'abrir una lista guardada
    
    'devuelve cantidad de archivos si tod esta ok (puede ser cero en archivo vacio o con todos inexistentes)
    '-1 algun error
    '-2 no hay archivo
    '-3 no era compatible y lo borre
    
    On Error GoTo ErrLee
    
    If FSO.FileExists(sPath) = False Then
        ListaAbrirDeDisco = -2
        Exit Function
    End If
    
    Dim Te As TextStream
    Dim A As Long
    Dim CadaArch As String, CadaTag As String
    Dim SP() As String
    Set Te = FSO.OpenTextFile(sPath, ForReading, False)
        Do While Not Te.AtEndOfStream
            CadaArch = Te.ReadLine
            
            '************************************
            'compatibilidad con archivos viejos!
            If InStr(CadaArch, Chr(5)) = 0 Then
                Te.Close
                'borrarlo y el proximo se grabara en el nuevo formato
                FSO.DeleteFile sPath
                ListaAbrirDeDisco = -3
                Exit Function
            End If
            '************************************
            
            SP = Split(CadaArch, Chr(5))
            CadaArch = SP(0)
            CadaTag = SP(1)
            'verificar la existencioa de todos y limpiarla !!
            If FSO.FileExists(CadaArch) Then
                A = A + 1
                ReDim Preserve ListaFalta(A): ListaFalta(A) = CadaArch
                ReDim Preserve ListaTag(A): ListaTag(A) = CadaTag
            End If
        Loop
    Te.Close
    
    ListaAbrirDeDisco = A
    Exit Function
ErrLee:
    ListaAbrirDeDisco = -1
End Function

Private Sub Class_Initialize()
    ReDim Preserve ListaFalta(0)
End Sub

Public Function GetTag(Index As Long) As String
    If Index > UBound(ListaFalta) Or Index = 0 Then
        GetTag = ""
        Exit Function
    End If
    
    GetTag = ListaTag(Index)
End Function

Public Function GetElementListaPath(Index As Long)
    If Index > UBound(ListaFalta) Or Index = 0 Then
        GetElementListaPath = ""
        Exit Function
    End If
    
    GetElementListaPath = ListaFalta(Index)
End Function

Public Function GetElementListaFileName(Index As Long)
    If Index > UBound(ListaFalta) Or Index = 0 Then
        GetElementListaFileName = ""
        Exit Function
    End If
    
    Dim TP As String
    TP = FSO.GetBaseName(ListaFalta(Index))
    
    'todo menos el largo de la extencion y el ultimo punto
    GetElementListaFileName = Left(TP, Len(TP) - Len(FSO.GetExtensionName(TP)))
    
End Function

Public Function GetElementListaExtencion(Index As Long) As String
    If Index > UBound(ListaFalta) Or Index = 0 Then
        GetElementListaExtencion = ""
        Exit Function
    End If
    
    GetElementListaExtencion = FSO.GetExtensionName(ListaFalta(Index))
End Function

Public Function GetElementListaLastFolder(Index As Long)
    If Index > UBound(ListaFalta) Or Index = 0 Then
        GetElementListaLastFolder = ""
        Exit Function
    End If
    
    Dim TMP() As String
    TMP = Split(ListaFalta(Index), "\")
    '''
    GetElementListaLastFolder = TMP(UBound(TMP) - 1)
End Function

Public Function GetMMType(Index As Long) As MMType
    If Index > UBound(ListaFalta) Or Index = 0 Then
        GetMMType = ""
        Exit Function
    End If
    
    Dim TMP As String
    TMP = LCase(GetElementListaExtencion(Index))
    
    Select Case TMP 'me falta identificar los karaokes MPG !!!
        Case "mp3", "wma", "ogg", "wav"
            GetMMType = mmtypeMUSICA
        Case "wmv", "avi", "mpg"
            GetMMType = mmtypeVIDEO
        Case "mn0", "mn1"
            GetMMType = mmtypeKARAOKEMN1
        Case Else
            GetMMType = mmtypeSinDEFINIR
    End Select
End Function

Public Function GetLastIndex() As Long
    GetLastIndex = UBound(ListaFalta)
End Function

Public Property Let GrabaAuto(nGraba As Boolean)
    mGrabaAuto = nGraba
End Property

Public Property Get GrabaAuto() As Boolean
    GrabaAuto = mGrabaAuto
End Property

Public Property Let Archivo(sArch As String)
    mArchivo = sArch
End Property

Public Property Get Archivo() As String
    Archivo = mArchivo
End Property
