VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Translator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'hay que agregar la referencia a Microsoft Scripting Runtime

Dim vLanguage As String
Dim vDictionary As New Dictionary

Private v(8) As String 'variables metidas
'dentro de las cadenas a traducir

Public Property Get Dictionary() As Dictionary
    Set Dictionary = vDictionary
End Property

Public Property Get Language() As String
    'ver si hay q devolver default en vez de cast.
    Language = IIf(vLanguage <> "", vLanguage, "Castellano")
End Property

Public Property Let Language(pValor As String)
    vLanguage = pValor
    LoadDictionary pValor
End Property

Public Property Get Languages() As String()
    Dim fso
    Dim folder
    Dim aux As String
    
    Set fso = CreateObject("scripting.filesystemobject")
    Set folder = fso.GetFolder(App.path + "\idiomas")
    Dim F
    For Each F In folder.Files
        aux = aux + Left(F.Name, Len(F.Name) - 4) + "*"
    Next
    Languages = Split(Left(aux, Len(aux) - 1), "*")
End Property

Public Sub SetVars(Optional ByVal s1, Optional ByVal s2, Optional ByVal s3, Optional ByVal s4, _
                   Optional ByVal s5, Optional ByVal s6, Optional ByVal s7, Optional ByVal s8)
    
    v(1) = CStr(s1)
    v(2) = CStr(s2)
    v(3) = CStr(s3)
    v(4) = CStr(s4)
    v(5) = CStr(s5)
    v(6) = CStr(s6)
    v(7) = CStr(s7)
    v(8) = CStr(s8)
    
End Sub

Public Property Get Trad(pSpanishWord As String) As String
Attribute Trad.VB_MemberFlags = "200"
    
    'tener en cuenta los parametros opcionales que son variables!!!
    'podria haber puesto args() pero tendria que tener matrices por todos lados
    'ademas de esta forma puedo poner fechas o numeros sin el cstr()
    'el problema para cadenas granbes puede ser la cantidad de variables necesarias
    'para traducciones comunes 8 debe ser mas que suficiente
    
    On Error GoTo e
    
    'sacarlo todos los %01%...%98%....%99%
    Dim pSPA As String
    pSPA = pSpanishWord
    'tERR.Anotar "ACNC20b", pSPA
    If pSPA <> "" Then
        Dim RealText As String, Ayuda As String
        Dim TMP() As String, B As Long, C As Long
        Dim par() As String
'        If InStr(pSPA, "1 can") <> 0 Then
'            C = 1
'        End If
        par = Split(pSPA, Chr$(31))
        'primero dejo solo lo que se va a traducir diferenciandolo de la ayuda
        B = InStr(par(0), "%98%")
        If B = 0 Then
            B = Len(par(0))
            RealText = Left(par(0), B - 4)
        Else
            RealText = Left(par(0), B - 1)
        End If
        'ahora le pongo saltos de carro donde van
        RealText = Replace(RealText, "%97%", vbCrLf)
    End If
        
    'tERR.Anotar "ACNC21", RealText
    
    Dim aux As String
    aux = vDictionary.Item(RealText)
    If aux = "" Then aux = RealText 'si no esta entonces uso el del idioma original
    
    'meter los valores de las variables donde van si es que existen
    'se supone que si existe "%01%" es por que existe v1, caso contrario puede dar errro
    aux = Replace(aux, "%01%", v(1))
    aux = Replace(aux, "%02%", v(2))
    aux = Replace(aux, "%03%", v(3))
    aux = Replace(aux, "%04%", v(4))
    aux = Replace(aux, "%05%", v(5))
    aux = Replace(aux, "%06%", v(6))
    aux = Replace(aux, "%07%", v(7))
    aux = Replace(aux, "%08%", v(8))
    
    Erase v 'limpio la matriz
    'tERR.Anotar "ACNC22", aux
    
    If aux <> "" Then
        Trad = aux
    Else
        'si no encuentra la palabra, devuelve la misma
        Trad = RealText
    End If
    
    Exit Property
e:
    'si no encuentra la palabra, devuelve la misma
    Translate = pSpanishWord
    
End Property

Private Function LoadDictionary(pLanguage As String) As Long

    On Error GoTo e
   
    Dim S As String
    Dim aux() As String
    Dim par() As String
    
    S = LeerArchivo(pLanguage)
    
    aux = Split(S, Chr$(30))
    vDictionary.RemoveAll
    For I = 0 To UBound(aux)
        If aux(I) <> "" Then
            par = Split(aux(I), Chr$(31))
            If vDictionary.Exists(par(0)) = False Then vDictionary.Add par(0), par(1)
        End If
    Next
    
    LoadDictionary = 0
    Exit Function
    
e:
    LoadDictionary = 1
    MsgBox TR.Trad("No se encontro el archivo de idioma.%99%")
End Function

Private Function LeerArchivo(path As String) As String
    Dim fso 'As FileSystemObject
    Dim F
        
    Dim S As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If path = "" Then
        LeerArchivo = ""
    Else
        Set F = fso.GetFile(path)
    End If
    Set ts = F.OpenAsTextStream(1)
    
    S = ts.ReadAll
    
    ts.Close
    LeerArchivo = S
End Function

