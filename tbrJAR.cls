VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrJAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clase para leer datos de un archivo jar

Private mFile As String 'archivo java abierto

'informacion de los manifest
Private Type PROP
    pName As String
    pValue As String
End Type

Private PROPs() As PROP

Private ManifPath As String 'path completo del archivo de manifiesto

Private Sub Class_Initialize()
    ReDim PROPs(0)
End Sub

Public Function Open_Jar(sFile As String) As Long
    If fso.FileExists(sFile) Then
        mFile = sFile
    Else
        Open_Jar = 1 'error de falta de archivo
        Exit Function
    End If
    
    On Local Error Resume Next
    
    'si ya hay un manifiesto borrarlo! (siemrpe usa la carpeta actual y es un lio cambiarla o saber cual es
    If fso.FileExists("META-INF\MANIFEST.MF") Then
        fso.DeleteFile "META-INF\MANIFEST.MF", True
    End If
    
    'imagino que al crear la carpeta esta es la ubicacion de la carpeta activa
    Dim res As Long, comEx As String 'respuesta y comando a ejecutar
    'comEx = LCase("cd /D " + Chr(34) + Fol + Chr(34)) 'el /D cambia tambien la unidad ademas de la carpeta se asegura la posicion alli
    
    comEx = "jar xf " + Chr(34) + mFile + Chr(34) + " META-INF\MANIFEST.MF *.jpg"
    res = Shell(comEx, vbHide) 'extraer solo el archivo de información , el vbHide hara que no se vea la ventana de DOS
    
    EsperarSec 1 'esperar a que se extraiga!
    
    ManifPath = fso.GetParentFolderName(sFile) + "\" + fso.GetBaseName(sFile) + ".MF"
    
    fso.CopyFile "META-INF\MANIFEST.MF", ManifPath 'si hay error lo saltea, parece que fileExists no labura sin path completo. Por eso me tiro el lanze de copiarlo
    
    'supuestamente si todo salio ok existira el archivo manifiesto ya descomprimido
    
    If fso.FileExists(ManifPath) = False Then
        ManifPath = ""
        Open_Jar = 2 'no se extrajo, puede ser que JAR no este en el path !!!
        Exit Function
    Else
        Open_Manifest ManifPath
        'solo para leerlos
        EscribirArch1Linea2 ManifPath, TraslateManifest
    End If
    
    Open_Jar = 0
End Function

'carga las variables segun un archivo de manifiesto directamente
Public Function Open_Manifest(sFile As String) As String
    'devuelve todo el archivo o vacio si no esta
    
    'ya se asegura antes que exista! pero ...
    If fso.FileExists(sFile) = False Then
        Open_Manifest = ""
        Exit Function
    End If
    
    Dim ManifInfo As String
    Dim TE As TextStream
    Set TE = fso.OpenTextFile(sFile, ForReading, False)
        ' ya aprecieron algunos dañados!
        If TE.AtEndOfStream Then
            ManifInfo = "Fichero dañado"
        Else
            ManifInfo = TE.ReadAll
        End If
    TE.Close
    
    Open_Manifest = ManifInfo
    
    Dim SP() As String 'cada uno de los renglones
    SP = Split(ManifInfo, vbCrLf) 'no estoy seguro que sea siempre este el separador de linea xxxx
    'el separadormde linea a vees no es ese (esos son mis amigos de linux haciendose odiar)
    
    If UBound(SP) < 2 Then
        SP = Split(ManifInfo, vbCr)
        If UBound(SP) < 2 Then
            SP = Split(ManifInfo, vbLf)
        End If
    End If
    
    
    Dim H As Long, SP2() As String 'dentro de cada renglon se separa con ":"
    Dim PropName As String, PropValue As String
    For H = 0 To UBound(SP)
        If InStr(SP(H), ":") Then 'evitar renglones vacios
            SP2 = Split(SP(H), ":")
            PropName = Trim(SP2(0))
            PropValue = Trim(Mid(SP(H), Len(SP2(0)) + 2, Len(SP(H)) - Len(SP2(0)) - 1)) 'el 1 es por ":"
            
            AddProp PropName, PropValue
        End If
    Next H
    
End Function

Public Function getStringManifest(sFileJar As String)
    'si ya esta creado lo devuelvo si no lo creo
    If ManifPath = "" Then
        Open_Jar sFileJar
    End If
    
    If ManifPath = "" Then 'no lo pudo abrir o encontrar!
        getStringManifest = ""
    Else
        getStringManifest = Open_Manifest(ManifPath)
    End If
    
End Function

Private Sub AddProp(newName As String, newValue As String)
    Dim J As Long
    
    J = UBound(PROPs) + 1
    ReDim Preserve PROPs(J)
    PROPs(J).pName = newName
    PROPs(J).pValue = newValue
End Sub

'traduce todo el chorizo en la info que me importa!
Public Function TraslateManifest() As String
    Dim H As Long, RET As String
    RET = ""
    For H = 0 To UBound(PROPs)
        Select Case LCase(PROPs(H).pName)
            'nomobre del juego
            Case LCase("MIDlet-Name"): RET = RET + "Nombre: " + PROPs(H).pValue + vbCrLf
            'version del juego
            Case LCase("MIDlet-Version"): RET = RET + "Versión: " + PROPs(H).pValue + vbCrLf
            'CLDC - 'dice CLDC 1.0 por ejemplo
            Case LCase("MicroEdition -Configuration"): RET = RET + PROPs(H).pValue + vbCrLf
            Case LCase("MicroEdition-Configuration"): RET = RET + PROPs(H).pValue + vbCrLf 'dice CLDC 1.0 por ejemplo
            'perfil MIDP
            Case LCase("MicroEdition -Profile"): RET = RET + PROPs(H).pValue + vbCrLf
            Case LCase("MicroEdition-Profile"): RET = RET + PROPs(H).pValue + vbCrLf
            'desacripcion
            Case LCase("MIDlet-Description"): RET = RET + "Descripcion: " + PROPs(H).pValue + vbCrLf
        End Select
    Next H
    
    TraslateManifest = RET
End Function

'////////////////////////////////////////////////
'ejemplo de manifest
'Manifest-Version: 1.0
'Ant-Version: Apache Ant 1.6.5
'Created-By: 1.4.2-b28 (Sun Microsystems Inc.)
'MIDlet-Name: Alien Quadrant 3D
'MIDlet -Vendor: Sense
'MIDlet-Version: 1.0.3
'MIDlet-1: Alien Quadrant 3D, /_.png, _
'MIDlet-Icon: /_.png
'MIDlet-Data-Size: 256
'MicroEdition -Configuration: CLDC -1.1
'MicroEdition -Profile: MIDP -2#
'MIDlet-Info-URL: http://my-cx65.net.ru
'MIDlet-Delete-Confirm: New Games & Applications For Siemens phones. More Info: http://my-cx65.net.ru

'algunas explicaciones
'# MIDlet-Name.Specifies the name of the application that will be shown to the user.
'# MIDlet-Jar-URL.The URL from where the JAR file can be downloaded.
'# MIDlet-Version.The MIDlet version.
'# MIDlet-Vendor.The MIDlet vendor.
'# MIDlet-Jar-Size.The size of the JAR file
'
'Optional attributes for JAD files:
'
'# MIDlet-Data-Size. The minimum number of bytes of persistent data required by the MIDlet. The device should ensure this memory to the application to be run, whereas it can provide optionally more.
'# MIDlet-Description.A brief description for the user.
'# MIDlet-Info-URL.A link to the application home page.
'# An icon that (depending on the particular device) will be associated with the application. The PNG image file format is used.
'# Application-specific attributes, used to configure the MIDlet via its JAD file. For example, the following line in a JAD file would indicate to the MIDlet to launch the configuration property:
