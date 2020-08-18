VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   4365
      Left            =   3330
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   180
      Width           =   5535
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   3075
   End
   Begin VB.Menu mnREG 
      Caption         =   "Registro de Fallas"
      Begin VB.Menu mnOpen 
         Caption         =   "Abrir Paquete"
      End
      Begin VB.Menu SEP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnANT 
         Caption         =   "Anteriores"
         Begin VB.Menu mnElija 
            Caption         =   "Elija"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnQuit 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FSO As New Scripting.FileSystemObject
Private AP As String
Dim Fol As String 'carpeta donde esta todo

Private Sub Form_Load()
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    'ver registros abierto anteriormente
    
    Dim F As Folder
    Dim F2 As Folder
    
    Set F = FSO.GetFolder(AP)
    Dim C As Long
    For Each F2 In F.SubFolders
        If Left(F2.Name, 3) = "200" Then
            C = C + 1
            Load mnElija(C)
            mnElija(C).Caption = F2.path
            mnElija(C).Enabled = True
        End If
    Next
End Sub

Private Sub List1_Click()
    ShowT1 Fol + List1
End Sub

Private Sub mnElija_Click(Index As Integer)
    Fol = mnElija(Index).Caption

    sOpen "", Fol
End Sub

Private Sub mnOpen_Click()
    Dim CM As New CommonDialog
    
    CM.Filter = "Archivos de registro empaquetados(*.JSA)|*.JSA"
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    Dim Cliente As String
    Cliente = InputBox("Para que cliente ?")
    
    Cliente = CStr(Year(Date)) + CStr(Month(Date) + 10) + CStr(Day(Date)) + "_" + Cliente
    
    Fol = AP + Cliente
    If FSO.FolderExists(Fol) = False Then FSO.CreateFolder Fol

    sOpen F, Fol
    
End Sub

Private Sub sOpen(Arch As String, Carp As String)

    List1.Clear

    If Arch = "" Then 'es uno que ya estaba descomprimido
        Dim F As File
        Dim FO As Folder
        Set FO = FSO.GetFolder(Carp)
        For Each F In FO.Files
            List1.AddItem F.Name
        Next
    Else
        'Descomprimir lo que llega
        Dim JS As New tbrJUSE.clsJUSE
        JS.ReadFile Arch
        Dim J As Long
        For J = 1 To JS.CantArchs
            JS.Extract Carp, J
            List1.AddItem JS.GetListFiles(J, False)
            If LCase(JS.GetListFiles(J, False)) = "cd4.pm" Then
                TRALIC Fol + "cd4.pm"
                List1.AddItem "cd4.pm4"
            End If
        Next J
    End If
    
    If Right(Fol, 1) <> "\" Then Fol = Fol + "\"
End Sub

Private Sub ShowT1(Arch As String)
    'mostrar como corresponda
    
    If FSO.GetBaseName(Arch) = "jumal" Then
        Text1.Text = tLt(Arch, Chr(5) + Chr(7) + Chr(6) + Chr(4), "Ingrese su pais de residencia")
        Exit Sub
    End If
    
    If FSO.GetBaseName(Arch) = "guen" Then
        Text1.Text = tLt(Arch, Chr(5) + Chr(6) + Chr(6) + Chr(5), "Telefono o fax")
        Exit Sub
    End If
    
    If FSO.GetBaseName(Arch) = "japi" Then
        Text1.Text = tLt(Arch, Chr(7) + Chr(7) + Chr(6) + Chr(5), "Email tecnico")
        Exit Sub
    End If
    
    If FSO.GetBaseName(Arch) = "buca" Then
        Text1.Text = tLt(Arch, Chr(4) + Chr(7) + Chr(6) + Chr(5), "Email administrativo")
        Exit Sub
    End If
    
    If FSO.GetBaseName(Arch) = "buda" Then
        Text1.Text = tLt(Arch, Chr(4) + Chr(6) + Chr(6) + Chr(4), "Gracias por confiar en tbrSoft")
        Exit Sub
    End If
    
    Dim EXT As String
    EXT = LCase(FSO.GetExtensionName(Arch))
    
    Dim TMP As String
    Dim TE As TextStream
    
    Select Case EXT
        Case "ona"
            Set TE = FSO.OpenTextFile(Arch, ForReading, False)
                If TE.AtEndOfStream Then
                    TMP = "Archivo PELADO!!"
                Else
                    TMP = TE.ReadAll
                    TMP = Encriptar(TMP, True)
                End If
            TE.Close
            Text1.Text = TMP
        
        Case Else '"log", "w15", "nga", "pm", "day", "txt", "pm4"
            Set TE = FSO.OpenTextFile(Arch, ForReading, False)
                If TE.AtEndOfStream Then
                    TMP = "Archivo PELADO!!"
                Else
                    TMP = TE.ReadAll
                End If
            TE.Close
            Text1.Text = TMP
    End Select
    
    
'    'ESTOS SON LOS ARCHIVOS QUE PUEDE ABRIR
'    AddFiles App.path, "log" 'REGISTRO BASICO + REGISTRO DE MMPLAYER
'    AddFiles App.path, "w15" 'ARCHIVOS W15
'    AddFile App.path + "\sf\marad.ona" 'CONFIGURACION DE 3PM
'    'OTRAS COSAS INTERESANTES
'    AddFile BasePath + "pindo.nga" 'lista de origenes de discos 'EX: sf+ "oddtb.jut"
'    AddFile BasePath + "cd3.pm" 'Copia clave sf + "c2LK.dll"
'    AddFile BasePath + "cccd3.pm" 'Copia clave sf + "c2LK.dll"
'    AddFile BasePath + "sf\cd4.pm" 'Archivo de licencia 3pm 7.0 (GENERADO)
'    AddFile BasePath + "sf\cd7.pm" 'Archivo RECIBIDO de licencia 3pm 7.0 COREGIDO Y EN USO
'    AddFile BasePath + "sf\rdc.day" 'registro diario del contador sf + "daily.cfg"
'    AddFile BasePath + "daliv.mp2" 'archivo con las claves para validar
    
    
End Sub

Private Function tLt(fil, sep, qi As String) As String
    
    Dim ftLt As String 'temporal para leerlo todo
    ftLt = ""
    LEE = GetText(CStr(fil))
    
    'leer cada dato
    Dim SP() As String
    SP = Split(LEE, sep)
    Dim H As Long
    For H = 0 To UBound(SP)
        ftLt = ftLt + pinchilon(SP(H), qi, True) + vbCrLf 'nuevo separador
    Next H
    
    tLt = ftLt
End Function

'CHOREADO DE 3PM
Public Function Encriptar(Valor, UnEncrypt As Boolean) As String
    'con esta funcion se puede encriptar y desencriptar
    'la uso para el GPF("config")
    
    'para saber si estoy leyendo algo encrytado le pongo algo identificativo
    Dim IdEstaEncryptado As String
    IdEstaEncryptado = "RMLVF"
    'encripta cualquier cosa y la transforma en string
    Dim ToEncrypt As String
    ToEncrypt = CStr(Valor)
    
    Dim Largo As Long, IND As Long, Letra As String, LetraE As String
    Dim FullE As String 'resultado de la encryptacion
    'ver si lo que se ingreso ya esta encrptado
    If UCase(Left(ToEncrypt, Len(IdEstaEncryptado))) = IdEstaEncryptado Then
        'ya esta encriptado
        If UnEncrypt Then
            'DESNCRIPTAR!!!
            'cambiar uno por uno los codigos
            Largo = Len(ToEncrypt)
            'empeiza despues del marcador
            For IND = Len(IdEstaEncryptado) + 1 To Largo
                Letra = Mid(ToEncrypt, IND, 1)
                'pasar todo a una letra distinta. Los saltos de carro no usarlos
                If Asc(Letra) = 0 Then Letra = "0"
                Select Case Letra
                    Case "0"
                        LetraE = vbCrLf
                    Case Else
                        LetraE = Chr(Asc(Letra) - 10)
                End Select
                FullE = FullE + LetraE
            Next
            Encriptar = FullE
        Else
            'no se puede encyprtar lo encryptado
            Encriptar = ToEncrypt
            Exit Function
        End If
    Else
        If UnEncrypt Then
            'no se puede desdencryptar lo desencryptado
            Encriptar = ToEncrypt
            Exit Function
        Else
            'Encriptar!!!!
            'cambiar uno por uno los codigos
            Largo = Len(ToEncrypt)
            For IND = 1 To Largo
                Letra = Mid(ToEncrypt, IND, 1)
                'pasar todo a una letra distinta. Los saltos de carro no usarlos
                Select Case Letra
                    Case vbCrLf ' Or vbCr
                        LetraE = "0"
                    Case Else
                        LetraE = Chr(Asc(Letra) + 10)
                End Select
                FullE = FullE + LetraE
            Next
            Encriptar = IdEstaEncryptado + FullE
        End If
        
    End If
    
End Function

Private Sub mnQuit_Click()
    Unload Me
End Sub

Private Sub TRALIC(Arch As String)
    Dim LL As New tbrDATA.clsTODO

    If LL.LiToLo(Arch, Arch + "4") = 1 Then
        Text1.Text = "No es un archivo de 3PM"
    End If
    
End Sub

Private Function GetText(F As String) As String

    If FSO.FileExists(F) = False Then
        GetText = "boloooo"
        Exit Function
    End If
    
    Dim TE As TextStream
    Set TE = FSO.OpenTextFile(F)
        If TE.AtEndOfStream = False Then
            GetText = TE.ReadAll
        Else
            GetText = "RE-boloooo"
        End If
    TE.Close
End Function

Private Function pinchilon(Texto As String, Clave As String, Invertido As Boolean) As String
    'encriptar
    'Cargo los datos
    
    If Texto = "" Then
        pinchilon = ""
        Exit Function
    End If
    
    Dim F As Integer
    Dim Buffer() As Byte
    'Buffer = Texto 'se meten de dos en dos las letras ??? sera por algo de ascii vs unicode
    
    ReDim Buffer(Len(Texto) - 1)
    For F = 1 To Len(Texto)
        Buffer(F - 1) = Asc(Mid(Texto, F, 1))
    Next F

    Dim xClave() As Byte
    'xClave = Clave
    
    ReDim xClave(Len(Clave) - 1)
    For F = 1 To Len(Clave)
        xClave(F - 1) = Asc(Mid(Clave, F, 1))
    Next F
    
    'Encripto
    
    Dim Char1 As Integer 'Caracter Original
    Dim Char2 As Integer 'Caracter ya Modificado (char1+char3) o (char1-char3)
    Dim Char3 As Integer 'Caracter de la Clave
    
    'Voy dando vueltas por la clave asi que necesito un indice
    Dim ContadorClave As Integer 'Indice de la clave
    ContadorClave = 0
    
    Dim i As Long
    Dim NuevoDato() As Byte
    
    ReDim NuevoDato(Len(Texto) - 1)
    
    For i = 0 To UBound(Buffer)
        Char1 = Buffer(i)
        Char3 = xClave(ContadorClave)
        If Invertido = True Then
            Char2 = Char1 - Char3
        Else
            Char2 = Char1 + Char3
        End If

        If Char2 < 0 Then
            Char2 = 256 + Char2
        End If
    
        If Char2 > 255 Then
            Char2 = Char2 Mod 256
        End If
    
        NuevoDato(i) = Char2
        
        ContadorClave = ContadorClave + 1
        If ContadorClave > UBound(xClave) Then ContadorClave = 0
    Next i
    
    Dim tRES As String
    For F = 0 To UBound(NuevoDato)
        tRES = tRES + Chr(NuevoDato(F))
    Next F
    
'    Dim Ver As String
'    For F = 0 To UBound(Buffer)
'        Ver = Ver + Chr(Buffer(F)) + " - " + Chr(NuevoDato(F)) + " * "
'    Next F
'    MsgBox Ver

    pinchilon = tRES
    
End Function
