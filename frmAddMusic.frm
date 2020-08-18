VERSION 5.00
Begin VB.Form frmAddMusic 
   BackColor       =   &H00004080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar musica a 3PM"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBar2 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   5670
      ScaleHeight     =   165
      ScaleWidth      =   15
      TabIndex        =   18
      Top             =   6660
      Width           =   15
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Agregar estas carpetas a 3PM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6810
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5610
      Width           =   4050
   End
   Begin VB.ListBox lstCarConMM 
      Height          =   3210
      Left            =   5640
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   810
      Width           =   6165
   End
   Begin VB.PictureBox PBar 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   5640
      ScaleHeight     =   165
      ScaleWidth      =   15
      TabIndex        =   9
      Top             =   4110
      Width           =   15
   End
   Begin VB.PictureBox P 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5640
      ScaleHeight     =   135
      ScaleWidth      =   6165
      TabIndex        =   8
      Top             =   4080
      Width           =   6225
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   660
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6810
      Width           =   3780
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Agregar Origen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Explorar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1650
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   630
      Width           =   1500
   End
   Begin VB.PictureBox P2 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5670
      ScaleHeight     =   135
      ScaleWidth      =   6165
      TabIndex        =   19
      Top             =   6630
      Width           =   6225
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3°"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   645
      Index           =   6
      Left            =   6210
      TabIndex        =   23
      Top             =   5520
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2°"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   645
      Index           =   5
      Left            =   5670
      TabIndex        =   22
      Top             =   120
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1°"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   645
      Index           =   4
      Left            =   90
      TabIndex        =   21
      Top             =   -60
      Width           =   525
   End
   Begin VB.Label lblBAR2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sin Tareas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   825
      Left            =   5670
      TabIndex        =   20
      Top             =   6870
      Width           =   5985
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAddMusic.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   765
      Index           =   1
      Left            =   5610
      TabIndex        =   17
      Top             =   4920
      Width           =   6225
   End
   Begin VB.Label lblWait 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Analizando disco.  Espere..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1725
      Left            =   180
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   5145
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Las carpetas elegidas se cargaran en el origen de datos original de 3PM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   525
      Index           =   0
      Left            =   7110
      TabIndex        =   15
      Top             =   6030
      Width           =   3705
   End
   Begin VB.Label lblTOT 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total de carpetas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   5760
      TabIndex        =   13
      Top             =   30
      Width           =   5985
   End
   Begin VB.Label lblBAR 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sin Tareas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   5640
      TabIndex        =   12
      Top             =   4290
      Width           =   5985
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   5490
      X2              =   5490
      Y1              =   120
      Y2              =   7080
   End
   Begin VB.Label Ltit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Carpetas encontradas con multimedia: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   5730
      TabIndex        =   10
      Top             =   510
      Width           =   5985
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Especificar ubicacion de los nuevos discos a agregar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Index           =   3
      Left            =   720
      TabIndex        =   7
      Top             =   90
      Width           =   4725
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   150
      X2              =   5280
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dejar como permanente un nuevo origen de datos multimedia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Index           =   2
      Left            =   1830
      TabIndex        =   5
      Top             =   2850
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Explorar para buscar los nuevos temas. Usese para buscar desde red u otros discos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Index           =   1
      Left            =   1830
      TabIndex        =   4
      Top             =   1650
      Width           =   3405
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tomar los nuevos temas desde un CD. 3PM busca automaticamente en todas las carpetas del CD insertado."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Index           =   0
      Left            =   1860
      TabIndex        =   3
      Top             =   600
      Width           =   3405
   End
End
Attribute VB_Name = "frmAddMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalArchMM As Long 'total de archivos multimedia en todas las carpetas
Dim CarpsConMM() As String
Dim X As New CommonDialog
Dim CarpetaDesdeCargar As String

Private Sub Command1_Click()
    X.CancelError = False
    X.InitDir = "" 'para que muestre todo
    X.DialogPrompt = "Elegir carpeta contenedora de los nuevos discos"
    X.ShowFolder
    
    If Len(X.InitDir) Then
        CarpetaDesdeCargar = X.InitDir
        lblWait.Visible = True
        lblWait.Refresh
        'buscar carpetas de multimedia
        CarpsConMM = FindCarpsConMM(CarpetaDesdeCargar)
        lblTOT = "Carpetas encontradas el la ubicacion elegida: " + CStr(UBound(CarpsConMM))
        lblTOT.Refresh
        'ver cuales tienen multimedia
        BuscarCarpetasMM
    End If
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub
Public Sub BuscarCarpetasMM()
    TotalArchMM = 0 'inicializar el contador de archivosmm(no carpetas)
    Dim TotF As Long
    TotF = UBound(CarpsConMM)
    PBar.Width = 0
    Dim TotCarpMM As Long
    Dim TMPfilesMM() As String
    lstCarConMM.Clear
    TotCarpMM = 0
    For a = 1 To TotF
        lblBAR = "Buscando en " + CarpsConMM(a)
        PBar.Width = P.Width / TotF * a
        PBar.Refresh
        TMPfilesMM = ObtenerArchMM(CarpsConMM(a))
        If UBound(TMPfilesMM) > 0 Then
            TotalArchMM = TotalArchMM + UBound(TMPfilesMM)
            lstCarConMM.AddItem CarpsConMM(a) + ", " + CStr(UBound(TMPfilesMM)) + " archivos"
            lstCarConMM.Selected(TotCarpMM) = True
            TotCarpMM = TotCarpMM + 1
            Ltit = "Carpetas multimedia encontradas: " + CStr(TotCarpMM)
            Ltit.Refresh
        End If
    Next
    lblWait.Visible = False
    lblWait.Refresh
    lblBAR = "Sin Tareas"
    PBar.Width = 0
    If TotCarpMM = 0 Then
        MsgBox "No se han encontrado carpetas multimedia en la ubicacion elegida"
    Else
        MsgBox "Se han encontrado " + CStr(TotCarpMM) + " carpetas multimedia en la ubicacion elegida"
    End If
End Sub

Private Sub Command4_Click()
    On Error GoTo LogERROR
    'TotArchMM sabe cuantos temas hay en total
    
    'grabar en AP+"discos" los nuevos datos multimedia
    Dim TMPfiles() As String
    Dim SoloCarp As String, NewCarp As String
    Dim PathArch As String, SoloArch As String
    Dim Ubic As String, TotMM As Long
    Dim ArchCopiados As Long
    ArchCopiados = 0
    'ver cuantos archivos efectivamente se copiaran
    Dim TotalACopiar As Long 'no cuenta los que no son multimedia
    TotalACopiar = 0
    For a = 0 To lstCarConMM.ListCount - 1
        If lstCarConMM.Selected(a) Then
            TotMM = Val(txtInLista(lstCarConMM.List(a), 1, ","))
            TotalACopiar = TotalACopiar + TotMM
        End If
    Next
    For a = 0 To lstCarConMM.ListCount - 1
        If lstCarConMM.Selected(a) Then
            TotMM = Val(txtInLista(lstCarConMM.List(a), 1, ","))
            Ubic = txtInLista(lstCarConMM.List(a), 0, ",")
            If Right(Ubic, 1) <> "\" Then Ubic = Ubic + "\"
            'hay que copiar solo los archivos MM
            SoloCarp = txtInLista(Ubic, 99998, "\") '99998 es el anteultimo
            NewCarp = AP + "discos\" + SoloCarp + "\"
            'crear la carpeta si no esta
            If FSO.FolderExists(NewCarp) = False Then FSO.CreateFolder NewCarp
            
            'NO OLVIDARSE DE TAPA.JPG Y DATA.TXT
            Dim ArchTapa As String
            ArchTapa = Ubic + "tapa.jpg"
            If FSO.FileExists(ArchTapa) Then
                'si existe ver los atributos
                If FSO.FileExists(NewCarp + "tapa.jpg") Then
                    aaa = GetAttr(NewCarp + "tapa.jpg")
                    If aaa = vbHidden Or aaa = vbReadOnly Then SetAttr NewCarp + "tapa.jpg", 0
                End If
                FSO.CopyFile ArchTapa, NewCarp + "tapa.jpg"
            End If
            
            Dim ArchDaTa As String
            ArchDaTa = Ubic + "data.txt"
            If FSO.FileExists(ArchDaTa) Then
                'si existe ver los atributos
                If FSO.FileExists(NewCarp + "data.txt") Then
                    aaa = GetAttr(NewCarp + "data.txt")
                    If aaa = vbHidden Or aaa = vbReadOnly Then SetAttr NewCarp + "data.txt", 0
                End If
                FSO.CopyFile ArchDaTa, NewCarp + "data.txt"
            End If
            
            TMPfiles = ObtenerArchMM(Ubic) 'deveuelve pathfull , solonombre
            c = 1
            Do While c <= TotMM 'se supone que es el total de esta carpeta
                PathArch = txtInLista(TMPfiles(c), 0, ",")
                SoloArch = txtInLista(TMPfiles(c), 1, ",")
                lblBAR2 = "Copiando " + PathArch
                lblBAR2.Refresh
                ArchCopiados = ArchCopiados + 1
                PBar2.Width = P2.Width / TotalACopiar * ArchCopiados
                PBar2.Refresh
                'si existe ver los atributos
                If FSO.FileExists(NewCarp + SoloArch) Then
                    aaa = GetAttr(NewCarp + SoloArch)
                    If aaa = vbHidden Or aaa = vbReadOnly Then SetAttr NewCarp + SoloArch, 0
                End If
                FSO.CopyFile PathArch, NewCarp + SoloArch, True
                c = c + 1
            Loop
            lblBAR2 = "Sin Tareas"
            PBar2.Width = 0
            
        End If
    Next
    MsgBox "Los archivos se copiaron correctamente"
    Exit Sub
LogERROR:
    WriteTBRLog "Error al cargar archivos MM. n° " + CStr(Err.Description) + " Descr: " + Err.Description, True
    Resume Next
End Sub

Private Sub Command5_Click()
    Dim DRs As Drives, DS As Drive
    Set DRs = FSO.Drives
    CarpetaDesdeCargar = "NO"
    For Each DS In DRs
        If DS.DriveType = 4 Then '4 es CDROM
            If DS.IsReady Then
                CarpetaDesdeCargar = DS.DriveLetter + ":\"
            Else
                MsgBox "El disco " + DS.DriveLetter + " no esta " + _
                "listo. Inserte un CD y reintente"
                Exit Sub
            End If
        End If
    Next
    If CarpetaDesdeCargar = "NO" Then
        MsgBox "No se encontro unidad de CD"
        Exit Sub
    End If
    lblWait.Visible = True
    lblWait.Refresh
    'buscar carpetas de multimedia
    CarpsConMM = FindCarpsConMM(CarpetaDesdeCargar)
    lblTOT = "Carpetas encontradas el la ubicacion elegida: " + CStr(UBound(CarpsConMM))
    lblTOT.Refresh
    'ver cuales tienen multimedia
    BuscarCarpetasMM
    
End Sub

Public Function FindCarpsConMM(Carp As String) As String()
    'devuelve una matriz con todas las carpetas que tengan multimedia
    Dim CarpetasEnQueBuscar() As String
    Dim Nivel2() As String
    Dim TodasLasCarpetas() As String
    Dim LastIni As Long, LastFin As Long
    CarpetasEnQueBuscar = GetFolders(Carp)
    LastIni = 1
    LastFin = UBound(CarpetasEnQueBuscar)
    Dim AgregadosEnVuelta
    Dim ContTotal As Long
    Do
        AgregadosEnVuelta = 0
        For a = LastIni To LastFin
            ContTotal = ContTotal + 1
            ReDim Preserve TodasLasCarpetas(ContTotal)
            TodasLasCarpetas(ContTotal) = CarpetasEnQueBuscar(a)
            Nivel2 = GetFolders(CarpetasEnQueBuscar(a)) 'el cero no existe
            If UBound(Nivel2) > 0 Then
                Dim LastCBuscar
                LastCBuscar = UBound(CarpetasEnQueBuscar)
                For Z = 1 To UBound(Nivel2)
                    ReDim Preserve CarpetasEnQueBuscar(LastCBuscar + Z)
                    CarpetasEnQueBuscar(LastCBuscar + Z) = Nivel2(Z)
                    AgregadosEnVuelta = AgregadosEnVuelta + 1
                Next
                GoTo NextMM
            Else
                If LastFin = UBound(CarpetasEnQueBuscar) Then
                    If a = LastFin Then
                        'no tiene y es el ultimo
                        Exit Do
                    Else
                        GoTo NextMM
                    End If
                End If
            End If
NextMM:
        Next
        LastIni = UBound(CarpetasEnQueBuscar) - AgregadosEnVuelta + 1
        LastFin = LastIni + AgregadosEnVuelta - 1
    Loop
    FindCarpsConMM = TodasLasCarpetas
End Function

' Devuelve un array de cadenas que incluye todos los subdirectorios
' contenidos en una ruta que coincide con los atributos de búsqueda
' opcionalmente, devuelve la ruta completa.

Function GetFolders(ruta As String) As String()
        Dim Resultado() As String
        Dim NombreDir As String, CONTADOR As Long
        Dim Ruta2 As String
        ReDim Resultado(0) As String
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
                CONTADOR = CONTADOR + 1
                ReDim Preserve Resultado(CONTADOR) As String
                ' incluir la ruta si se pide
                NombreDir = Ruta2 & NombreDir
                Resultado(CONTADOR) = NombreDir
            End If
            NombreDir = Dir$
        Loop
        
        ' proporciona el array resultante
        ReDim Preserve Resultado(CONTADOR) As String
        GetFolders = Resultado
End Function

Sub ShowDriveList()
    On Local Error Resume Next
    Dim fs, d, dc, s, n
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set dc = fs.Drives
    For Each d In dc
        s = s & d.DriveLetter & " - "
        Select Case d.DriveType
            Case 0: t = "Desconocido"
            Case 1: t = "Separable"
            Case 2: t = "Fijo"
            Case 3: t = "Red"
            Case 4: t = "CD-ROM"
            Case 5: t = "Disco RAM"
        End Select
        If d.DriveType = 3 Then
            n = d.ShareName
        Else
            n = d.VolumeName
        End If
        s = s & n & "Tipo: " & t & vbCrLf
    Next
    MsgBox s
End Sub

Private Sub Form_Load()
    AjustarFRM Me, 12000
End Sub
