VERSION 5.00
Begin VB.Form frmAddMusic 
   BackColor       =   &H00004080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar musica a 3PM"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00004080&
      Caption         =   "Origenes disponibles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   8070
      TabIndex        =   24
      Top             =   1350
      Width           =   3645
      Begin VB.TextBox txtInfoOrig 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1830
         Width           =   3375
      End
      Begin VB.ListBox lstOrigenes 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         IntegralHeight  =   0   'False
         Left            =   150
         TabIndex        =   25
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CD Audio"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   50
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   480
      Width           =   1200
   End
   Begin VB.PictureBox PBar2 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   8010
      ScaleHeight     =   165
      ScaleWidth      =   15
      TabIndex        =   14
      Top             =   6720
      Width           =   15
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Agregar estos discos a mi fonola"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7980
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6090
      Width           =   3870
   End
   Begin VB.ListBox lstCarConMM 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5010
      Left            =   150
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   2280
      Width           =   7725
   End
   Begin VB.PictureBox PBar 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   150
      ScaleHeight     =   165
      ScaleWidth      =   15
      TabIndex        =   7
      Top             =   7320
      Width           =   15
   End
   Begin VB.PictureBox P 
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   150
      ScaleHeight     =   135
      ScaleWidth      =   7665
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   7725
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8040
      Width           =   3720
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Explorar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   50
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1200
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CD/DVD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   50
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   900
      Width           =   1200
   End
   Begin VB.PictureBox P2 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7980
      ScaleHeight     =   135
      ScaleWidth      =   3765
      TabIndex        =   15
      Top             =   6690
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.Label lblP 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "% libre"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   3990
      TabIndex        =   27
      Top             =   7890
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAddMusic.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1185
      Index           =   0
      Left            =   8430
      TabIndex        =   23
      Top             =   60
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3°"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   555
      Index           =   2
      Left            =   7920
      TabIndex        =   22
      Top             =   120
      Width           =   525
   End
   Begin VB.Label lblWait 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Analizando disco.  Espere..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   660
      TabIndex        =   12
      Top             =   1830
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde aqui podra trandsformar un CD de audio en ficheros mp3."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   1300
      TabIndex        =   21
      Top             =   570
      Width           =   7635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4°"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   555
      Index           =   6
      Left            =   7890
      TabIndex        =   19
      Top             =   4260
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2°"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   645
      Index           =   5
      Left            =   90
      TabIndex        =   18
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1°"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   4
      Left            =   60
      TabIndex        =   17
      Top             =   0
      Width           =   525
   End
   Begin VB.Label lblBAR2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sin Tareas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   1065
      Left            =   7980
      TabIndex        =   16
      Top             =   6930
      Width           =   3630
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAddMusic.frx":00A3
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1605
      Index           =   1
      Left            =   8400
      TabIndex        =   13
      Top             =   4290
      Width           =   3315
   End
   Begin VB.Label lblBAR 
      BackStyle       =   0  'Transparent
      Caption         =   "Sin Tareas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   150
      TabIndex        =   10
      Top             =   7530
      Width           =   7815
   End
   Begin VB.Label Ltit 
      BackStyle       =   0  'Transparent
      Caption         =   "Carpetas encontradas con multimedia: 0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   690
      TabIndex        =   8
      Top             =   1950
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Especificar ubicacion de los nuevos discos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   5685
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   60
      X2              =   7950
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Explore usted por nuevos discos. Ususe para discos duros o unidades de red."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Index           =   1
      Left            =   1300
      TabIndex        =   3
      Top             =   1290
      Width           =   6435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "3PM busca automaticamente en todas las carpetas del CD insertado."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   0
      Left            =   1300
      TabIndex        =   2
      Top             =   870
      Width           =   6405
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
    On Error GoTo MiErr
    X.CancelError = False
    X.InitDir = "" 'para que muestre todo
    tERR.Anotar "achx", IDIOMA
    Select Case IDIOMA
        Case "Español"
            X.DialogPrompt = "Elegir carpeta contenedora de los nuevos discos"
        Case "English"
            X.DialogPrompt = "Eligeu carpetau"
        Case "Francois"
        Case "Italiano"
    End Select
    
    X.ShowFolder
    
    If Len(X.InitDir) Then
        CarpetaDesdeCargar = X.InitDir
        tERR.Anotar "acig", CarpetaDesdeCargar
        lblWait.Visible = True
        lblWait.Refresh
        'buscar carpetas de multimedia
        CarpsConMM = FindCarpsConMM(CarpetaDesdeCargar)
        'ver cuales tienen multimedia
        BuscarCarpetasMM
    End If
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".achy"
    Resume Next
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub
Public Sub BuscarCarpetasMM()
    On Error GoTo MiErr
    
    TotalArchMM = 0 'inicializar el contador de archivosmm(no carpetas)
    Dim TotF As Long
    TotF = UBound(CarpsConMM)
    PBar.Width = 0
    Dim TotCarpMM As Long
    Dim TMPfilesMM() As String
    lstCarConMM.Clear
    TotCarpMM = 0
    tERR.Anotar "acih", TotF
    For A = 1 To TotF
        Select Case IDIOMA
            Case "Español"
                lblBAR = "Buscando en " + CarpsConMM(A)
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
        
        PBar.Width = P.Width / TotF * A
        PBar.Refresh
        TMPfilesMM = ObtenerArchMM(CarpsConMM(A))
        If UBound(TMPfilesMM) > 0 Then
            TotalArchMM = TotalArchMM + UBound(TMPfilesMM)
            Select Case IDIOMA
                Case "Español"
                    lstCarConMM.AddItem CarpsConMM(A) + "# " + CStr(UBound(TMPfilesMM)) + " archivos"
                Case "English"
                    lstCarConMM.AddItem CarpsConMM(A) + "# " + CStr(UBound(TMPfilesMM)) + " files"
                Case "Francois"
                Case "Italiano"
            End Select
            
            lstCarConMM.Selected(TotCarpMM) = True
            TotCarpMM = TotCarpMM + 1
            Select Case IDIOMA
                Case "Español"
                    Ltit = "Carpetas multimedia encontradas: " + CStr(TotCarpMM)
                Case "English"
                    Ltit = "Multimedia Folders Find: " + CStr(TotCarpMM)
                Case "Francois"
                Case "Italiano"
            End Select
            
            Ltit.Refresh
        End If
    Next
    lblWait.Visible = False
    lblWait.Refresh
    Select Case IDIOMA
        Case "Español"
            lblBAR = "Sin Tareas"
        Case "English"
            lblBAR = "Without Work"
        Case "Francois"
        Case "Italiano"
    End Select
    
    PBar.Width = 0
    If TotCarpMM = 0 Then
        Select Case IDIOMA
            Case "Español"
                MsgBox "No se han encontrado carpetas multimedia en la ubicacion elegida"
            Case "English"
                MsgBox "Nou se encontraron"
            Case "Francois"
            Case "Italiano"
        End Select
        
    Else
        Select Case IDIOMA
            Case "Español"
                MsgBox "Se han encontrado " + CStr(TotCarpMM) + " carpetas multimedia en la ubicacion elegida"
            Case "English"
                MsgBox "Has Been Find " + CStr(TotCarpMM) + " multimedia folders in this ubication"
            Case "Francois"
            Case "Italiano"
        End Select
    End If
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".achz"
    Resume Next
End Sub

Private Sub Command4_Click()
    On Error GoTo MiErr
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
    tERR.Anotar "acij", lstCarConMM.ListCount
    For A = 0 To lstCarConMM.ListCount - 1
        If lstCarConMM.Selected(A) Then
            TotMM = Val(txtInLista(lstCarConMM.List(A), 1, "#"))
            TotalACopiar = TotalACopiar + TotMM
            tERR.Anotar "acik", A, lstCarConMM.List(A)
        End If
    Next
    
    For A = 0 To lstCarConMM.ListCount - 1
        If lstCarConMM.Selected(A) Then
            TotMM = Val(txtInLista(lstCarConMM.List(A), 1, "#"))
            'ubic es la ubicacion en origen
            Ubic = txtInLista(lstCarConMM.List(A), 0, "#")
            If Right(Ubic, 1) <> "\" Then Ubic = Ubic + "\"
            'hay que copiar solo los archivos MM
            SoloCarp = txtInLista(Ubic, 99998, "\") '99998 es el anteultimo
            'ver a donde lo va a grabar!
            If Right(lstOrigenes, 1) <> "\" Then
                NewCarp = lstOrigenes + "\" + SoloCarp + "\"
            Else
                NewCarp = lstOrigenes + SoloCarp + "\"
            End If
            'antes siempre copiaba al unico origen existente!
            'NewCarp = AP + "discos\" + SoloCarp + "\"
            
            'crear la carpeta si no esta
            NewCarp = Replace(NewCarp, ",", "")
            tERR.Anotar "acil", A, NewCarp, TotMM
            If FSO.FolderExists(NewCarp) = False Then FSO.CreateFolder NewCarp
            
            'NO OLVIDARSE DE TAPA.JPG Y DATA.TXT
            Dim ArchTapa As String
            ArchTapa = Ubic + "tapa.jpg"
            If FSO.FileExists(ArchTapa) Then
                'si existe ver los atributos
                If FSO.FileExists(NewCarp + "tapa.jpg") Then
                    AAA = GetAttr(NewCarp + "tapa.jpg")
                    tERR.Anotar "acim", AAA
                    If AAA = vbHidden Or AAA = vbReadOnly Then SetAttr NewCarp + "tapa.jpg", 0
                End If
                tERR.Anotar "acin", ArchTapa, NewCarp
                FSO.CopyFile ArchTapa, NewCarp + "tapa.jpg"
            End If
            
            Dim ArchDaTa As String
            ArchDaTa = Ubic + "data.txt"
            If FSO.FileExists(ArchDaTa) Then
                tERR.Anotar "acio"
                'si existe ver los atributos
                If FSO.FileExists(NewCarp + "data.txt") Then
                    AAA = GetAttr(NewCarp + "data.txt")
                    If AAA = vbHidden Or AAA = vbReadOnly Then SetAttr NewCarp + "data.txt", 0
                End If
                tERR.Anotar "acip", ArchDaTa, NewCarp
                FSO.CopyFile ArchDaTa, NewCarp + "data.txt"
            End If
            TMPfiles = ObtenerArchMM(Ubic) 'deveuelve pathfull , solonombre
            tERR.Anotar "aciq", Ubic
            c = 1
            Do While c <= TotMM 'se supone que es el total de esta carpeta
                PathArch = txtInLista(TMPfiles(c), 0, "#")
                SoloArch = txtInLista(TMPfiles(c), 1, "#")
                'SI SOLO ARCH TIENE COMAS?
                SoloArch = Replace(SoloArch, ",", "")
                'soloarch es para el destino por lo que puedo modificarlo
                Select Case IDIOMA
                    Case "Español"
                        lblBAR2 = "Copiando " + PathArch
                    Case "English"
                        lblBAR2 = "Copiyng " + PathArch
                    Case "Francois"
                    Case "Italiano"
                End Select
                
                lblBAR2.Refresh
                tERR.Anotar "acir", ArchCopiados
                ArchCopiados = ArchCopiados + 1
                PBar2.Width = P2.Width / TotalACopiar * ArchCopiados
                PBar2.Refresh
                'si existe ver los atributos
                If FSO.FileExists(NewCarp + SoloArch) Then
                    tERR.Anotar "acis", NewCarp + SoloArch
                    AAA = GetAttr(NewCarp + SoloArch)
                    If AAA = vbHidden Or AAA = vbReadOnly Then SetAttr NewCarp + SoloArch, 0
                End If
                tERR.Anotar "acit", PathArch, NewCarp + SoloArch
                
                'sacar las comas de los nombres en el destino (por que el origen
                'puede ser CD o DVD de solo lectura)
                
                FSO.CopyFile PathArch, NewCarp + SoloArch, True
                
                Dim MbT As Long, MbF As Long, PL As Single
                txtInfoOrig = InfoDisco2(Left(lstOrigenes, 1), MbT, MbF, PL)
                lblP = CStr(PL)
                If PL < 10 Then
                    MsgBox "Queda menos del 10% de espacio en el disco!" + _
                        vbCrLf + "No se seguira copiando en este origen" + _
                        vbCrLf + "Use otra particion u otro disco con mas espacio"
                    Exit Sub
                End If
                
                c = c + 1
            Loop
            
            Select Case IDIOMA
                Case "Español"
                    lblBAR2 = "Sin Tareas"
                Case "English"
                    lblBAR2 = "Without Work"
                Case "Francois"
                Case "Italiano"
            End Select
            
            PBar2.Width = 0
            
        End If
    Next
    
    Select Case IDIOMA
        Case "Español"
            MsgBox "Los archivos se copiaron correctamente"
        Case "English"
            MsgBox "The files se copiaron"
        Case "Francois"
        Case "Italiano"
    End Select
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acia"
    Resume Next
End Sub

Private Sub Command5_Click()
    On Error GoTo MiErr
    
    Dim DRs As Drives, DS As Drive
    Set DRs = FSO.Drives
    CarpetaDesdeCargar = "NO"
    Dim CDsDisponibles() As String, ContCDs As Long
    ContCDs = -1
    For Each DS In DRs
        tERR.Anotar "aciu", DS.DriveType
        If DS.DriveType = 4 Then '4 es CDROM
            'ver cuantos hay disponibles!!!!
            ContCDs = ContCDs + 1
            ReDim Preserve CDsDisponibles(ContCDs)
            CDsDisponibles(ContCDs) = DS.DriveLetter
        End If
    Next
    'Que el tipo eliga la unidad que desea si es que hay mas de una
    If ContCDs = -1 Then
        Select Case IDIOMA
            Case "Español"
                MsgBox "No hay unidades de CD en su PC!"
            Case "English"
                MsgBox "There is not CD is PC"
            Case "Francois"
            Case "Italiano"
        End Select
        
        Exit Sub
    End If
    tERR.Anotar "aciv", ContCDs
    If ContCDs = 0 Then
        'no hay nada que legir
        Set DS = FSO.GetDrive(CDsDisponibles(0))
        GoTo ElegidoCD
    End If
    If ContCDs > 0 Then
        Set DS = FSO.GetDrive(CDsDisponibles(ContCDs))
        For A = 0 To ContCDs
            Set DS = FSO.GetDrive(CDsDisponibles(A))
            'muestra un mensaje completo si esta listo y si no solo la letra
            If DS.IsReady Then
                Select Case IDIOMA
                    Case "Español"
                        MSG = "Desea bucar en la unidad de CD:" + vbCrLf + _
                            DS.DriveLetter + "-" + DS.VolumeName + vbCrLf + _
                            "No = Unidad Siguiente"
                    Case "English"
                        MSG = "Want to search in:" + vbCrLf + _
                            DS.DriveLetter + "-" + DS.VolumeName + vbCrLf + _
                            "No = Next disc"
                    Case "Francois"
                    Case "Italiano"
                End Select
            Else
                Select Case IDIOMA
                    Case "Español"
                        MSG = "Desea bucar en la unidad de CD:" + vbCrLf + _
                            DS.DriveLetter + " (no esta listo)" + vbCrLf + _
                            "No = Unidad Siguiente"
                    Case "English"
                        MSG = "Want to search in:" + vbCrLf + _
                            DS.DriveLetter + " (no esta listo)" + vbCrLf + _
                            "No = Next disc"
                    Case "Francois"
                    Case "Italiano"
                End Select
                
            End If
            If MsgBox(MSG, vbYesNo) = vbYes Then GoTo ElegidoCD
        Next
        'si llego hasta aca y no eligio se caga por boludo
        Exit Sub
    End If
    
ElegidoCD:

    If DS.IsReady Then
        CarpetaDesdeCargar = DS.DriveLetter + ":\"
        tERR.Anotar "aciw", CarpetaDesdeCargar
    Else
        Select Case IDIOMA
            Case "Español"
                MsgBox "El disco " + DS.DriveLetter + " no esta " + _
                    "listo. Inserte un CD y reintente"
            Case "English"
                MsgBox "The disc " + DS.DriveLetter + " is not " + _
                    "ready. Insert a CD and try again"
            Case "Francois"
            Case "Italiano"
        End Select
        
        Exit Sub
    End If
    
    If CarpetaDesdeCargar = "NO" Then
        Select Case IDIOMA
            Case "Español"
                MsgBox "No se encontro unidad de CD"
            Case "English"
                MsgBox "Not Find CD drive"
            Case "Francois"
            Case "Italiano"
        End Select
        
        Exit Sub
    End If
    lblWait.Visible = True
    lblWait.Refresh
    'buscar carpetas de multimedia
    CarpsConMM = FindCarpsConMM(CarpetaDesdeCargar)
    lblWait = "Carpetas encontradas el la ubicacion elegida: " + CStr(UBound(CarpsConMM))
    lblWait.Refresh
    'ver cuales tienen multimedia
    tERR.Anotar "acix"
    BuscarCarpetasMM
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acib"
    Resume Next
    
End Sub

Public Function FindCarpsConMM(Carp As String) As String()
    On Error GoTo MiErr
    
    'devuelve una matriz con todas las carpetas que tengan multimedia
    Dim CarpetasEnQueBuscar() As String
    Dim Nivel2() As String
    Dim TodasLasCarpetas() As String
    Dim LastIni As Long, LastFin As Long
    CarpetasEnQueBuscar = GetFolders(Carp)
    LastIni = 1
    LastFin = UBound(CarpetasEnQueBuscar)
    tERR.Anotar "aciy", LastFin
    Dim AgregadosEnVuelta
    Dim ContTotal As Long
    Do
        AgregadosEnVuelta = 0
        If LastIni = 1 And LastFin = 0 Then
            'es una carpeta sin subcarpetas
            tERR.Anotar "aciz"
            Select Case IDIOMA
                Case "Español"
                    MsgBox "3PM no ha encontrado subcarpetas en la " + _
                        "ubicacion elegida. Pruebe buscar en un nivel " + _
                        "superior del arbol de directorios"
                Case "English"
                Case "Francois"
                Case "Italiano"
            End Select
            
            ReDim Preserve FindCarpsConMM(0)
            Exit Function
        End If
        For A = LastIni To LastFin
            ContTotal = ContTotal + 1
            ReDim Preserve TodasLasCarpetas(ContTotal)
            TodasLasCarpetas(ContTotal) = CarpetasEnQueBuscar(A)
            Nivel2 = GetFolders(CarpetasEnQueBuscar(A)) 'el cero no existe
            tERR.Anotar "acja", A, ContTotal, UBound(Nivel2)
            If UBound(Nivel2) > 0 Then
                Dim LastCBuscar
                LastCBuscar = UBound(CarpetasEnQueBuscar)
                For z = 1 To UBound(Nivel2)
                    ReDim Preserve CarpetasEnQueBuscar(LastCBuscar + z)
                    CarpetasEnQueBuscar(LastCBuscar + z) = Nivel2(z)
                    AgregadosEnVuelta = AgregadosEnVuelta + 1
                    tERR.Anotar "acjb", z, AgregadosEnVuelta
                Next
                GoTo NextMM
            Else
                tERR.Anotar "acjc", LastFin, UBound(CarpetasEnQueBuscar)
                If LastFin = UBound(CarpetasEnQueBuscar) Then
                    If A = LastFin Then
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
        tERR.Anotar "acjd", LastIni, LastFin
    Loop
    FindCarpsConMM = TodasLasCarpetas
    
    Exit Function
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acic"
    Resume Next
End Function

' Devuelve un array de wcadenas que incluye todos los subdirectorios
' contenidos en una ruta
Function GetFolders(ruta As String) As String()
    On Error GoTo MiErr
    
    Dim Resultado() As String
    Dim NombreDir As String, ContadorArch As Long
    Dim Ruta2 As String
    ReDim Resultado(0) As String
    ' genera el nombre de ruta + barra invertida
    Ruta2 = ruta
    If Right$(Ruta2, 1) <> "\" Then Ruta2 = Ruta2 & "\"
    NombreDir = Dir$(Ruta2 & "*.*", vbDirectory)
    tERR.Anotar "acjd", Ruta2, NombreDir
    Do While Len(NombreDir)
        If NombreDir = "." Or NombreDir = ".." Then
            ' excluir las entradas "." y ".."
        ElseIf (GetAttr(Ruta2 & NombreDir) And vbDirectory) = 0 Then
            ' este es un archivo normal
        Else
            ' es un directorio
            ContadorArch = ContadorArch + 1
            ReDim Preserve Resultado(ContadorArch) As String
            ' incluir la ruta si se pide
            NombreDir = Ruta2 & NombreDir
            Resultado(ContadorArch) = NombreDir
        End If
        NombreDir = Dir$
    Loop
    
    ' proporciona el array resultante
    ReDim Preserve Resultado(ContadorArch) As String
    GetFolders = Resultado
    tERR.Anotar "acje"
    Exit Function
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acid"
    Resume Next
    
End Function

Sub ShowDriveList()
    On Local Error Resume Next
    Dim fs, D, dc, s, n
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set dc = fs.Drives
    For Each D In dc
        s = s & D.DriveLetter & " - "
        Select Case D.DriveType
            Case 0: T = "Desconocido"
            Case 1: T = "Separable"
            Case 2: T = "Fijo"
            Case 3: T = "Red"
            Case 4: T = "CD-ROM"
            Case 5: T = "Disco RAM"
        End Select
        
        If D.DriveType = 3 Then
            n = D.ShareName
        Else
            n = D.VolumeName
        End If
        s = s & n & "Tipo: " & T & vbCrLf
        tERR.Anotar "acjf", s
    Next
    MsgBox s
End Sub

Private Sub Form_Activate()
    On Error GoTo MiErr
    tERR.Anotar "acjg"
    Select Case IDIOMA
        Case "Español"
            Label1(3) = "Especificar ubicacion de los nuevos discos"
            Command6.Caption = "CD Audio"
            Command5.Caption = "CD/DVD"
            Command1.Caption = "Explorar"
            Label1(7) = "Desde aqui podra trandsformar un CD de audio en ficheros mp3."
            Label1(0) = "3PM busca automaticamente en todas las carpetas del CD insertado."
            Label1(1) = "Explore usted por nuevos discos. Ususe para discos duros o unidades de red."
            Label2(1) = "Revise y controle la lista para asegurarse que el material encontrado es el deseado. Solo se agregaran aquellos discos que esten seleccionados. Quite todo el material que no sea necesario. Una vez terminado presione el boton AGREGAR"
            Command4.Caption = "Agregar estos discos a mi fonola"
            Command3.Caption = "SALIR"
            lblBAR.Caption = "Sin Tareas"
            lblBAR2.Caption = "Sin Tareas"
            lblWait = "Analizando disco.  Espere..."
            'lblInfoDisco = "Informacion del disco"
        Case "English"
            Label1(3) = "Especify ubication for the new music"
            Command6.Caption = "Audio CD"
            Command5.Caption = "CD/DVD"
            Command1.Caption = "Explore"
            Label1(7) = "Encode audio CD in mp3 files."
            Label1(0) = "3PM find automatically all folder for inserted CD"
            Label1(1) = "Find manually for new discs. Use it for hard disks or lan conbections"
            Label2(1) = "Revise y controle la lista para asegurarse que el material encontrado es el deseado. Solo se agregaran aquellos discos que esten seleccionados. Quite todo el material que no sea necesario. Una vez terminado presione el boton AGREGAR"
            Command4.Caption = "Agregar estos discos a mi fonola"
            Command3.Caption = "SALIR"
            lblBAR.Caption = "Sin Tareas"
            lblBAR2.Caption = "Sin Tareas"
            lblWait = "Analizando disco.  Espere..."
            'lblInfoDisco = "Informacion del disco"
        Case "Francois"
        
        Case "Italiano"
        
    End Select
    
Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acie"
    Resume Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case TeclaCerrarSistema
            YaCerrar3PM
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo MiErr
    tERR.Anotar "acjh", KeyCode, Shift
    If KeyCode = TeclaNewFicha Then
        LTE 1
        VarCreditos CSng(TemasPorCredito)
        
    End If
    
    If KeyCode = TeclaNewFicha2 Then
        LTE 2
        VarCreditos CSng(CreditosBilletes)
        
    End If
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acif"
    Resume Next
End Sub

Private Sub Form_Load()
    AjustarFRM Me, 12000
    
    Dim MtxTmpOrigenes() As String
    Dim Origenes As String
    Origenes = LeerArch1Linea(GPF("origs"))
    
    Dim PartOrigenes() As String
    PartOrigenes = Split(Origenes, "*")
    
    Dim AAA As Long: lstOrigenes.Clear
    For AAA = 0 To UBound(PartOrigenes)
        lstOrigenes.AddItem PartOrigenes(AAA)
        tERR.Anotar "acfc8", PartOrigenes(AAA)
    Next AAA
    lstOrigenes.ListIndex = 0
End Sub

Private Sub lstOrigenes_Click()
    Dim MbT As Long, MbF As Long, PL As Single
    txtInfoOrig = InfoDisco2(Left(lstOrigenes, 1), MbT, MbF, PL)
    lblP = CStr(PL)
End Sub

