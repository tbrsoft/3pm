VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmAddMusic 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar música a 3PM"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Orígenes disponibles"
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
      Height          =   2505
      Left            =   8280
      TabIndex        =   25
      Top             =   1320
      Width           =   3555
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
         Height          =   1245
         IntegralHeight  =   0   'False
         Left            =   60
         TabIndex        =   27
         Top             =   270
         Width           =   3375
      End
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
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1560
         Width           =   3375
      End
   End
   Begin tbrFaroButton.fBoton Command3 
      Height          =   615
      Left            =   8310
      TabIndex        =   24
      Top             =   7860
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command4 
      Height          =   525
      Left            =   8310
      TabIndex        =   23
      Top             =   5700
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   926
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar estos discos a mi fonola"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin VB.PictureBox PBar2 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   8325
      ScaleHeight     =   195
      ScaleWidth      =   15
      TabIndex        =   10
      Top             =   6285
      Width           =   15
   End
   Begin VB.ListBox lstCarConMM 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   2430
      Width           =   7695
   End
   Begin VB.PictureBox PBar 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   150
      ScaleHeight     =   165
      ScaleWidth      =   15
      TabIndex        =   4
      Top             =   7520
      Width           =   15
   End
   Begin VB.PictureBox P 
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   150
      ScaleHeight     =   135
      ScaleWidth      =   7485
      TabIndex        =   3
      Top             =   7500
      Visible         =   0   'False
      Width           =   7545
   End
   Begin VB.PictureBox P2 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8325
      ScaleHeight     =   135
      ScaleWidth      =   3435
      TabIndex        =   11
      Top             =   6285
      Visible         =   0   'False
      Width           =   3495
   End
   Begin tbrFaroButton.fBoton Command6 
      Height          =   435
      Left            =   30
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "CD Audio"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton command5 
      Height          =   435
      Left            =   30
      TabIndex        =   21
      Top             =   780
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "CD/DVD"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   435
      Left            =   30
      TabIndex        =   22
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Explorar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5452834
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Deje marcados solo aquellos discos que desee copiar."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   8
      Left            =   1260
      TabIndex        =   28
      Top             =   2220
      Width           =   6345
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   150
      TabIndex        =   19
      Top             =   7980
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
      ForeColor       =   &H00E0E0E0&
      Height          =   1230
      Index           =   0
      Left            =   8430
      TabIndex        =   18
      Top             =   90
      Width           =   3435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Index           =   2
      Left            =   7860
      TabIndex        =   17
      Top             =   120
      Width           =   525
   End
   Begin VB.Label lblWait 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Analizando disco. Espere..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   630
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   7035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde aquí podrá trandsformar un CD de audio en ficheros mp3."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   7
      Left            =   1305
      TabIndex        =   16
      Top             =   450
      Visible         =   0   'False
      Width           =   6525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Index           =   6
      Left            =   7800
      TabIndex        =   15
      Top             =   3840
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   5
      Left            =   90
      TabIndex        =   14
      Top             =   1920
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   4
      Left            =   60
      TabIndex        =   13
      Top             =   -30
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
      ForeColor       =   &H00E0E0E0&
      Height          =   1005
      Left            =   8325
      TabIndex        =   12
      Top             =   6525
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAddMusic.frx":00A4
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1635
      Index           =   1
      Left            =   8250
      TabIndex        =   9
      Top             =   3900
      Width           =   3615
   End
   Begin VB.Label lblBAR 
      BackStyle       =   0  'Transparent
      Caption         =   "Sin Tareas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   150
      TabIndex        =   7
      Top             =   7740
      Width           =   7665
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
      Left            =   960
      TabIndex        =   5
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   2
      Top             =   60
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
      Caption         =   "Explore usted por nuevos discos. Use para discos duros o unidades de red."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   1
      Left            =   1305
      TabIndex        =   1
      Top             =   1200
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
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Index           =   0
      Left            =   1305
      TabIndex        =   0
      Top             =   780
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
    X.DialogPrompt = TR.Trad("Elegir carpeta contenedora de los nuevos discos%99%")
    
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
    
    TotalArchMM = 0 'inicializar el contador de archivosmm (no carpetas)
    Dim TotF As Long
    TotF = UBound(CarpsConMM)
    PBar.Width = 0
    Dim TotCarpMM As Long
    Dim TMPfilesMM() As String
    lstCarConMM.Clear
    TotCarpMM = 0
    tERR.Anotar "acih", TotF
    For A = 1 To TotF
        lblBAR = TR.Trad("Buscando en%99%") + " " + CarpsConMM(A)
        PBar.Width = P.Width / TotF * A
        PBar.Refresh
        'OM- estoy navegando por todas las carpetas de algun lugar viendo si estan tienen multimedia
        'si el cliente no define que esta buscando quizas debería buscar automáticamente
        'tener en cuenta si se va a copiar todo lo que hay o solo lo que importa de la carpeta
        'es importante por el llenado del disco y que las carpetas no vengan con agregados que
        'no se cuentan en el tamaño del disco
        'mm91
        Dim Perf As Long
        If VentaExtras Then
            Perf = 1
            TMPfilesMM = ObtenerArchMM(CarpsConMM(A), , Perf)
        Else
            TMPfilesMM = ObtenerArchMM(CarpsConMM(A))
            Perf = 1
        End If
        'mm91
        If UBound(TMPfilesMM) > 0 Then
            TotalArchMM = TotalArchMM + UBound(TMPfilesMM)
            TR.SetVars "# " + CStr(UBound(TMPfilesMM)) + " "
            
            Select Case Perf
                Case 1
                    lstCarConMM.AddItem CarpsConMM(A) + TR.Trad("%01% canciones%98%La " + "variable 1 es un numero que se calcula%99%")
                Case 2
                    lstCarConMM.AddItem CarpsConMM(A) + TR.Trad("%01% ringtones%98%La " + "variable 1 es un numero que se calcula%99%")
                Case 3
                    lstCarConMM.AddItem CarpsConMM(A) + TR.Trad("%01% wallpapers%98%La " + "variable 1 es un numero que se calcula%99%")
                Case 4
                    lstCarConMM.AddItem CarpsConMM(A) + TR.Trad("%01% juegos java%98%La " + "variable 1 es un numero que se calcula%99%")
                Case 5
                    lstCarConMM.AddItem CarpsConMM(A) + TR.Trad("%01% imágenes ISO%98%La " + "variable 1 es un numero que se calcula%99%")
                Case 6
                    lstCarConMM.AddItem CarpsConMM(A) + TR.Trad("%01% videos 3GP%98%La " + "variable 1 es un numero que se calcula%99%")
                Case 7
                    lstCarConMM.AddItem CarpsConMM(A) + TR.Trad("%01% temas para movil%98%La " + "variable 1 es un numero que se calcula%99%")
            End Select
            
            lstCarConMM.Selected(TotCarpMM) = True
            TotCarpMM = TotCarpMM + 1
            Ltit = TR.Trad("Carpetas multimedia encontradas: %99%") + CStr(TotCarpMM)
            Ltit.Refresh
        End If
    Next
    lblWait.Visible = False
    lblWait.Refresh
    lblBAR = TR.Trad("Sin Tareas%99%")
    
    PBar.Width = 0
    If TotCarpMM = 0 Then
        MsgBox TR.Trad("No se han encontrado carpetas multimedia en " + _
            "la ubicacion elegida%99%")
    Else
        TR.SetVars TotCarpMM
        MsgBox TR.Trad("Se han encontrado %01% carpetas multimedia en " + _
            "la ubicacion elegida%99%")
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
            NewCarp = Replace(NewCarp, "#", "")
            tERR.Anotar "acil", A, NewCarp, TotMM
            If fso.FolderExists(NewCarp) = False Then fso.CreateFolder NewCarp
            
            'NO OLVIDARSE DE TAPA.JPG Y DATA.TXT
            Dim ArchTapa As String
            
            'debo agregar TODOS los jpg (que son ilustraciones de objetos o tapas no achicadas)
            'y todos los TXT que son mis descripciones
            'corregido en set 2008
            
'            ArchTapa = Ubic + "tapa.jpg"
'            If fso.FileExists(ArchTapa) Then
'                'si existe ver los atributos
'                If fso.FileExists(NewCarp + "tapa.jpg") Then
'                    AAA = GetAttr(NewCarp + "tapa.jpg")
'                    tERR.Anotar "acim", AAA
'                    If AAA = vbHidden Or AAA = vbReadOnly Then SetAttr NewCarp + "tapa.jpg", 0
'                End If
'                tERR.Anotar "acin", ArchTapa, NewCarp
'                fso.CopyFile ArchTapa, NewCarp + "tapa.jpg"
'            End If
            
'            Dim ArchDaTa As String
'            ArchDaTa = Ubic + "data.txt"
'            If fso.FileExists(ArchDaTa) Then
'                tERR.Anotar "acio"
'                'si existe ver los atributos
'                If fso.FileExists(NewCarp + "data.txt") Then
'                    AAA = GetAttr(NewCarp + "data.txt")
'                    If AAA = vbHidden Or AAA = vbReadOnly Then SetAttr NewCarp + "data.txt", 0
'                End If
'                tERR.Anotar "acip", ArchDaTa, NewCarp
'                fso.CopyFile ArchDaTa, NewCarp + "data.txt"
'            End If

            ArchTapa = Dir(Ubic + "*.jpg") 'dir devuelve solo el nombre de archivo+ext sin el path
            Do While ArchTapa <> ""
                If ArchTapa <> "." And ArchTapa <> ".." Then
                    'si existe ver los atributos para reemplazar sin error
                    If fso.FileExists(NewCarp + ArchTapa) Then
                        AAA = GetAttr(NewCarp + ArchTapa)
                        tERR.Anotar "acim", AAA
                        If AAA = vbHidden Or AAA = vbReadOnly Then SetAttr NewCarp + ArchTapa, 0
                    End If
                    tERR.Anotar "acin", ArchTapa, NewCarp
                    fso.CopyFile Ubic + ArchTapa, NewCarp + ArchTapa
                End If
                ArchTapa = Dir
            Loop
            
            ArchTapa = Dir(Ubic + "*.txt") 'dir devuelve solo el nombre de archivo+ext sin el path
            Do While ArchTapa <> ""
                If ArchTapa <> "." And ArchTapa <> ".." Then
                    'si existe ver los atributos para reemplazar sin error
                    If fso.FileExists(NewCarp + ArchTapa) Then
                        AAA = GetAttr(NewCarp + ArchTapa)
                        tERR.Anotar "acim", AAA
                        If AAA = vbHidden Or AAA = vbReadOnly Then SetAttr NewCarp + ArchTapa, 0
                    End If
                    tERR.Anotar "acin", ArchTapa, NewCarp
                    fso.CopyFile Ubic + ArchTapa, NewCarp + ArchTapa
                End If
                ArchTapa = Dir
            Loop
                        
            'OM- empezando a copiar todo lo de la lista elegida en asistente para agregar multimedia
            'mm91
            If VentaExtras Then
                TMPfiles = ObtenerArchMM(Ubic, , 1) 'deveuelve pathfull , solonombre
            Else
                TMPfiles = ObtenerArchMM(Ubic) 'deveuelve pathfull , solonombre
            End If
            
            tERR.Anotar "aciq", Ubic
            C = 1
            Do While C <= TotMM 'se supone que es el total de esta carpeta
                PathArch = txtInLista(TMPfiles(C), 0, "#")
                SoloArch = txtInLista(TMPfiles(C), 1, "#")
                'SI SOLO ARCH TIENE COMAS?
                SoloArch = Replace(SoloArch, ",", "")
                'soloarch es para el destino por lo que puedo modificarlo
                lblBAR2 = TR.Trad("Copiando %99%") + PathArch
                
                lblBAR2.Refresh
                tERR.Anotar "acir", ArchCopiados
                ArchCopiados = ArchCopiados + 1
                PBar2.Width = P2.Width / TotalACopiar * ArchCopiados
                PBar2.Refresh
                'si existe ver los atributos
                If fso.FileExists(NewCarp + SoloArch) Then
                    tERR.Anotar "acis", NewCarp + SoloArch
                    AAA = GetAttr(NewCarp + SoloArch)
                    If AAA = vbHidden Or AAA = vbReadOnly Then SetAttr NewCarp + SoloArch, 0
                End If
                tERR.Anotar "acit", PathArch, NewCarp + SoloArch
                
                'sacar las comas de los nombres en el destino (por que el origen
                'puede ser CD o DVD de solo lectura)
                
                fso.CopyFile PathArch, NewCarp + SoloArch, True
                
                Dim MbT As Long, MbF As Long, PL As Single
                txtInfoOrig = InfoDisco2(Left(lstOrigenes, 1), MbT, MbF, PL)
                lblP = CStr(PL)
                If PL < 10 Then
                    MsgBox TR.Trad("Queda menos del 10% de espacio en el disco!" + vbCrLf + _
                        "No se seguira copiando en este origen de discos. " + _
                        "Use otra particion u otro disco con mas espacio%99%")
                    Exit Sub
                End If
                
                C = C + 1
            Loop
            lblBAR2 = TR.Trad("Sin Tareas%99%")
            
            PBar2.Width = 0
            
        End If
    Next
    
    MsgBox TR.Trad("Los archivos se copiaron correctamente%99%") + vbCrLf + _
        TR.Trad("El contenido agregado estará disponible la próxima vez que inicie 3PM%99%")
    
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acia"
    Resume Next
End Sub

Private Sub Command5_Click()
    On Error GoTo MiErr
    
    Dim DRs As Drives, DS As Drive
    Set DRs = fso.Drives
    
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
        MsgBox TR.Trad("No hay unidades de CD en su PC!%99%")
        Exit Sub
    End If
    tERR.Anotar "aciv", ContCDs
    If ContCDs = 0 Then
        'no hay nada que legir
        Set DS = fso.GetDrive(CDsDisponibles(0))
        GoTo ElegidoCD
    End If
    If ContCDs > 0 Then
        Set DS = fso.GetDrive(CDsDisponibles(ContCDs))
        For A = 0 To ContCDs
            Set DS = fso.GetDrive(CDsDisponibles(A))
            'muestra un mensaje completo si esta listo y si no solo la letra
            If DS.IsReady Then
                TR.SetVars DS.DriveLetter, DS.VolumeName
                msg = TR.Trad("Desea bucar en la unidad de CD:" + vbCrLf + _
                     "%01%-%02%" + vbCrLf + _
                    "No = Unidad Siguiente%98%La variable 1 es una letra de " + _
                    "unidad como D o E y la variable 2 es la etiqueta de ese " + _
                    "volumen%99%")
            Else
                TR.SetVars DS.DriveLetter
                msg = TR.Trad("Desea bucar en la unidad de CD:" + vbCrLf + _
                    "%01% (unidad no lista)" + vbCrLf + _
                    "No = Unidad Siguiente%99%")
            End If
            If MsgBox(msg, vbYesNo) = vbYes Then GoTo ElegidoCD
        Next
        'si llego hasta aca y no eligio se caga por boludo
        Exit Sub
    End If
    
ElegidoCD:

    If DS.IsReady Then
        CarpetaDesdeCargar = DS.DriveLetter + ":\"
        tERR.Anotar "aciw", CarpetaDesdeCargar
    Else
        TR.SetVars DS.DriveLetter
        MsgBox TR.Trad("El disco %01% no esta listo. Inserte un CD y reintente%99%")
        
        Exit Sub
    End If
    
    If CarpetaDesdeCargar = "NO" Then
        MsgBox TR.Trad("No se encontro unidad de CD%99%")
        Exit Sub
    End If
    lblWait.Visible = True
    lblWait.Refresh
    'buscar carpetas de multimedia
    CarpsConMM = FindCarpsConMM(CarpetaDesdeCargar)
    lblWait = TR.Trad("Carpetas encontradas el la ubicación elegida: %99%") + _
        CStr(UBound(CarpsConMM))
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
            MsgBox TR.Trad("3PM no ha encontrado subcarpetas en la ubicación " + _
                "elegida. Pruebe buscar en un nivel superior del árbol " + _
                "de directorios%99%")
            
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
                For Z = 1 To UBound(Nivel2)
                    ReDim Preserve CarpetasEnQueBuscar(LastCBuscar + Z)
                    CarpetasEnQueBuscar(LastCBuscar + Z) = Nivel2(Z)
                    AgregadosEnVuelta = AgregadosEnVuelta + 1
                    tERR.Anotar "acjb", Z, AgregadosEnVuelta
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
        tERR.Anotar "acjd-2", LastIni, LastFin
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
    tERR.Anotar "acjd-1", Ruta2, NombreDir
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
    Dim fs, D, dc, S, n
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set dc = fs.Drives
    For Each D In dc
        S = S & D.DriveLetter & " - "
        Select Case D.DriveType
            Case 0: T = TR.Trad("Desconocido%99%")
            Case 1: T = TR.Trad("Separable%98%Se refiere a una unidad de " + _
                "disco que se puede extraer de la PC%99%")
            Case 2: T = TR.Trad("Fijo%99%")
            Case 3: T = TR.Trad("Red%99%")
            Case 4: T = TR.Trad("CD-ROM%99%")
            Case 5: T = TR.Trad("Disco RAM%99%")
        End Select
        
        If D.DriveType = 3 Then
            n = D.ShareName
        Else
            n = D.VolumeName
        End If
        S = S & n & TR.Trad("Tipo: %99%") & T & vbCrLf
        tERR.Anotar "acjf", S
    Next
    MsgBox S
End Sub

Private Sub fBoton3_Click()

End Sub

Private Sub fBoton4_Click()

End Sub

Private Sub Form_Activate()
    On Error GoTo MiErr
    tERR.Anotar "acjg"
    Label1(3) = TR.Trad("Especificar ubicación de los nuevos discos%99%")
    Command6.Caption = TR.Trad("CD Audio%99%")
    Command5.Caption = TR.Trad("CD/DVD%99%")
    Command1.Caption = TR.Trad("Explorar%99%")
    Label1(7) = TR.Trad("Desde aquí podrá trandsformar un CD de audio en " + _
        "ficheros mp3.%99%")
    Label1(0) = TR.Trad("3PM busca automaticamente en todas las carpetas " + _
        "del CD insertado.%99%")
    Label1(1) = TR.Trad("Explore usted por nuevos discos. Use para discos " + _
        "duros o unidades de red.%99%")
    Label2(1) = TR.Trad("Revise y controle la lista para asegurarse que " + _
        "el material encontrado es el deseado. Solo se agregaran aquellos " + _
        "discos que esten seleccionados. Quite todo el material que " + _
        "no sea necesario. Una vez terminado presione el boton AGREGAR%99%")
        
    Command4.Caption = TR.Trad("Agregar estos discos a mi fonola%99%")
    Command3.Caption = TR.Trad("SALIR%99%")
    lblBAR.Caption = TR.Trad("Sin Tareas%99%")
    lblBAR2.Caption = TR.Trad("Sin Tareas%99%")
    lblWait = TR.Trad("Analizando disco. Espere...%99%")
    'lblInfoDisco = "Informacion del disco"
Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acie"
    Resume Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case TeclaCerrarSistema
            Unload Me
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
    Traducir 'Agregado por el complemento traductor
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acif"
    Resume Next
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
    AjustarFRM Me, 12000, 9000
    
'    Dim MtxTmpOrigenes() As String
'    Dim Origenes As String
'    Origenes = LeerArch1Linea(GPF("origs"))
'
'    PartOrigenes = Split(Origenes, "*")
'       ya es publico se carga en index load
'
    Dim AAA As Long: lstOrigenes.Clear
    For AAA = 0 To UBound(PartOrigenes)
        lstOrigenes.AddItem PartOrigenes(AAA)
        tERR.Anotar "acfc8", PartOrigenes(AAA)
    Next AAA
    lstOrigenes.ListIndex = 0
    'tbrPintar fondo, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
    Frame1.Refresh
End Sub

Private Sub lstOrigenes_Click()
    Dim MbT As Long, MbF As Long, PL As Single
    txtInfoOrig = InfoDisco2(Left(lstOrigenes, 1), MbT, MbF, PL)
    lblP = CStr(PL)
End Sub

'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Frame1.Caption = TR.Trad("Origenes disponibles%99%")
    Command6.Caption = TR.Trad("CD Audio%99%")
    Command4.Caption = TR.Trad("Agregar estos discos a mi fonola%99%")
    Command3.Caption = TR.Trad("SALIR%99%")
    Command1.Caption = TR.Trad("Explorar%99%")
    Command5.Caption = TR.Trad("CD/DVD%99%")
    lblP.Caption = TR.Trad("% libre%99%")
    Label2(0).Caption = TR.Trad("Indique a que origen de discos se copiará la " + _
        "música. Revise el espacio disponible en el disco a usar. NO SIGA " + _
        "CARGANDO SI EL ESPACIO DISPONIBLE ES MENOR AL 10%.%99%")
    lblWait.Caption = TR.Trad("Analizando disco.  Espere...%99%")
    Label1(7).Caption = TR.Trad("Desde aqui podra trandsformar un CD de " + _
        "audio en ficheros mp3%99%")
    lblBAR2.Caption = TR.Trad("Sin Tareas%99%")
    Label2(1).Caption = TR.Trad("Revise y controle la lista para asegurarse " + _
        "que el material encontrado es el deseado. Solo se agregaran " + _
        "aquellos discos que esten seleccionados. Quite todo el material " + _
        "que no sea necesario. Una vez terminado presione el boton AGREGAR%99%")
    lblBAR.Caption = TR.Trad("Sin Tareas%99%")
    Ltit.Caption = TR.Trad("Carpetas encontradas con multimedia: 0%99%")
    Label1(3).Caption = TR.Trad("Especificar ubicación de los nuevos discos%99%")
    Label1(1).Caption = TR.Trad("Explore usted por nuevos discos. Usese para " + _
        "discos duros o unidades de red.%99%")
    Label1(0).Caption = TR.Trad("3PM busca automaticamente en todas las " + _
        "carpetas del CD insertado.%99%")
End Sub

