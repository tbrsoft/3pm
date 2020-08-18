VERSION 5.00
Begin VB.UserControl tbrPassImg 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   ControlContainer=   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   2475
   Begin VB.Timer RELOJ 
      Left            =   630
      Top             =   990
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "xxx"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image IMG 
      Height          =   2250
      Left            =   90
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2250
   End
End
Attribute VB_Name = "tbrPassImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mIntervalBetwenIMGs As Long 'segundos entre foto y foto.
'Minimo 10 segundos para que no joda al programa

Private mArchPictures() As String 'lista de las fotos a usar. Si es solo una que ni active el timer
'el cero no lo usa para no tener drama. Ya se cargan ordenados del clsPUB
Private mTotalImagenes As Long
Private mUltimaReproducida As Long 'ultimo n° de imagen para saber cual sigue
Private mActivarPUBS As Boolean 'saber si esta activo!
Public Event ChangeImg()

Public Property Let ActivarPUBS(Activar As Boolean)
    mActivarPUBS = Activar
End Property

Public Property Get ActivarPUBS() As Boolean
    ActivarPUBS = mActivarPUBS
End Property

Public Property Let IntervalBetwenIMGs(Secs As Long)
    mIntervalBetwenIMGs = Secs
End Property

Public Sub Picture(ArchImagen As String)
    'permite poner una imagen fija sin leer pubs
    'esto para las SuperLicencias que esten activas y antes habia un IMAGE1.PICTURE= y ahora
    'debe haber un tbrPASS.PICTURE=
    IMG.Picture = LoadPicture(ArchImagen)
End Sub

Public Sub Refresh()
    IMG.Refresh
End Sub

Public Property Get IntervalBetwenIMGs() As Long
    IntervalBetwenIMGs = mIntervalBetwenIMGs
End Property

Private Sub UserControl_Initialize()
    tERR.Anotar "PASI001"
    mTotalImagenes = 0
    IMG.Stretch = True
    mUltimaReproducida = 0 'de entrada va al 1
    ReDim Preserve mArchPictures(0)
End Sub

Private Sub UserControl_Resize()
    IMG.Width = UserControl.Width
    IMG.Height = UserControl.Height
    IMG.Top = 0
    IMG.Left = 0
    Label1.Width = UserControl.Width
    Label1.Left = 0
End Sub

Public Sub AddArchivoIMG(Arch As String)
    mTotalImagenes = UBound(mArchPictures) + 1
    ReDim Preserve mArchPictures(mTotalImagenes)
    mArchPictures(mTotalImagenes) = Arch
End Sub

Public Sub ClearList()
    Erase mArchPictures
    ReDim Preserve mArchPictures(0)
End Sub

Public Sub IniciarPASS()
    'empezar a reproducir
    'si no hay mas de una ni moverse.
    'si hay cero ni hablar
    
    'si no activo se caga
    If mActivarPUBS Then
        If mTotalImagenes >= 1 Then
            RELOJ.Interval = mIntervalBetwenIMGs * 1000
        End If
    End If
End Sub

Public Sub Detener()
    RELOJ.Interval = 0
End Sub

Private Sub Reloj_Timer()
    mUltimaReproducida = mUltimaReproducida + 1
    'si me paso se va al primero ya
    If mUltimaReproducida > mTotalImagenes Then mUltimaReproducida = 1
    '...
    '...
    'aca debe ir algun efecto. Ponete las pilas ANDRES
    '...
    '...
    IMG.Picture = LoadPicture(mArchPictures(mUltimaReproducida))
    Label1.Caption = mArchPictures(mUltimaReproducida)
    RaiseEvent ChangeImg
End Sub


