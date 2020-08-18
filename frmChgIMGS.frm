VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmChgIMGS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambiar componentes gráficos"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Imagen de inicio de 3PM"
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
      Height          =   1815
      Left            =   270
      TabIndex        =   4
      Top             =   2280
      Width           =   4725
      Begin tbrFaroButton.fBoton Command23 
         Height          =   465
         Left            =   270
         TabIndex        =   5
         Top             =   510
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Elegir imagen"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton XxBoton5 
         Height          =   465
         Left            =   270
         TabIndex        =   6
         Top             =   990
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Quitar"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.Image IMG2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1440
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   300
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Imagen de fondo de las portadas"
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
      Height          =   1815
      Left            =   270
      TabIndex        =   1
      Top             =   300
      Width           =   4725
      Begin tbrFaroButton.fBoton Command22 
         Height          =   465
         Left            =   270
         TabIndex        =   2
         Top             =   510
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Elegir imagen"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton XxBoton4 
         Height          =   465
         Left            =   270
         TabIndex        =   3
         Top             =   990
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Quitar"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.Image IMG1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1440
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   300
         Width           =   1995
      End
   End
   Begin tbrFaroButton.fBoton XxBoton1 
      Height          =   345
      Left            =   4290
      TabIndex        =   0
      Top             =   4260
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Tapa predetermianda de discos"
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
      Height          =   1815
      Left            =   5220
      TabIndex        =   7
      Top             =   300
      Width           =   4725
      Begin tbrFaroButton.fBoton Command24 
         Height          =   465
         Left            =   270
         TabIndex        =   8
         Top             =   510
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Elegir imagen"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton XxBoton6 
         Height          =   465
         Left            =   270
         TabIndex        =   9
         Top             =   990
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Quitar"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.Image IMG3 
         BorderStyle     =   1  'Fixed Single
         Height          =   1440
         Left            =   2610
         Stretch         =   -1  'True
         Top             =   270
         Width           =   1995
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Tapa predetermianda de ranking"
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
      Height          =   1815
      Left            =   5190
      TabIndex        =   10
      Top             =   2280
      Width           =   4725
      Begin tbrFaroButton.fBoton Command25 
         Height          =   465
         Left            =   270
         TabIndex        =   11
         Top             =   510
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Elegir imagen"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton XxBoton7 
         Height          =   465
         Left            =   270
         TabIndex        =   12
         Top             =   990
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Quitar"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.Image IMG4 
         BorderStyle     =   1  'Fixed Single
         Height          =   1440
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   300
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmChgIMGS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function GetArch(sTitle As String, sFilter As String) As String
    
    Dim CM As New CommonDialog
    CM.DialogPrompt = sTitle
    CM.DialogTitle = sTitle
    CM.Filter = sFilter
    
    CM.ShowOpen
    Dim F As String
    F = CM.FileName
    
    Set CM = Nothing
    
    GetArch = F
    
End Function

Private Sub Form_Load()
    Pintar_fBoton Me
    Traducir 'Agregado por el complemento traductor
    XxBoton4.Enabled = (K.LICENCIA("3pm") = HSuperLicencia)
    Command22.Enabled = XxBoton4.Enabled
    
    XxBoton5.Enabled = XxBoton4.Enabled
    Command23.Enabled = XxBoton4.Enabled
    
    XxBoton6.Enabled = XxBoton4.Enabled
    Command24.Enabled = XxBoton4.Enabled
    
    XxBoton7.Enabled = XxBoton4.Enabled
    Command25.Enabled = XxBoton4.Enabled

    'mostrar las imagenes usadas actualemente
    If fso.FileExists(GPF("iischu")) Then
        IMG1.Picture = LoadPicture(GPF("iischu"))
    Else
        IMG1.Picture = LoadPicture(ExtraData.GetImagePath("FondoDeLasTapas"))
    End If
    
    If fso.FileExists(GPF("iisl67")) Then
        IMG2.Picture = LoadPicture(GPF("iisl67"))
    Else
        IMG2.Picture = LoadPicture(ExtraData.GetImagePath("iniciasys"))
    End If
    
    If fso.FileExists(GPF("tddp322")) Then
        IMG3.Picture = LoadPicture(GPF("tddp322"))
    Else
        IMG3.Picture = LoadPicture(ExtraData.GetImagePath("tapapredeterminada"))
    End If
    
    If fso.FileExists(GPF("tddp323")) Then
        IMG4.Picture = LoadPicture(GPF("tddp323"))
    Else
        IMG4.Picture = LoadPicture(ExtraData.GetImagePath("taparanking"))
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmIndex.Timer3.Enabled = True
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub

Private Sub XxBoton1_Click()
    Unload Me
End Sub

Private Sub Command22_Click()
    Dim F As String
    F = GetArch(TR.Trad("Elegir imagen%99%"), TR.Trad("Imagenes%99%") + " (*.jpg, *.bmp, *.gif)|*.jpg; *.gif;*.bmp")
    If F = "" Then Exit Sub
    fso.CopyFile F, GPF("iischu"), True
    IMG1.Picture = LoadPicture(GPF("iischu"))
End Sub

Private Sub XxBoton4_Click()
    fso.DeleteFile GPF("iischu"), True
    'vuelve a la del skin original que si o si existe
    IMG1.Picture = LoadPicture(ExtraData.GetImagePath("FondoDeLasTapas"))
End Sub

Private Sub Command23_Click()
    Dim F As String
    F = GetArch(TR.Trad("Elegir imagen%99%"), TR.Trad("Imagenes%99%") + " (*.jpg, *.bmp, *.gif)|*.jpg; *.gif;*.bmp")
    If F = "" Then Exit Sub
    'lo pongo en la ubicacion
    fso.CopyFile F, GPF("iisl67"), True
    IMG2.Picture = LoadPicture(GPF("iisl67"))
End Sub

Private Sub XxBoton5_Click()
    fso.DeleteFile GPF("iisl67"), True
    'vuelve a la del skin original que si o si existe
    IMG2.Picture = LoadPicture(ExtraData.GetImagePath("iniciasys"))
End Sub

Private Sub Command24_Click()
    Dim F As String
    F = GetArch(TR.Trad("Elegir imagen%99%"), TR.Trad("Imagenes%99%") + " (*.jpg, *.bmp, *.gif)|*.jpg; *.gif;*.bmp")
    If F = "" Then Exit Sub
    fso.CopyFile F, GPF("tddp322"), True
    IMG3.Picture = LoadPicture(GPF("tddp322"))
End Sub

Private Sub XxBoton6_Click()
    fso.DeleteFile GPF("tddp322"), True
    'vuelve a la del skin original que si o si existe
    IMG3.Picture = LoadPicture(ExtraData.GetImagePath("tapapredeterminada"))
End Sub

Private Sub Command25_Click()
    Dim F As String
    F = GetArch(TR.Trad("Elegir imagen%99%"), TR.Trad("Imagenes%99%") + " (*.jpg, *.bmp, *.gif)|*.jpg; *.gif;*.bmp")
    If F = "" Then Exit Sub
    'lo pongo en la ubicacion
    fso.CopyFile F, GPF("tddp323"), True
    IMG4.Picture = LoadPicture(GPF("tddp323"))
End Sub

Private Sub XxBoton7_Click()
    fso.DeleteFile GPF("tddp323"), True
    'vuelve a la del skin original que si o si existe
    IMG4.Picture = LoadPicture(ExtraData.GetImagePath("taparanking"))
End Sub


'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Frame2.Caption = TR.Trad("Imagen de inicio de 3PM%99%")
    Command23.Caption = TR.Trad("Elegir imagen%99%")
    XxBoton5.Caption = TR.Trad("Quitar%99%")
    Frame1.Caption = TR.Trad("Imagen de fondo de las portadas%99%")
    Command22.Caption = TR.Trad("Elegir imagen%99%")
    XxBoton4.Caption = TR.Trad("Quitar%99%")
    XxBoton1.Caption = TR.Trad("SALIR%99%")
    Frame3.Caption = TR.Trad("Tapa predetermianda de discos%99%")
    Command24.Caption = TR.Trad("Elegir imagen%99%")
    XxBoton6.Caption = TR.Trad("Quitar%99%")
    Frame4.Caption = TR.Trad("Tapa predetermianda de ranking%99%")
    Command25.Caption = TR.Trad("Elegir imagen%99%")
    XxBoton7.Caption = TR.Trad("Quitar%99%")
End Sub
