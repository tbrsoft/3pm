VERSION 5.00
Begin VB.Form frmConfigVIS 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Graficos de 3PM"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "cambiar detalles"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5460
      TabIndex        =   9
      Top             =   840
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   0
      ScaleHeight     =   7335
      ScaleWidth      =   11775
      TabIndex        =   6
      Top             =   1320
      Width           =   11775
      Begin VB.PictureBox imgMarco 
         BackColor       =   &H00000000&
         Height          =   5805
         Left            =   930
         ScaleHeight     =   5745
         ScaleWidth      =   9405
         TabIndex        =   7
         Top             =   360
         Width           =   9465
         Begin VB.PictureBox imgFONDO 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            Height          =   5505
            Left            =   90
            ScaleHeight     =   5445
            ScaleWidth      =   9135
            TabIndex        =   8
            Top             =   150
            Width           =   9195
            Begin VB.Image imgTapaSel 
               BorderStyle     =   1  'Fixed Single
               Height          =   2505
               Left            =   120
               Stretch         =   -1  'True
               Top             =   150
               Width           =   2895
            End
            Begin VB.Image imgTapaSel2 
               BorderStyle     =   1  'Fixed Single
               Height          =   2505
               Left            =   3150
               Stretch         =   -1  'True
               Top             =   150
               Width           =   2895
            End
            Begin VB.Image ii3 
               BorderStyle     =   1  'Fixed Single
               Height          =   2505
               Left            =   6120
               Stretch         =   -1  'True
               Top             =   150
               Width           =   2895
            End
            Begin VB.Image ii4 
               BorderStyle     =   1  'Fixed Single
               Height          =   2505
               Left            =   150
               Stretch         =   -1  'True
               Top             =   2790
               Width           =   2895
            End
            Begin VB.Image ii5 
               BorderStyle     =   1  'Fixed Single
               Height          =   2505
               Left            =   3150
               Stretch         =   -1  'True
               Top             =   2790
               Width           =   2895
            End
            Begin VB.Image ii6 
               BorderStyle     =   1  'Fixed Single
               Height          =   2505
               Left            =   6150
               Stretch         =   -1  'True
               Top             =   2790
               Width           =   2895
            End
            Begin VB.Image imgDISC 
               BorderStyle     =   1  'Fixed Single
               Height          =   1965
               Index           =   0
               Left            =   210
               Stretch         =   -1  'True
               Top             =   270
               Width           =   2685
            End
            Begin VB.Image imgDISC 
               BorderStyle     =   1  'Fixed Single
               Height          =   1965
               Index           =   1
               Left            =   3240
               Stretch         =   -1  'True
               Top             =   270
               Width           =   2685
            End
            Begin VB.Image imgDISC 
               BorderStyle     =   1  'Fixed Single
               Height          =   1965
               Index           =   2
               Left            =   6210
               Stretch         =   -1  'True
               Top             =   270
               Width           =   2685
            End
            Begin VB.Image imgDISC 
               BorderStyle     =   1  'Fixed Single
               Height          =   1965
               Index           =   3
               Left            =   270
               Stretch         =   -1  'True
               Top             =   2910
               Width           =   2685
            End
            Begin VB.Image imgDISC 
               BorderStyle     =   1  'Fixed Single
               Height          =   1965
               Index           =   4
               Left            =   3240
               Stretch         =   -1  'True
               Top             =   2940
               Width           =   2685
            End
            Begin VB.Image imgDISC 
               BorderStyle     =   1  'Fixed Single
               Height          =   1965
               Index           =   5
               Left            =   6240
               Stretch         =   -1  'True
               Top             =   2970
               Width           =   2685
            End
         End
      End
      Begin VB.Image imgVUMSel 
         BorderStyle     =   1  'Fixed Single
         Height          =   5835
         Left            =   60
         Stretch         =   -1  'True
         Top             =   360
         Width           =   840
      End
      Begin VB.Image imgTouchSel 
         BorderStyle     =   1  'Fixed Single
         Height          =   885
         Left            =   60
         Stretch         =   -1  'True
         Top             =   6240
         Width           =   1695
      End
      Begin VB.Image imgVumSel2 
         BorderStyle     =   1  'Fixed Single
         Height          =   5835
         Left            =   10470
         Stretch         =   -1  'True
         Top             =   330
         Width           =   840
      End
      Begin VB.Image imgTouchSel2 
         BorderStyle     =   1  'Fixed Single
         Height          =   885
         Left            =   9690
         Stretch         =   -1  'True
         Top             =   6210
         Width           =   1635
      End
   End
   Begin VB.Timer Timer1 
      Left            =   7980
      Top             =   120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Crear nuevo skin"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5460
      TabIndex        =   5
      Top             =   300
      Width           =   1935
   End
   Begin VB.ComboBox cmbSK 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   300
      Width           =   5295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "grabar y salir"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4170
      TabIndex        =   2
      Top             =   780
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "salir sin grabar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2850
      TabIndex        =   1
      Top             =   780
      Width           =   1200
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Puede elegir manualmente cada grafica o cargar un skin de esta lista"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   90
      TabIndex        =   3
      Top             =   30
      Width           =   6495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Los cambios tendrán efectos en el proximo inicio de 3PM"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   810
      Width           =   2415
   End
End
Attribute VB_Name = "frmConfigVIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DirVum As Long

Private Sub cmbSK_Click()
    Timer1.Interval = 0
    
    'cargar el skin!
    Dim F As Long
    F = ExtraData.AbrirSKIN(AP + "skin\" + cmbSK + ".skin")
    If F = 1 Then 'alguien le cambio el nombre al original!
        MsgBox "Este skin tenia otro nombre y ha sido monidicado. Devuelva el archivo SKIN a su nombre" + _
            " original para poder utilizarlo"
        Exit Sub
    End If
    'mostrar cada una de las imagenes donde corresponde!
    imF = ExtraData.GetImagePath("vumetroapagado")
    imgVUMSel.Picture = LoadPicture(imF)
    imgVumSel2.Picture = LoadPicture(imF)
    
    imF = ExtraData.GetImagePath("MarcoFondodelosdiscos")
    imgMarco.PaintPicture LoadPicture(imF), 0, 0, imgMarco.Width, imgMarco.Height
    
    imF = ExtraData.GetImagePath("FondoDeLasTapas")
    imgFONDO.PaintPicture LoadPicture(imF), 0, 0, imgFONDO.Width, imgFONDO.Height
    
    imF = ExtraData.GetImagePath("marcodiscocomun")
    imgTapaSel.Picture = LoadPicture(imF)
    ii3.Picture = imgTapaSel.Picture
    ii4.Picture = imgTapaSel.Picture
    ii5.Picture = imgTapaSel.Picture
    ii6.Picture = imgTapaSel.Picture
    
    imF = ExtraData.GetImagePath("marcodiscoelegido")
    imgTapaSel2.Picture = LoadPicture(imF)
    
    imF = ExtraData.GetImagePath("touchizqnormal")
    'imF = ExtraData.GetImagePath("touchiapretado")
    imgTouchSel.Picture = LoadPicture(imF)
    
    imF = ExtraData.GetImagePath("touchderechanormal")
    'imF = ExtraData.GetImagePath("touchderechaapretado")
    imgTouchSel2.Picture = LoadPicture(imF)
    
    'ver si es superlicencia y usa otra tapa predeterminada
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(GPF("tddp323")) Then
            imF = GPF("tddp323")
        Else
            imF = ExtraData.GetImagePath("taparanking")
        End If
    Else
        imF = ExtraData.GetImagePath("taparanking")
    End If
    
    imgDISC(0).Picture = LoadPicture(imF)
    
    'ver si es superlicencia y usa otra tapa predeterminada
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(GPF("tddp322")) Then
            imF = GPF("tddp322")
        Else
            imF = ExtraData.GetImagePath("tapapredeterminada")
        End If
    Else
        imF = ExtraData.GetImagePath("tapapredeterminada")
    End If
    
    imgDISC(1).Picture = LoadPicture(imF)
    imgDISC(2).Picture = LoadPicture(imF)
    imgDISC(3).Picture = LoadPicture(imF)
    imgDISC(4).Picture = LoadPicture(imF)
    imgDISC(5).Picture = LoadPicture(imF)
    
    Timer1.Interval = 200
    
End Sub

Private Sub Command1_Click()
    'dejo cargado el skin que quedo grabado!!!!
    ExtraData.AbrirSKIN mySKIN
    Unload Me
End Sub

Private Sub Command2_Click()
    ChangeConfig "mySKIN", AP + "skin\" + cmbSK + ".skin"
    Unload Me
End Sub

Private Sub Command3_Click()
    frmCrearSKIN.Show 1
End Sub

Private Sub Command4_Click()
    frmChgIMGS.Show 1
End Sub

Private Sub Form_Load()
    
    'mostrar la lista de skins disponibles
    Dim lSK As String, SK() As String
    ReDim SK(0)
    Dim CC As Long
    CC = 0
    
    lSK = Dir(AP + "skIn\*.skin")
    Do While lSK <> ""
        ReDim Preserve SK(UBound(SK) + 1)
        SK(UBound(SK)) = lSK
        cmbSK.AddItem Mid(lSK, 1, Len(lSK) - 5)
        
        'seleccionar el que esta elegido
        If LCase(FSO.GetBaseName(mySKIN)) = LCase(cmbSK.List(CC)) Then
            cmbSK.ListIndex = CC
        End If
        
        lSK = Dir
        CC = CC + 1
    Loop
    
    DirVum = 10
    
    imgVUMSel.BorderStyle = 0
    imgVumSel2.BorderStyle = 0
    imgMarco.BorderStyle = 0
    imgFONDO.BorderStyle = 0
    imgTouchSel.BorderStyle = 0
    imgTouchSel2.BorderStyle = 0
    imgTapaSel.BorderStyle = 0
    imgTapaSel2.BorderStyle = 0
    ii3.BorderStyle = 0
    ii4.BorderStyle = 0
    ii5.BorderStyle = 0
    ii6.BorderStyle = 0
    imgDISC(0).BorderStyle = 0
    imgDISC(1).BorderStyle = 0
    imgDISC(2).BorderStyle = 0
    imgDISC(3).BorderStyle = 0
    imgDISC(4).BorderStyle = 0
    imgDISC(5).BorderStyle = 0
    
'    Dim F As String, NN As Long, N2 As String
'    Dim indIMG As Long 'cantidad de imagenes encontradas
'    F = AP + "sf\"
'    indIMG = 0
'    Dim ArFinal As String
'    For NN = 1 To 40
'        If NN < 10 Then N2 = "0" + CStr(NN)
'        If NN >= 10 Then N2 = CStr(NN)
'        ArFinal = F + "f74_" + N2 + ".dlw"
'        If FSO.FileExists(ArFinal) Then
'            If indIMG > 0 Then
'                Load imgTAPAS(indIMG)
'                imgTAPAS(indIMG).Left = imgTAPAS(indIMG - 1).Left + imgTAPAS(indIMG - 1).Width + 60
'                imgTAPAS(indIMG).Top = imgTAPAS(indIMG - 1).Top
'                picTapas.Width = imgTAPAS(indIMG).Left + imgTAPAS(indIMG).Width + 60
'                'la barra
'                hsTAPA.MAX = Abs(picTapas.Width - picBASETapa.Width) / 10 'es integer!!!
'                hsTAPA.LargeChange = hsTAPA.MAX / 10
'
'                'LO MISMO PARA EL ELEGIDO!
'                Load imgTAPAS2(indIMG)
'                imgTAPAS2(indIMG).Left = imgTAPAS2(indIMG - 1).Left + imgTAPAS2(indIMG - 1).Width + 60
'                imgTAPAS2(indIMG).Top = imgTAPAS2(indIMG - 1).Top
'                picTapas2.Width = imgTAPAS2(indIMG).Left + imgTAPAS2(indIMG).Width + 60
'                'la barra
'                hsTAPA2.MAX = Abs(picTapas2.Width - picBASETapa2.Width) / 10 'es integer!!!
'                hsTAPA2.LargeChange = hsTAPA2.MAX / 10
'            End If
'
'            imgTAPAS(indIMG).Tag = ArFinal
'            imgTAPAS(indIMG).Picture = LoadPicture(ArFinal)
'            imgTAPAS(indIMG).Visible = True
'
'            imgTAPAS2(indIMG).Tag = ArFinal
'            imgTAPAS2(indIMG).Picture = LoadPicture(ArFinal)
'            imgTAPAS2(indIMG).Visible = True
'
'            indIMG = indIMG + 1
'        End If
'    Next NN
'
'    hsTAPA.SmallChange = imgTAPAS(0).Width / 20
'    hsTAPA2.SmallChange = imgTAPAS(0).Width / 20
'
'    indIMG = 0
'    'touch es el 70 71 72 73
'    For NN = 1 To 40
'        If NN < 10 Then N2 = "0" + CStr(NN)
'        If NN >= 10 Then N2 = CStr(NN)
'        ArFinal = F + "f70_" + N2 + ".dlw"
'        If FSO.FileExists(ArFinal) Then
'            If indIMG > 0 Then
'                Load imgTouchs(indIMG)
'                imgTouchs(indIMG).Left = imgTouchs(indIMG - 1).Left + imgTouchs(indIMG - 1).Width + 60
'                imgTouchs(indIMG).Top = imgTouchs(indIMG - 1).Top
'                picTouchs.Width = imgTouchs(indIMG).Left + imgTouchs(indIMG).Width + 60
'                'la barra
'                hsTOUCHS.MAX = Abs(picTouchs.Width - picBaseTouch.Width) / 10  'es integer!!!
'                hsTOUCHS.LargeChange = hsTOUCHS.MAX / 10
'            End If
'
'            imgTouchs(indIMG).Tag = ArFinal
'            imgTouchs(indIMG).Picture = LoadPicture(ArFinal)
'            imgTouchs(indIMG).Visible = True
'
'            indIMG = indIMG + 1
'        End If
'    Next NN
'
'    hsTOUCHS.SmallChange = imgTouchs(0).Width / 20
'
'    'los vumetros
'    indIMG = 0
'    'touch es el 77 78
'    Dim ArFinal2 As String
'    For NN = 1 To 40
'        If NN < 10 Then N2 = "0" + CStr(NN)
'        If NN >= 10 Then N2 = CStr(NN)
'        ArFinal = F + "f77_" + N2 + ".dlw"
'        ArFinal2 = F + "f78_" + N2 + ".dlw"
'        If FSO.FileExists(ArFinal) Then
'            If indIMG > 0 Then
'                Load imgVUMs(indIMG)
'                imgVUMs(indIMG).Left = imgVUMs(indIMG - 1).Left + imgVUMs(indIMG - 1).Width + 60
'                imgVUMs(indIMG).Top = imgVUMs(indIMG - 1).Top
'
'                Load imgVUMs2(indIMG)
'                imgVUMs2(indIMG).Left = imgVUMs(indIMG).Left
'                imgVUMs2(indIMG).Top = imgVUMs(indIMG).Top
'
'                picVUMs.Width = imgVUMs(indIMG).Left + imgVUMs(indIMG).Width + 60
'                'la barra
'                hsVUMs.MAX = Abs(picVUMs.Width - picBaseVUM.Width) / 10   'es integer!!!
'                hsVUMs.LargeChange = hsVUMs.MAX / 10
'            End If
'
'            imgVUMs(indIMG).Tag = ArFinal
'            imgVUMs(indIMG).Picture = LoadPicture(ArFinal)
'            imgVUMs(indIMG).Visible = True
'
'            imgVUMs2(indIMG).Tag = ArFinal2
'            imgVUMs2(indIMG).Picture = LoadPicture(ArFinal2)
'            imgVUMs2(indIMG).Visible = True
'
'            imgVUMs(indIMG).ZOrder 'aca va el apagado por arriba tapando alguna parte
'
'            indIMG = indIMG + 1
'        End If
'    Next NN
'
'    hsTOUCHS.SmallChange = imgTouchs(0).Width / 20
'
'    Timer1.Interval = 200
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Timer1.Interval = 0
End Sub

'Private Sub hsTAPA_Change()
'    picTapas.Left = -CLng(hsTAPA.Value) * 10 'es integer!!!
'End Sub
'
'Private Sub hsTOUCHS_Change()
'    picTouchs.Left = -CLng(hsTOUCHS.Value) * 10 'es integer!!!
'End Sub
'
'Private Sub hsVUMs_Change()
'    picVUMs.Left = -CLng(hsVUMs.Value) * 10 'es integer!!!
'End Sub

'Private Sub imgTAPAS_Click(Index As Integer)
'    imgTapaSel.Picture = imgTAPAS(Index).Picture
'    ii3.Picture = imgTapaSel.Picture
'    ii4.Picture = imgTapaSel.Picture
'    ii5.Picture = imgTapaSel.Picture
'    ii6.Picture = imgTapaSel.Picture
'
'    imgTapaSel.Tag = imgTAPAS(Index).Tag
'End Sub
'
'Private Sub hsTAPA2_Change()
'    picTapas2.Left = -CLng(hsTAPA2.Value) * 10 'es integer!!!
'End Sub
'
'Private Sub imgTAPAS2_Click(Index As Integer)
'    imgTapaSel2.Picture = imgTAPAS2(Index).Picture
'    imgTapaSel2.Tag = imgTAPAS2(Index).Tag
'End Sub
'
'Private Sub imgTouchs_Click(Index As Integer)
'    imgTouchSel.Picture = imgTouchs(Index).Picture
'    imgTouchSel.Tag = imgTouchs(Index).Tag
'    imgTouchSel2.Picture = LoadPicture(AP + "sf\f72" + Right(imgTouchSel.Tag, 7))
'End Sub
'
'Private Sub imgVUMs_Click(Index As Integer)
'    imgVUMSel.Picture = imgVUMs2(Index).Picture
'    imgVumSel2.Picture = imgVUMs2(Index).Picture
'    imgVUMSel.Tag = imgVUMs2(Index).Tag
'End Sub
'
'Private Sub imgVUMs2_Click(Index As Integer)
'    imgVUMSel.Picture = imgVUMs2(Index).Picture
'    imgVumSel2.Picture = imgVUMs2(Index).Picture
'    imgVUMSel.Tag = imgVUMs2(Index).Tag
'End Sub
'
'Private Sub Timer1_Timer()
'    DirVum = DirVum + 5
'    If DirVum > 100 Then DirVum = 5
'    Dim M As Long
'    For M = 0 To imgVUMs.UBound
'        imgVUMs(M).Height = imgVUMs2(M).Height * DirVum / 100
'    Next M
'End Sub

Private Sub Timer1_Timer()
    Dim FF As Long
    Randomize
    FF = Int(Rnd * 100)
    
    Dim F2 As Long
    F2 = FF Mod 6
    Select Case F2
        Case 0
            'mostrar cada una de las imagenes donde corresponde!
            imF = ExtraData.GetImagePath("vumetroapagado")
            imgVUMSel.Picture = LoadPicture(imF)
            imgVumSel2.Picture = LoadPicture(imF)
            
            imF = ExtraData.GetImagePath("touchizqnormal")
            imgTouchSel.Picture = LoadPicture(imF)
            
            imF = ExtraData.GetImagePath("touchderechanormal")
            imgTouchSel2.Picture = LoadPicture(imF)
        Case 1
            'mostrar cada una de las imagenes donde corresponde!
            imF = ExtraData.GetImagePath("vumetroprendido")
            imgVUMSel.Picture = LoadPicture(imF)
            imgVumSel2.Picture = LoadPicture(imF)
            
            imF = ExtraData.GetImagePath("touchizqapretado")
            imgTouchSel.Picture = LoadPicture(imF)
            
            imF = ExtraData.GetImagePath("touchderechaapretado")
            imgTouchSel2.Picture = LoadPicture(imF)
        Case 2
            'mostrar cada una de las imagenes donde corresponde!
            imF = ExtraData.GetImagePath("vumetroapagado")
            imgVUMSel.Picture = LoadPicture(imF)
            imgVumSel2.Picture = LoadPicture(imF)
            
            imF = ExtraData.GetImagePath("touchizqapretado")
            imgTouchSel.Picture = LoadPicture(imF)
            
            imF = ExtraData.GetImagePath("touchderechaapretado")
            imgTouchSel2.Picture = LoadPicture(imF)
        Case 3
            'mostrar cada una de las imagenes donde corresponde!
            imF = ExtraData.GetImagePath("vumetroprendido")
            imgVUMSel.Picture = LoadPicture(imF)
            imgVumSel2.Picture = LoadPicture(imF)
            
            imF = ExtraData.GetImagePath("touchizqapretado")
            imgTouchSel.Picture = LoadPicture(imF)
            
            imF = ExtraData.GetImagePath("touchderechanormal")
            imgTouchSel2.Picture = LoadPicture(imF)
        Case 4
            imF = ExtraData.GetImagePath("touchizqnormal")
            imgTouchSel.Picture = LoadPicture(imF)
            
            imF = ExtraData.GetImagePath("touchderechaapretado")
            imgTouchSel2.Picture = LoadPicture(imF)
            
        Case 5
            imF = ExtraData.GetImagePath("touchderechanormal")
            imgTouchSel2.Picture = LoadPicture(imF)
            
            imF = ExtraData.GetImagePath("touchizqapretado")
            imgTouchSel.Picture = LoadPicture(imF)
            
    End Select
    
'    imF = ExtraData.GetImagePath("marcodiscocomun")
'    imgTapaSel.Picture = LoadPicture(imF)
'    ii3.Picture = imgTapaSel.Picture
'    ii4.Picture = imgTapaSel.Picture
'    ii5.Picture = imgTapaSel.Picture
'    ii6.Picture = imgTapaSel.Picture
'    imgTapaSel2.Picture = LoadPicture(imF)
'
'    imF = ExtraData.GetImagePath("marcodiscoelegido")
'    If F2 = 0 Then imgTapaSel2.Picture = LoadPicture(imF)
'    If F2 = 1 Then imgTapaSel.Picture = LoadPicture(imF)
'    If F2 = 2 Then ii3.Picture = LoadPicture(imF)
'    If F2 = 3 Then ii4.Picture = LoadPicture(imF)
'    If F2 = 4 Then ii5.Picture = LoadPicture(imF)
'    If F2 = 5 Then ii6.Picture = LoadPicture(imF)
    
End Sub
