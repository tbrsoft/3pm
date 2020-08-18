VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmCarritoReview 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin tbrFaroButton.fBoton btOKPachaCartDel 
      Height          =   840
      Left            =   3090
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3630
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1482
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Quitar"
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin VB.Image tUPDel 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   4800
      Top             =   3870
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image tDownDel 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   2280
      Top             =   3870
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-Contenido de la compra-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   825
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   7875
   End
   Begin VB.Image CD1 
      Height          =   1425
      Index           =   0
      Left            =   300
      Stretch         =   -1  'True
      Top             =   690
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Label teX1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "SALIR. He terminado de eliminar"
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
      Height          =   615
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   2130
      Visible         =   0   'False
      Width           =   2445
   End
End
Attribute VB_Name = "frmCarritoReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sSel As Long 'seleccion elegida

Private Sub btOKPachaCartDel_Click()
    Form_KeyDown TeclaOK, 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim II As Long
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        
        Case TeclaCarrito, TeclaOK 'borrar el elegido
            If sSel = 0 Then
                'salir
                Unload Me
                
                frmCarrito.Show 1
            Else
                Carrito.CleanSel sSel
                LoadLista
            End If
            
        Case TeclaIZQ
            II = sSel - 1
            If II < 0 Then II = teX1.Count - 1
            SELD II
        Case TeclaDER
            II = sSel + 1
            If II > (teX1.Count - 1) Then II = 0
            SELD II
    End Select
End Sub

Private Sub SELD(I As Long)
    CD1(sSel).BorderStyle = 0
    teX1(sSel).ForeColor = vbWhite
    
    CD1(I).BorderStyle = 1
    teX1(I).ForeColor = vbYellow
    
    sSel = I
End Sub

Private Sub Form_Load()
    EsSaving = True 'para que no se lance ni el protector ni temas al azar!
    
    Label3.Caption = "-Contenido de la compra-" + vbCrLf + _
        "Elija la selección a eliminar"
        
    Dim IMF As String
    
    IMF = ExtraData.getDef.getImagePath("touchderechanormal")
    tUPDel.Picture = LoadPicture(IMF)

    IMF = ExtraData.getDef.getImagePath("touchizqnormal")
    tDownDel.Picture = LoadPicture(IMF)
    
    tUPDel.BorderStyle = 0
    tDownDel.BorderStyle = 0
End Sub

Private Sub Form_Resize()
    Label3.Left = 0
    Label3.Width = Me.Width
    
    Label3.Caption = "Elija el elemento a eliminar y presione 'OK' o la tecla de carrito para eliminarlo"
    
    'que quede igual!
    tDownDel.Top = frmIndex.picFondoPacha.Top + frmIndex.t1.Top
    tDownDel.Left = frmIndex.picFondoPacha.Left + frmIndex.t1.Left
    
    tUPDel.Top = frmIndex.picFondoPacha.Top + frmIndex.t3.Top
    tUPDel.Left = frmIndex.picFondoPacha.Left + frmIndex.t3.Left
    'este boton es más grande!
    'btOKPachaCart.Top = frmIndex.picFondoPacha.Top + frmIndex.btOKPacha.Top
    btOKPachaCartDel.Top = Me.Height - btOKPachaCartDel.Height + 60
    btOKPachaCartDel.Left = frmIndex.picFondoPacha.Left + frmIndex.btOKPacha.Left
    btOKPachaCartDel.Width = frmIndex.btOKPacha.Width
    'aqui tengo mas lugar y necesito más texto
    'btOKPachaCart.Height = frmIndex.btOKPacha.Height
    btOKPachaCartDel.Caption = "Quitar"
    
    tDownDel.Visible = True
    tUPDel.Visible = True
    btOKPachaCartDel.Visible = True
        
    LoadLista
    
End Sub

Private Function LoadLista()
    
    'descargar todo por las dudas
    Dim H As Long
    For H = 1 To teX1.Count - 1
        Unload CD1(H)
        Unload teX1(H)
    Next H
    
    sSel = 0
    
    
    
    'ver si se puede mostrar todo. Si quiedara muy chiquito ponemos algun mensaje
    Dim MinH As Long 'minimo de alto que muestro
    
    tERR.Anotar "daaq", Carrito.GetFileCant
    
    CD1(0).Top = Label3.Top + Label3.Height + 160
    teX1(0).Top = CD1(0).Top + CD1(0).Height
    
    CD1(0).Left = Label3.Left + 120
    teX1(0).Left = CD1(0).Left
    
    CD1(0).Visible = True
    teX1(0).Visible = True
    
    If Carrito.GetFileCant > 0 Then
        For H = 1 To Carrito.GetFileCant
            ShowElem H
        Next H
    End If
    
    SELD 0
End Function

Private Function ShowElem(I As Long)
    On Local Error GoTo MER
    
    Load CD1(I)
    Load teX1(I)
    
    Dim IMG As String
    IMG = Carrito.GetElementPath(I) + "tapa.jpg"
    If fso.FileExists(IMG) Then
        If FileLen(IMG) > TamanoTapaPermitido * 1024 Then
            GoTo TapaDef3
        End If
        tERR.Anotar "daar", IMG
        CD1(I).Picture = LoadPicture(IMG)
    Else
TapaDef3:
        'ver si tiene programado una imagen de SL
        If K.sabseee(dcr("1Vx0YVGhEoIisHPLAZMHXw==")) >= Supsabseee Then
            If fso.FileExists(GPF("tddp322")) Then
                IMF = GPF("tddp322")
                tERR.Anotar "daas", IMF
                CD1(I).Picture = LoadPicture(IMF)
            Else
                tERR.Anotar "daat"
                CD1(I).Picture = frmIndex.imgTapaDefBUP.Picture
            End If
        Else
            tERR.Anotar "daau"
            CD1(I).Picture = frmIndex.imgTapaDefBUP.Picture
        End If
    End If
    
    teX1(I).Caption = Carrito.GetElementName(I)
     
    If I = 1 Then 'si es el primero dar la primera referencia
        CD1(I).Top = CD1(I - 1).Top
        teX1(I).Top = teX1(I - 1).Top
        
        CD1(I).Left = CD1(I - 1).Left + CD1(I - 1).Width + 90
        teX1(I).Left = teX1(I - 1).Left + teX1(I - 1).Width + 90
    Else
        If CD1(I - 1).Left + (2 * CD1(0).Width) > Me.Width Then
            'empezar otro renglon
            CD1(I).Top = teX1(I - 1).Top + teX1(I - 1).Height + 60
            teX1(I).Top = CD1(I).Top + CD1(I).Height
            
            CD1(I).Left = CD1(1).Left
            teX1(I).Left = teX1(1).Left
        Else
            CD1(I).Top = CD1(I - 1).Top
            teX1(I).Top = teX1(I - 1).Top
            
            CD1(I).Left = CD1(I - 1).Left + CD1(I - 1).Width + 90
            teX1(I).Left = teX1(I - 1).Left + teX1(I - 1).Width + 90
        End If
    End If
    
    CD1(I).Visible = True
    teX1(I).Visible = True
    
    tERR.Anotar "daav"
    
    Exit Function
MER:
    tERR.AppendLog tERR.ErrToTXT(Err), "cpCC2"
    Resume Next
End Function

Private Sub tDownDel_Click()
    Form_KeyDown TeclaIZQ, 0
End Sub

Private Sub tUPDel_Click()
    Form_KeyDown TeclaDER, 0
End Sub
