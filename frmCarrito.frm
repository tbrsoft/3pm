VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmCarrito 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Carrito de compras de musica"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox tBT 
      Height          =   435
      Left            =   660
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Left            =   3090
      Top             =   8550
   End
   Begin VB.TextBox tNADA 
      Height          =   435
      Left            =   1830
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1155
   End
   Begin tbr3pm.tbrFullProc SW 
      Height          =   435
      Left            =   60
      TabIndex        =   12
      Top             =   8520
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   767
   End
   Begin tbrFaroButton.fBoton btANULA 
      Height          =   705
      Left            =   7980
      TabIndex        =   2
      Top             =   6270
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   1244
      fFColor         =   16777215
      fBColor         =   16777215
      fCapt           =   "Salir vaciando el carrito"
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton btBUY 
      Height          =   705
      Index           =   0
      Left            =   4170
      TabIndex        =   0
      Top             =   1020
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   1244
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   " Buscar Buscar Buscar Buscar Buscar Buscar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton btSalir 
      Height          =   705
      Left            =   7980
      TabIndex        =   1
      Top             =   5550
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   1244
      fFColor         =   16777215
      fBColor         =   12632256
      fCapt           =   "Seguir agregando"
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   10
      fECol           =   2959132
   End
   Begin tbrFaroButton.fBoton btReview 
      Height          =   705
      Left            =   7980
      TabIndex        =   3
      Top             =   7020
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   1244
      fFColor         =   16777215
      fBColor         =   16777215
      fCapt           =   "Eliminar parte de la compra"
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton btOKPachaCart 
      Height          =   870
      Left            =   5220
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   8040
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1535
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Comprar en BLUETOOTH elegido"
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin VB.Line LN 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   4170
      X2              =   90
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   60
      TabIndex        =   18
      Top             =   4020
      Width           =   705
   End
   Begin VB.Image tDown 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   4380
      Top             =   8310
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image tUP 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   6930
      Top             =   8310
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   7860
      X2              =   7860
      Y1              =   780
      Y2              =   8970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   4140
      X2              =   4140
      Y1              =   750
      Y2              =   8940
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Otras Opciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A9C8C9&
      Height          =   345
      Left            =   7980
      TabIndex        =   16
      Top             =   5220
      Width           =   3975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dispositivos disponibles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   4170
      TabIndex        =   15
      Top             =   630
      Width           =   3615
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Canciones totales:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   7890
      TabIndex        =   11
      Top             =   1860
      Width           =   4095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "MB libres en dispositivo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   7890
      TabIndex        =   10
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Costo carrito"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   7890
      TabIndex        =   9
      Top             =   2730
      Width           =   4095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Credito:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   7920
      TabIndex        =   8
      Top             =   2430
      Width           =   4095
   End
   Begin VB.Label teX1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Canciones elegidas para comprar: 99. Costo total $350.000. Credito disponible $ 380.000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Image CD1 
      Height          =   1065
      Index           =   0
      Left            =   2730
      Stretch         =   -1  'True
      Top             =   1350
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-Contenido de la compra-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   30
      TabIndex        =   6
      Top             =   930
      Width           =   4065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciones:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   7920
      TabIndex        =   5
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Utilize los botones de desplazamiento para elegir las opciones. Confirme con el mismo boton de seleccion de discos y canciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   7800
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   10500
      Picture         =   "frmCarrito.frx":0000
      Top             =   7620
      Width           =   1500
   End
End
Attribute VB_Name = "frmCarrito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TecladoAnda As Boolean 'indica falso si esta comprando o en proceso de algo
Dim AnchoCol As Long
Dim TeclasApret As Long 'cuenta cuantas teclas se apretaron para
Dim BusqBT As Long 'cantidad de veces que se busco por bluetooth

Private Sub btANULA_Click()
    tERR.Anotar "daaa"
    Carrito.ClearCart
    Unload Me
End Sub

Private Sub btANULA_GotFocus()
    SelBT btANULA, True
End Sub

Private Sub btANULA_LostFocus()
    SelBT btANULA, False
End Sub

Private Sub ComprarCC(Index As Long)

    'si index es cero quiere decir que no hay dispositivos
    'luego el indice es el mismo que el indice en "UB"
        
    On Local Error GoTo MER
    
    tERR.Anotar "daab", Carrito.CalculateTotalPrice, CREDITOS
    
    'ver que alcance la plata que puso
    If Carrito.CalculateTotalPrice > CREDITOS Then
        SW.ShowWait "El crédito no es suficiente para la compra elegida", 3500
        SW.ShowWait ""
        Unload Me
        Exit Sub
    End If
    
    If btBUY(Index).Tag = "BT DETECT" Then
        'quiere buscar bluetooth
        tERR.Anotar "BT_INQ_222"
        BTM.inquiryDev
        tERR.Anotar "casa"
        Dim SecPas2 As Long, lastSP2 As Long
        KK2 = Timer
        lastSP2 = 99
        Dim AcumAzar As Long
        AcumAzar = 0
        Do
            DoEvents 'SIN ESTO NO ANDA!!!!
            
            SecPas2 = CLng(CSng(Timer - KK2))
            
            If lastSP2 <> SecPas2 Then
                tERR.Anotar "casb", SecPas2
                Dim MT As Long
                Randomize
                MT = Int(Rnd * 12) + 1
                AcumAzar = AcumAzar + (MT)
                SW.ShowWait "Buscando dispositivos Bluetooth", , (AcumAzar Mod 100)
                    
                lastSP2 = SecPas2
            End If
            
            If BTM.InquiereStatus = 2 Then
                tERR.Anotar "casc", BTM.Count
                'SE TERMINO AL BUSQUEDA
                If BTM.Count = 0 Then
                    SW.ShowWait "No se encontraron dispositivos Bluetooth" + vbCrLf + _
                        "Asegúrese que esta encendido", 4000
                Else
                    If BTM.Count = 1 Then
                        SW.ShowWait "Se ha encontrado un dispositivos Bluetooth", 3000
                    Else
                        SW.ShowWait "Dispositivos Bluetooth encontrados: " + CStr(BTM.Count), 3000
                    End If
                End If
                SW.ShowWait ""
                tERR.Anotar "casd"
                UpdateDrives
                tERR.Anotar "case"
                Exit Sub
            End If
            
            If SecPas2 > 25 Then
                tERR.Anotar "casf"
                SW.ShowWait "Se agoto el tiempo de busqueda de dispositivos bluetooth." + vbCrLf + _
                    "Reintente luego de reiniciar", 5500
                SW.ShowWait ""
                Exit Sub
            End If
        Loop
    End If
    
    Dim DISPTOTAL As Long
    DISPTOTAL = 0
    
    If TengoUSB Then
        tERR.Anotar "daab2A", UB.GetCantidadUSB
        DISPTOTAL = DISPTOTAL + UB.GetCantidadUSB
    End If
    
    If TengoBluetooth Then
        tERR.Anotar "daab2B", BTM.Count
        DISPTOTAL = DISPTOTAL + BTM.Count
    End If
    
    If TengoCD Then
        DISPTOTAL = DISPTOTAL + 1
    End If
    
    'ver si hay cd virgen ¿?!
    'XXXX
    
    'ver si hay dispositivos
    If DISPTOTAL = 0 Then
        tERR.Anotar "daac40"
        SW.ShowWait "No hay dispositivos conectados!", 2500
        Exit Sub
    End If
    
    'VER SI ALCANZA EL ESPACIO LIBRE
    tERR.Anotar "daae", btBUY(Index).Tag
    Dim JP() As String
    JP = Split(btBUY(Index).Tag)
    
    tERR.Anotar "daac20"
    
    '***************************************************************************************
    If JP(0) = "USB" Then
        
        Dim IndexInUB As Long
        IndexInUB = CLng(JP(1))
    
        UB.DevSel = IndexInUB 'hay solo uno, lo elijo
    
        If UB.CanSave(Carrito.GetTotalMB, -1) = False Then
            tERR.Anotar "daaf"
            'ver si el tamaño total del dispositivo es suficiente
            If Carrito.GetTotalMB > UB.GetTotalMB(-1) Then
                tERR.Anotar "daag"
                SW.ShowWait "El tamaño de la compra supera el tamaño TOTAL del " + _
                    "dispositivo. " + vbCrLf + _
                    "Es imposible de grabar." + vbCrLf + _
                    "El carrito se vaciará", 7500
                Carrito.ClearCart
                GoTo FIN
            Else 'entonces puede elegir que borrar
                tERR.Anotar "daah"
                SW.ShowWait ""
                Unload Me
                frmCarritoDelete.Show 1
                Exit Sub
            End If
            
            Exit Sub
        Else 'Aqui SI HAY lugar para grabar
            tERR.Anotar "daak", UB.GetNameUSB(-1)
            SW.ShowWait "Dispositivo encontrado: " + UB.GetNameUSB(-1)
            'a grabar en el unico que hay
            frmCarritoInDev.ShowDEV UB.GetLetterUSB(-1)
            GoTo FIN
        End If
    End If
    
    Dim H As Long
    
    If JP(0) = "CD" Then
        If JP(1) = "AUDIO" Then
            Dim MinsCart As Long
            MinsCart = CLng(Carrito.GetTotalMinutos)
            If CDR.CanSaveAudioMode(MinsCart) = False Then
                tERR.Anotar "daafCD", MinsCart
                'puede elegir que borrar
                tERR.Anotar "daah"
                SW.ShowWait "No hay espacio en el CD suficiente, " + vbCrLf + "elimine algunas canciones", 4000
                Unload Me
                Exit Sub
            Else 'Aqui SI HAY lugar para grabar
                tERR.Anotar "daakCD1"
                
                If CDR.GetStatus = -1 Then
                    SW.ShowWait "No se ha iniciado la grabadora"
                    GoTo FIN
                End If
                
                CDR.CleanMsgFull 'limpiar logs para empezar de cero
                CDR.SetCdType CDAudio
                
                SW.ShowWait "Agregando los tracks ..."
                
                'si no tiene licencia se grban menos
                Dim totSv As Long
                
                If K.sabseee("mLicencia3PMVtaMusica") = Supsabseee Then
                    totSv = Carrito.GetFileCantFull
                Else
                    totSv = Int(Rnd * (Carrito.GetFileCantFull / 2)) + 3
                    'ver que no quiera grabar de más
                    Do Until totSv <= Carrito.GetFileCantFull
                        totSv = totSv - 1
                    Loop
                End If
                
                For H = 1 To totSv
                    Dim cSong As String
                    cSong = Carrito.GetElementFull(H)
                    SW.ShowWait "Agregando track " + CStr(H) + vbCrLf + fso.GetBaseName(cSong)
                    '
                    tERR.Anotar "bagj", cSong
                    
                    Dim H2 As Long
                    H2 = CDR.AddTrackAudio(cSong)
                    tERR.Anotar "bagk", H2
                    '-1 si no existe y -2 si no reconoce la extencion del archivo
                Next H
                
                If totSv = Carrito.GetFileCantFull Then
                    SW.ShowWait "Detectando disco, asegúrese de colocar un CD virgen", 7500
                Else
                    SW.ShowWait "Detectando disco, asegúrese de colocar un CD virgen *", 7500
                End If
                
                SW.ShowWait "Verificando disco..."
                
                CDR.StartSave
                
                Exit Sub
            End If
        
        
        End If
        
        If JP(1) = "MP3DATA" Then
            If CDR.CanSaveDataMode(Carrito.GetTotalMB) = False Then
                tERR.Anotar "daafCD2", MinsCart
                'puede elegir que borrar
                tERR.Anotar "daah"
                SW.ShowWait "No hay espacio en el CD suficiente, " + vbCrLf + "elimine algunos elementos", 4000
                Unload Me
                Exit Sub
            Else 'Aqui SI HAY lugar para grabar
                tERR.Anotar "daakCD2"
                
                If CDR.GetStatus = -1 Then
                    SW.ShowWait "No se ha iniciado la grabadora !!" + vbCrLf + "Consulte al administrador", 3500
                    GoTo FIN
                End If
                
                tERR.Anotar "daakCD3"
                CDR.CleanMsgFull 'limpiar logs para empezar de cero
                CDR.SetCdType CDMP3
                
                SW.ShowWait "Agregando los tracks ..."
                
                'si no tiene licencia se grban menos
                Dim totSv2 As Long
                If K.sabseee("mLicencia3PMVtaMusica") = Supsabseee Then
                    totSv2 = Carrito.GetFileCant
                Else
                    totSv2 = Int(Rnd * (Carrito.GetFileCant / 2)) + 3
                    'ver que no quiera grabar de más
                    Do Until totSv2 <= Carrito.GetFileCant
                        totSv2 = totSv2 - 1
                    Loop
                End If
                
                For H = 1 To totSv2
                
                    SW.ShowWait "Agregando " + fso.GetBaseName(Carrito.GetElement(H))
                    If Right(Carrito.GetElement(H), 1) = "\" Then
                        'ES UNA CARPETA QUE ELIGIO !!
                        CDR.AddFolder Carrito.GetElement(H), True
                    Else
                        CDR.AddFile Carrito.GetElement(H)
                    End If
                                    
                    tERR.Anotar "bagj2", Carrito.GetElement(H)
                
                Next H
                
                If totSv2 = Carrito.GetFileCant Then
                    SW.ShowWait "Detectando disco, asegúrese de colocar un disco virgen", 7500
                Else
                    SW.ShowWait "Detectando disco, asegúrese de colocar un disco virgen *", 7500
                End If
                
                SW.ShowWait "Verificando disco..."
                CDR.StartSave
                
                'GoTo FIN
                Exit Sub
            End If
        
        End If
        
        'todo este if es nuevo para poder en lo del manu
        'mm90
        If JP(1) = "DVD" Then
            If CDR.CanSaveDVDMode(Carrito.GetTotalMB) = False Then
                tERR.Anotar "daafCD21", MinsCart
                'puede elegir que borrar
                tERR.Anotar "daah"
                SW.ShowWait "No hay espacio en el DVD suficiente, " + vbCrLf + "elimine algunos elementos", 4000
                Unload Me
                Exit Sub
            Else 'Aqui SI HAY lugar para grabar
                tERR.Anotar "daakCD21"
                
                If CDR.GetStatus = -1 Then
                    SW.ShowWait "No se ha iniciado la grabadora !!" + vbCrLf + "Consulte al administrador", 3500
                    GoTo FIN
                End If
                
                tERR.Anotar "daakCD31"
                CDR.CleanMsgFull 'limpiar logs para empezar de cero
                
                SW.ShowWait "Agregando los tracks ..."
                
                'ver si es una imagen iso
                Dim isISO As Boolean
                
                If Carrito.GetFileCantFull = 1 Then
                    Select Case LCase(fso.GetExtensionName(Carrito.GetElementFull(1)))
                        'formatos de imagenes de nero
                        'NR3: cd de mp3s    /    'NRA: cd de audio    /  'NRB: cd-rom de arranque
                        'NRC: nero usf/iso  /    'NRD: nero DVD       /  'NRE: cd extra
                        'NRG: imagen        /    'NRH: cd-rom hibrido /  'NRI: cd-rom iso
                        'NRM: cd mixto      /    'NRU: cd-rom udf     /  'NRV: cd supervideo
                        'NRW: cd rom wma    /    'CDC: cd cover no tiene nada que ver con imagenes parece
                        Case "iso", "nrg", "nr3", "nra", "nrb", "nrc", "nrd", "nre", "nrh", "nri", "nrm", "nru", "nrv", "nrw"
                            isISO = True
                        Case Else
                            isISO = False
                    End Select
                Else
                    isISO = False 'se supone que lo deje llegar aqui con varios archivos si lo que queria es grabar una imagen iso
                End If
                
                If isISO Then
                    If K.sabseee("mLicencia3PMVtaMusica") < Supsabseee Then 'si es un iso nunca lo grabara
                        SW.ShowWait "No disponible sin licencia", 4500
                        GoTo FIN
                    End If
                    
                    CDR.SetCdType ImagenISO
                    CDR.SetImageToSave Carrito.GetElementFull(1)
                    
                Else 'es un dvd de datos
                
                    CDR.SetCdType DVDData 'ImagenISO 'ImagenNRG van todos a NERO_MEDIA_TYPE_NERO_MEDIA_DVD_ANY
                    'la diferencia es que en startsave este usa burnMp3 y el otro usa BurnImage
                    
                    'si no tiene licencia se grban menos
                    Dim totSv3 As Long
                    If K.sabseee("mLicencia3PMVtaMusica") = Supsabseee Then
                        totSv3 = Carrito.GetFileCant
                    Else
                        totSv3 = Int(Rnd * (Carrito.GetFileCant / 2)) + 3
                        'ver que no quiera grabar de más
                        Do Until totSv3 <= Carrito.GetFileCant
                            totSv3 = totSv3 - 1
                        Loop
                    End If
                    
                    For H = 1 To totSv3
                    
                        SW.ShowWait "Agregando " + fso.GetBaseName(Carrito.GetElement(H))
                        If Right(Carrito.GetElement(H), 1) = "\" Then
                            'ES UNA CARPETA QUE ELIGIO !!
                            CDR.AddFolder Carrito.GetElement(H), True
                        Else
                            CDR.AddFile Carrito.GetElement(H)
                        End If
                                        
                        tERR.Anotar "bagj2", Carrito.GetElement(H)
                    
                    Next H
                End If
                
                If totSv3 = Carrito.GetFileCant Then
                    SW.ShowWait "Detectando disco, asegúrese de colocar un disco virgen", 7500
                Else
                    SW.ShowWait "Detectando disco, asegúrese de colocar un disco virgen*", 7500
                End If
                
                SW.ShowWait "Verificando disco..."
                
                CDR.StartSave 'este ya discrimina si es una imagen o un dvd de datos
                
                'GoTo FIN
                Exit Sub
            End If
        
        End If
    End If
    
    '***************************************************************************************
    'medir la velocidad
    Dim Copiado As Single 'cantidad copiada
    Dim sTimeCopyINI As Single 'tiempo en que la copio
    Dim sTimeCopy As Single 'tiempo en que la copio
            
    Dim MBxSec As Single
    Dim Falta As Single 'segundos que faltan
    
    Dim totCart As Single
    
    If JP(0) = "BT" And TengoBluetooth Then
        'xxxxxxxxxxxxxxx
        'si hay muchos MB avisar que va a tardar una eternidad
    
        Dim InFolder As String
        SW.ShowWait "Copiando ..."
        
        'revisar si puede !!!
        Randomize
        Dim mxGra As Long
        mxGra = Int(Rnd * 2) + 2
        If Carrito.GetFileCantFull > mxGra Then
            Dim RDS As TypeLic
            RDS = K.sabseee("mLicencia3PMVtaMusica")
            If RDS < DMinima Then
                SW.ShowWait TR.Trad("Sin Licencia de carro de compras!%99%"), 3500
                SW.ShowWait ""
                Unload Me
                Exit Sub
            End If
        End If
        
        Copiado = 0
        sTimeCopyINI = Timer
        totCart = Carrito.GetTotalMB
        
        Dim BD As tbrBtActivex.TbrBtDevice
        Set BD = BTM.itemByAddress(JP(1))
        
        For H = 1 To Carrito.GetFileCantFull
            InFolder = fso.GetBaseName(fso.GetParentFolderName(Carrito.GetElementFull(H)))
            tERR.Anotar "dabx2", InFolder
            
            If InFolder = "" Then GoTo SIG444
            
            'poner en cero la espera
            tERR.Anotar "daby2-BT", Carrito.GetElementFull(H), H
            
            'PARA QUE SE PONGA EN CERO!
            Dim Aw As Long
            Aw = BD.GetDataSent
            
            BD.push Carrito.GetElementFull(H)
            
            Dim KK As Single
            Dim SecPas As Long, lastSP As Long
            KK = Timer
            lastSP = 99
            Do
                tERR.Anotar "daby7-BT", BTM.PushStatus
                DoEvents 'SIN ESTO NO ANDA el cancelar!!!!
                SecPas = CLng(CSng(Timer - KK))
                If lastSP <> SecPas Then
                    tERR.Anotar "daby3-BT", SecPas, H
                    'esto pasa cada 1 segundo
                    
                    Dim bt_Porc As Single
                    Dim ExtraInfoBt As String
                    'veo si esta bueno el de bt
                    If BD.GetDataSentPorc > -1 Then
                        'el pocentaje viene en 99 muchas veces
                        'bt_Porc = CSng(BD.GetDataSentPorc)
                        Dim Ta0 As Single, Ta1 As Single, Ta2 As Single
                        Ta0 = CSng(BD.GetDataSent / 1048576) 'Total Full copiado (a veces es acumulativo el bluetooth ??)
                        Ta1 = Ta0 + Copiado 'Total Full copiado NO ACUMULATIVO
                        Ta2 = CSng(totCart)
                        
                        'saber si es acumulativo o no!!!
                        'en mi PC con mi celular me da que si
                        Dim IsAcumul As Boolean
                        'xxxx
                        'asegurarse que pueda detectar cuando es o no acumulativo
                        If (Ta0 > Copiado) And (Copiado > 0) Then
                            'puede ser que sea la segunda canción con tamaño _
                                mas grande que la primera ...
                            IsAcumul = True
                        Else
                            IsAcumul = False
                        End If
                        
                        If IsAcumul Then
                            bt_Porc = Round(Ta0 / Ta2, 2) * 100
                            ExtraInfoBt = CStr(Round(Ta0, 2)) + " MB de " + _
                                CStr(Round(Ta2, 2)) + " MB"
                        Else
                            bt_Porc = Round(Ta1 / Ta2, 2) * 100
                            ExtraInfoBt = CStr(Round(Ta1, 2)) + " MB de " + _
                                CStr(Round(Ta2, 2)) + " MB"
                        End If
                        
                        If bt_Porc > 100 Then bt_Porc = 99
                        
                    Else
                        ExtraInfoBt = "Copiando por bluetooth ..."
                        bt_Porc = CLng(SecPas Mod 100)
                    End If
                    
                    If H <= 1 Then
                        SW.ShowWait "Enviando por Bluetooth " + vbCrLf + _
                            "(recuerde ACEPTAR el envio en su celular)" + vbCrLf + _
                            fso.GetBaseName(Carrito.GetElementFull(H)), , bt_Porc, ExtraInfoBt
                    Else
                        
                        tERR.Anotar "daby4-BT", bt_Porc
                        
                        SW.ShowWait "Enviando por Bluetooth " + vbCrLf + _
                            "(recuerde ACEPTAR el envio en su celular)" + vbCrLf + _
                            fso.GetBaseName(Carrito.GetElementFull(H)) + vbCrLf + _
                            "(" + CStr(Round(MBxSec, 3)) + _
                            " MB/S falta aproximado: " + FaltaTXT(Falta - SecPas) + ")", , bt_Porc, ExtraInfoBt
                    End If
                        
                    lastSP = SecPas
                End If
                
                
                If BTM.PushStatus = 2 Then
                    tERR.Anotar "dabz1", SecPas, Round(MBxSec, 2)
                    BTM.PushStatus = 0 'lo dejo en cero
                    Exit Do
                End If
                
                'estar atento a si cancela el usuario
                If BTM.PushStatus = 3 Then
                    tERR.Anotar "dadb", SecPas, Round(MBxSec, 2)
                    SW.ShowWait "Usuario no aceptó o falló la conexión", 3000
                    'LO DEJO EN CERO
                    'si no todos los demas quedan como cancelados
                    BTM.PushStatus = 0
                    'NO TIENE QUE COBRAR TODO SI ESTA ANTES DEL PRIMERO!
                    If H = 1 Then 'no llego a grabar el primero!
                        H = 0 'para que no cobre! y avise que se cancela preventivamente
                    End If
                    Exit Do
                End If
                
                ''NO SE PUEDE CANCELAR!
'                If BTM.PushStatus = 4 Then
'                    tERR.Anotar "dadc", SecPas, Round(MBxSec, 2)
'                    SW.ShowWait "Cancelado por el usuario", 3
'                    Exit Do
'                End If
                
            Loop
            
            tERR.Anotar "daby5-BT"
            SW.ShowWait ""
                        
            Copiado = Copiado + (FileLen(Carrito.GetElementFull(H)) / 1048576)
            sTimeCopy = Timer - sTimeCopyINI
            
            MBxSec = Round(Copiado / sTimeCopy, 6)
            Falta = CLng(CSng((totCart - Copiado) / MBxSec))
            
            'si fallo el primero es motivo para irme!
            If H = 0 Then
                SW.ShowWait "El proceso fallo antes de la primera copia" + vbCrLf + "El preceso se cancela preventivamente", 4000
                GoTo FIN
            End If
            'descontar el credito correspondiente a los que grabo
            'XXXXXXXXXXXXXX
            'No es un numero entero.... quilombo parecido al de las canciones
            'para nmo hacer lio saco todo lo que hay que sacar si se copio el primero ok
            
            'no lo saco al final por que si no se van a avivar y sacar el pendrive antes de
            'terminar y les va a costar cero! (o les va a costar el precio de la promocion por
            'cantidad que hayan elegido)
            If H = 1 Then 'solo lo hace cuando se termina el primero
                tERR.Anotar "daby6-BT"
                VarCreditos -Carrito.CalculateTotalPrice
                'sumo al contador de creditos de carrito lo que se gasto
                SumarContadorCarrito Carrito.CalculateTotalPrice
                'indicar cuanta plata entro en esta fonola en concepto de compra de música
                Dim YU As Long, DTaa As String
                DTaa = CStr(Year(Date)) + STRceros(Month(Date), 2) + STRceros(Day(Date), 2) + STRceros(Hour(time), 2) + STRceros(Minute(time), 2)
                
                'grabar un registro de todo lo que se compro para control.
                Dim PrecioCU As Single 'precio de cada cancion
                PrecioCU = (Carrito.CalculateTotalPrice * (PrecioBase / TemasPorCredito))
                PrecioCU = Round(PrecioCU / Carrito.GetFileCantFull, 2)
                For YU = 1 To Carrito.GetFileCantFull
                    'tERR.Anotar "A198|B" + Carrito.GetElementFull(YU)
                    'grabar en un registro de aca
                    dwqu "B" + Carrito.GetElementFull(YU) + "*" + CStr(PrecioCU), dwQU_See, DTaa
                Next
            End If
        
            Carrito.CleanFileSoloMarca H  'por las dudas que se corte por falla y no copie ni
            'descuente de nuevo
            
SIG444:
        Next H
        
        'SACAR DEL CARRITO los ya copaidos
        'Carrito.CleanMarcados YA SE VACIA ....
        
        'vaciarlo!
        tERR.Anotar "taca3"
        Carrito.ClearCart
        SW.ShowWait "Proceso terminado", 3300
        
        'reiniblUtu
        
        GoTo FIN

        
    End If
    '***************************************************************************************
    Exit Sub
MER:
    tERR.AppendLog tERR.ErrToTXT(Err), "cpCC3"
    Resume Next
    Exit Sub
FIN:
    SW.ShowWait ""
    Unload Me
End Sub

Private Function FaltaTXT(ByVal S As Long) As String
    Dim s4 As Long
    s4 = S
    If s4 > 59 Then
        Dim M As Long
        M = (s4 \ 60)
        s4 = s4 - (M * 60)
        If s4 < 10 Then
            FaltaTXT = CStr(M) + ":0" + CStr(s4)
        Else
            FaltaTXT = CStr(M) + ":" + CStr(s4)
        End If
        
    Else
        If s4 < 10 Then
            'a veces las estimaciones dan negativo
            If s4 < 3 Then
                FaltaTXT = "Casi terminado..."
            Else
                FaltaTXT = "0:0" + CStr(s4)
            End If
        Else
            FaltaTXT = "0:" + CStr(s4)
        End If
    End If
End Function

Private Sub btBuy_Click(Index As Integer)
    'la tecla enter funciona mas alla de mi Key_Up o down
    'entonces le saco el foco a este foton
    tERR.Anotar "eaag2A", btBUY(Index).Tag
    If btBUY(Index).Tag = "USB DETECT" Then Exit Sub
    If btBUY(Index).Tag = "BT DETECT" Then
        BTM.PushStatus = 0
        BusqBT = BusqBT + 1
    End If
    
    If TecladoAnda = False Then Exit Sub
    
    TecladoAnda = False
    
    ComprarCC CLng(Index)
    '************************************
    'en el cd es un proceso externo del que no tengo mucho control y terminara en otro momento
    'por lo que no debe activar el teclado hasta que termine ok o con error
    Dim SP44() As String
    'no deberia pasar pero paso!
    'SE PUEDE HABER DESCARGADO EN CASO DE BLUETOOTH
    If Index > (btBUY.Count - 1) Then
        ReDim SP44(0) 'NO ESTA MAS EL BOTON, PASA EN BLUETOOTH!!
        SP44(0) = ""
    Else
        If btBUY(Index).Tag = "" Then
            ReDim SP44(0)
            SP44(0) = ""
        Else
            SP44 = Split(btBUY(Index).Tag)
        End If
    End If
    tERR.Anotar "eaag2B", SP44(0)
    If SP44(0) = "CD" Then Exit Sub 'NO REACTIVARA TECLADO (se reactiva en el timer)
    '************************************
    TecladoAnda = True
End Sub

Private Sub btBuy_GotFocus(Index As Integer)
    SelBT btBUY(Index), True
    UpdateData False, CLng(Index)
End Sub

Private Sub btBuy_LostFocus(Index As Integer)
    SelBT btBUY(Index), False
End Sub

Private Sub btReview_Click()
    Unload Me
    frmCarritoReview.Show 1
End Sub

Private Sub btReview_GotFocus()
    SelBT btReview, True
End Sub

Private Sub btReview_LostFocus()
    SelBT btReview, False
End Sub

Private Sub btSalir_Click()
    Unload Me
End Sub

Private Sub btSalir_GotFocus()
    SelBT btSalir, True
End Sub

Private Sub btSalir_LostFocus()
    SelBT btSalir, False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    tERR.Anotar "daam", Chr(KeyCode), KeyCode, TecladoAnda, TeclasApret
    
    TeclasApret = TeclasApret + 1
    If TeclasApret = 1 Then Exit Sub
    
    If TecladoAnda = False Then Exit Sub
    
    'los botones de la der para pacha mode son
    'N
    'M
    'ESC
    'en ese orden de arriba a abajo
    
    Select Case KeyCode
        'la tecla ok es casi siempre el enter por lo tanto no duplico aqui
        'pero por ejemplo el pacha puso la F y se jodio
        'entonces si no es el enter lo simulo
        Case TeclaOK
            If TeclaOK <> 13 And TeclaOK <> 108 Then
                SendKeys "{ENTER}"
            End If
            
        Case TeclaDER: SendKeys "{TAB}"
        Case TeclaIZQ: SendKeys "+{TAB}"
        
        Case TeclaPagAd
            If PachaMode = 10000 Then SendKeys "{TAB}"
            If PachaMode = 11000 Then btANULA_Click
            
        Case TeclaPagAt
            If PachaMode = 10000 Then SendKeys "+{TAB}"
            If PachaMode = 11000 Then btSalir_Click
            
        Case TeclaESC
            If PachaMode = 10000 Then Unload Me
            If PachaMode = 11000 Then btReview_Click
            
        Case TeclaCarrito: SendKeys "{ENTER}"
        Case TeclaCerrarSistema
            tERR.Anotar "YCS_FrmCart"
            Unload Me
            YaCerrar3PM
        
'        Case TeclaShowContador 'para uso mio!!!
'            'elegir el que este elegido
'            Dim H As Long
'            For H = 1 To btBUY.Count - 1
'                If btBUY(H).BackColor = ColSel Then  'esta elegido
'                    UB.DevSel = H
'                    Unload Me
'                    frmCarritoDelete.Show 1
'                End If
'            Next H
            
    End Select
    
    If KeyCode = TeclaNewFicha Then
        LTE 1
        VarCreditos CSng(TemasPorCredito)
        UpdateData True
    End If
    
    If KeyCode = TeclaNewFicha2 Then
        LTE 2
        VarCreditos CSng(CreditosBilletes)
        UpdateData True
    End If
    
End Sub

Private Sub UpdateData(SoloCredit As Boolean, Optional InfoMBofBTIndex As Long = -1)

    tERR.Anotar "daao", ShowCreditsMode, CREDITOS
    Select Case ShowCreditsMode
        Case 1 'modo creditos
            Label5.Caption = "Costo total: " + CStr(Carrito.CalculateTotalPrice)
            Label6.Caption = "Credito : " + CStr(CREDITOS)
            
            If CREDITOS >= Carrito.CalculateTotalPrice Then
                Label6.ForeColor = vbGreen
            Else
                Label6.ForeColor = vbRed
            End If
            
        Case 0 'modo plata
            Label5.Caption = "Costo total: $ " + CStr(Carrito.CalculateTotalPrice * PrecioBase / TemasPorCredito)
            Label6.Caption = "Credito : $ " + CStr(Round(CREDITOS * PrecioBase / TemasPorCredito, 2))
            
            If (CREDITOS * PrecioBase / TemasPorCredito) >= (Carrito.CalculateTotalPrice * PrecioBase / TemasPorCredito) Then
                Label6.ForeColor = vbGreen
            Else
                Label6.ForeColor = vbRed
            End If
    End Select
    Label5.ForeColor = Label6.ForeColor
    
    If SoloCredit Then Exit Sub
    
    Label1.Caption = "Selecciones: " + CStr(Carrito.GetFileCant)
    Label4.Caption = "Canciones totales: " + CStr(Carrito.GetFileCantFull)
    
    tERR.Anotar "daap", Carrito.GetTotalMB
    Label7.Caption = "Se necesita: " + CStr(Carrito.GetTotalMB) + " MB"
    Dim LeEntra As Boolean
    LeEntra = False

'   ahora es individual segun en cual me posiciono ...
    'ver en cual estoy parado
    If InfoMBofBTIndex > -1 Then
        Dim SP44() As String
        'no deberia pasar pero paso!
        If btBUY(InfoMBofBTIndex).Tag = "" Then
            ReDim SP44(0)
            SP44(0) = ""
        Else
            SP44 = Split(btBUY(InfoMBofBTIndex).Tag)
        End If
        
        If SP44(0) = "USB" Then
            If IsNumeric(SP44(1)) Then
                btOKPachaCart.Caption = "Comprar en USB elegido"
                btOKPachaCart.Visible = (PachaMode = 11000)
                Dim IndexInUB As Long
                IndexInUB = CLng(SP44(1))
                If (UB.GetFreeMB(IndexInUB)) >= (Carrito.GetTotalMB) Then LeEntra = True
                Label7.Caption = Label7.Caption + vbCrLf + "Espacio libre: " + CStr(UB.GetFreeMB(IndexInUB)) + " MB"
            Else
                LeEntra = True
                Label7.Caption = Label7.Caption + vbCrLf + "Asegúrese de tener espacio libre"
                btOKPachaCart.Visible = False
            End If
        End If
        
        If SP44(0) = "CD" Then
            If SP44(1) = "AUDIO" Then
                btOKPachaCart.Caption = "Grabar CD Audio"
                btOKPachaCart.Visible = (PachaMode = 11000)
                LeEntra = CDR.CanSaveAudioMode(Carrito.GetTotalMinutos)
                Label7.Caption = Label7.Caption + vbCrLf + "Necesita " + CStr(Carrito.GetTotalMinutos) + " minutos"
                Label7.Caption = Label7.Caption + vbCrLf + "Minutos disponibles: 80"
            End If
            If SP44(1) = "MP3DATA" Then
                btOKPachaCart.Caption = "Grabar CD MP3s"
                btOKPachaCart.Visible = (PachaMode = 11000)
                LeEntra = CDR.CanSaveDataMode(Carrito.GetTotalMB)
                Label7.Caption = Label7.Caption + vbCrLf + "700 MB disponibles"
            End If
            'MM90 muestra de espacio disponible!
            If SP44(1) = "DVD" Then
                btOKPachaCart.Caption = "Grabar DVD"
                btOKPachaCart.Visible = (PachaMode = 11000)
                LeEntra = CDR.CanSaveDVDMode(Carrito.GetTotalMB)
                Label7.Caption = Label7.Caption + vbCrLf + "4482 MB disponibles"
            End If
            
        End If
        
        If SP44(0) = "BT" Then
            If SP44(1) = "DETECT" Then
                btOKPachaCart.Caption = "Comenzar búsqueda"
                btOKPachaCart.Visible = (PachaMode = 11000)
            Else
                btOKPachaCart.Caption = "Comprar en BLUETOOTH elegido"
                btOKPachaCart.Visible = (PachaMode = 11000)
                Label7.Caption = Label7.Caption + vbCrLf + "Asegúrese de tener espacio libre"
            End If
            LeEntra = True
        End If
    Else
        LeEntra = True
    End If
        
    If LeEntra Then
        Label7.ForeColor = vbGreen
    Else
        Label7.ForeColor = vbRed
    End If
    
End Sub

Private Function Promos() As String
    'ver las promociones ya grabadas
    Dim TMP As String
    
    Dim H As Long
    Dim SN As Single
    
    For H = 1 To Carrito.GetTotalPricesAudio
        If Carrito.GetPricesAudioBase(H) > 0 Then
            SN = Round(Carrito.GetPricesAudioBase(H) * PrecioBase / TemasPorCredito, 2)
            
            If H = 1 Then
                TMP = TMP + "1 fichero de AUDIO " + " por $ " + CStr(SN) + vbCrLf
            Else
                TMP = TMP + CStr(H) + " ficheros de AUDIO " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)" + vbCrLf
            End If
        End If
    Next H
    
    For H = 1 To Carrito.GetTotalPricesVideo
        If Carrito.GetPricesVideoBase(H) > 0 Then
            SN = Round(Carrito.GetPricesVideoBase(H) * PrecioBase / TemasPorCredito, 2)
            If H = 1 Then
                TMP = TMP + "1 fichero de VIDEO " + " por $ " + CStr(SN) + vbCrLf
            Else
                TMP = TMP + CStr(H) + " ficheros de VIDEO " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)" + vbCrLf
            End If
        End If
    Next H
    
    For H = 1 To Carrito.GetTotalPricesJAVA
        If Carrito.GetPricesJAVABase(H) > 0 Then
            SN = Round(Carrito.GetPricesJAVABase(H) * PrecioBase / TemasPorCredito, 2)
            If H = 1 Then
                TMP = TMP + "1 juego java " + " por $ " + CStr(SN) + vbCrLf
            Else
                TMP = TMP + CStr(H) + " juegos java " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)" + vbCrLf
            End If
        End If
    Next H
    
    For H = 1 To Carrito.GetTotalPricesRingtones
        If Carrito.GetPricesRingtonesBase(H) > 0 Then
            SN = Round(Carrito.GetPricesRingtonesBase(H) * PrecioBase / TemasPorCredito, 2)
            If H = 1 Then
                TMP = TMP + "1 ringtone " + " por $ " + CStr(SN) + vbCrLf
            Else
                TMP = TMP + CStr(H) + " ficheros de ringtones " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)" + vbCrLf
            End If
        End If
    Next H
    
    For H = 1 To Carrito.GetTotalPricesWallpapers
        If Carrito.GetPricesWallpapersBase(H) > 0 Then
            SN = Round(Carrito.GetPricesWallpapersBase(H) * PrecioBase / TemasPorCredito, 2)
            If H = 1 Then
                TMP = TMP + "1 wallpaper " + " por $ " + CStr(SN) + vbCrLf
            Else
                TMP = TMP + CStr(H) + " ficheros de wallpapers " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)" + vbCrLf
            End If
        End If
    Next H
    
    Promos = TMP
End Function


'-------Agregado por el complemento traductor------------
Private Sub Form_Load()
    
    'martino quiere laburar si logos ...
    'pero solo estos meses. apartir de 2009 se va a ver
    If ClaveAdmin = "martino" And Year(Date) < 2009 Then Image1.Visible = False
    
    tERR.Anotar "eaac", TengoBluetooth, PachaMode
    KeyPress = 0
    TecladoAnda = False
    TeclasApret = 0  'por las teclas que vienen de frmindex
    BusqBT = 0
    
    tUP.BorderStyle = 0
    tDown.BorderStyle = 0
    
    If TengoBluetooth Then
        tERR.Anotar "eaad", tBT.HWND
        BTM.UseEventMSG tBT.HWND
    End If
    
    If TengoCD Then
        CDR.SetStatus 0 'por las dudas
    End If
    
    If PachaMode = 11000 Then
        Label2.Caption = "Utilize los botones de desplazamiento para elegir DISPOSITIVOS. Confirme con el mismo boton de seleccion de discos y canciones"
    End If
    
    Pintar_fBoton Me
    Me.AutoRedraw = True
    
    'si esta en modo pacha las opciones del costado no entran en tabstop
    If PachaMode = 11000 Then
        btSalir.TabStop = False
        btANULA.TabStop = False
        btReview.TabStop = False
        
        Dim IMF As String
        IMF = ExtraData.getDef.getImagePath("tocuharribacomun")
        tUP.Picture = LoadPicture(IMF)
    
        IMF = ExtraData.getDef.getImagePath("touchabajocomun")
        tDown.Picture = LoadPicture(IMF)
    Else
        btSalir.TabStop = True
        btSalir.TabIndex = 1 'el 0 (primero) es siempre el primer dispositivo
        btANULA.TabStop = True
        btANULA.TabIndex = 2
        btReview.TabStop = True
        btReview.TabIndex = 3
    End If
    
    tERR.Anotar "eaae", tBT.HWND
    
    CD1(0).Top = Label3.Top + Label3.Height + 60
    teX1(0).Top = CD1(0).Top + CD1(0).Height
    'CD1(0).Left = Line1.X1 - CD1(0).Width
    teX1(0).Left = CD1(0).Left
    tERR.Anotar "daan", Carrito.GetFileCant, Carrito.GetFileCantFull
    
    If TengoUSB Then UB.UseEventMSG tNADA.HWND
    
    tERR.Anotar "eaaf", tNADA.HWND
    TecladoAnda = True
    
    'CUANDO HAY ALGUN LECTOR DE MEMORIA YA SE CARGA COMO USB
    'ENTONCES APARECE COMO DISPOSITIVO DE CERO MB Y AL CONECTARLE ALGO
    'NO LANZA EVENTO YA QUE EL DISPOSITIVO YA EXISTIA, SOLO CAMBIA SU TAMAÑO EN MB
    Timer1.Interval = 1000
    Me.KeyPreview = True
    
    tERR.Anotar "eaag"
'    Dim RDS As TypeLic
'    RDS = K.sabseee("mLicencia3PMVtaMusica")
'    If RDS < DMinima Then
'        btSalir.Enabled = False
'        btBUY.Enabled = False'
'        btANULA.Enabled = False
'    End If
    
End Sub

Private Sub UpdateDrives()

    'si tiene activado el bluetooth entonces hay un boton fijo para buscar por bluetooth
    'la cantidad y el orden del los botones es el siguiente
    '1- dispositivos bluetooth y un boton para detectarlos (solo si esta configurado asi)
    '2- dispositivos usb o un boton que pida que inserte
    
    tERR.Anotar "xsaa"
    UnloadBtBuy
    
    Dim UltTitUsado As Long 'un titulo por cada tipo de dispositivo
    
    tERR.Anotar "xsaf", TengoBluetooth
    If TengoBluetooth Then
        Label9(0).Caption = "BLUETOOTH"
        tERR.Anotar "xsag", BTM.Count
        'si o si el boton de detectar
        If BusqBT = 0 Then
            btBUY(0).Caption = "Buscar dispositivos bluetooth"
        Else
            btBUY(0).Caption = "Buscar nuevamente dispositivos bluetooth"
        End If
        
        btBUY(0).Top = Label9(0).Top + Label9(0).Height
        btBUY(0).Tag = "BT DETECT"
        
        If BTM.Count > 0 Then
            H = btBUY.Count
            tERR.Anotar "xsah", H
            
            Dim CBT As tbrBtActivex.TbrBtDevice
            tERR.Anotar "xsah2"
            
            For Each CBT In BTM
                tERR.Anotar "xsah3", H
                
                Load btBUY(H)
                btBUY(H).Top = btBUY(H - 1).Top + btBUY(H - 1).Height + 60
                btBUY(H).Left = btBUY(H - 1).Left
                tERR.Anotar "xsah4", btBUY(H).Top
                
                tERR.Anotar "xsai", CBT.Name, CBT.getAddress
                btBUY(H).Caption = "Comprar en Bluetooth: " + CBT.Name + vbCrLf + " (" + CBT.getAddress + ")"
                    
                btBUY(H).Tag = "BT " + CBT.getAddress  'PARA PODER USARLO
                
                btBUY(H).Visible = True
                btBUY(H).TabIndex = btBUY(H - 1).TabIndex + 1
                'que se cargue despintado!!
                SelBT btBUY(H), False
                
                H = H + 1
            Next
            tERR.Anotar "xsaj"
        End If
        
        Load Label9(1)
        Label9(1).Caption = "USB"
        Label9(1).Top = btBUY(btBUY.Count - 1).Top + btBUY(btBUY.Count - 1).Height + 220
        Label9(1).Visible = True
        UltTitUsado = 1
        H = btBUY.Count
    Else
        Label9(0).Caption = "USB"
        UltTitUsado = 0
        H = 0 'no hay otros medios por el momento
    End If
    
    tERR.Anotar "xsabA", TengoUSB, TengoBluetooth, TengoCD
    
    If TengoUSB Then
        'si o si agrega el titulo de usb
        tERR.Anotar "xsab", UB.GetCantidadUSB
        If UB.GetCantidadUSB = 0 Then
            If H > 0 Then Load btBUY(H) 'es cero cuando no esta activado el bluetooth
            btBUY(H).Caption = "Inserte dispositivo USB" + vbCrLf + "Se detectan instantáneamente"
            btBUY(H).Tag = "USB DETECT" 'se ignora no hay busqueda es automático
            btBUY(H).Top = Label9(UltTitUsado).Top + Label9(UltTitUsado).Height  'btBUY(H - 1).Top + btBUY(H - 1).Height + 60
            If H > 0 Then btBUY(H).Left = btBUY(H - 1).Left
            btBUY(H).Visible = True
            If H > 0 Then
                btBUY(H).TabIndex = btBUY(H - 1).TabIndex + 1
                'que se cargue despintado!!
                SelBT btBUY(H), False
            Else
                btBUY(0).TabIndex = 0
                'no hay nada mas por eso el setfocus
                'ya se pinta alli
                btBuy_GotFocus 0
                'si lanzo el evento setfocus no funciona!
            End If
            
            H = H + 1
            
        Else
            'SI HAY MAS DE UNO QUE SEPA QUIEN ES QUIEN
            tERR.Anotar "xsad", UB.GetCantidadUSB
            
            Dim H2 As Long
            For H2 = H To UB.GetCantidadUSB + H - 1
                If H2 > 0 Then Load btBUY(H2)
                'xxxxxxxxxxxxxxxxxxxx error 68 disp no disponible
                UB.RefreshValues H2 - H + 1
                If H2 > H Then
                    btBUY(H2).Top = btBUY(H2 - 1).Top + btBUY(H2 - 1).Height + 60
                Else
                    btBUY(H2).Top = Label9(UltTitUsado).Top + Label9(UltTitUsado).Height
                End If
                
                If H2 > 0 Then btBUY(H2).Left = btBUY(H2 - 1).Left
                tERR.Anotar "xsae", UB.GetNameUSB(H2 - H + 1)
                
                btBUY(H2).Caption = "Comprar por USB: " + UB.GetNameUSB(H2 - H + 1) + _
                    " (" + UB.GetLetterUSB(H2 - H + 1) + ":\)" + vbCrLf + _
                    "Tiene " + CStr(UB.GetFreeMB(H2 - H + 1)) + " MB libres"
                
                btBUY(H2).Tag = "USB " + CStr(H2 - H + 1) 'el segundo es el indice en "UB"
                
                btBUY(H2).Visible = True
                If H2 > 0 Then
                    btBUY(H2).TabIndex = btBUY(H2 - 1).TabIndex + 1
                    'que se cargue despintado!!
                    SelBT btBUY(H2), False
                Else
                    'no hay nada mas por eso el setfocus
                    btBUY(0).TabIndex = 0
                    'no hay nada mas por eso el setfocus
                    'ya se pinta alli
                    btBuy_GotFocus 0
                    'si lanzo el evento setfocus no funciona!
                End If
                
                
            Next H2
            
            H = H + UB.GetCantidadUSB
        End If
        
        UltTitUsado = UltTitUsado + 1
    End If
    
    If TengoCD Then
        
        tERR.Anotar "xsabD", CDR.GetStatus
        '***********************************
        'necesitaria confirmacion de que el carrito tiene solo archivo de musica!!! XXXX
        '***********************************
        If UltTitUsado > 0 Then Load Label9(UltTitUsado)
        Label9(UltTitUsado).Caption = "CD GRABABLE"
        Label9(UltTitUsado).Top = btBUY(btBUY.Count - 1).Top + btBUY(btBUY.Count - 1).Height + 220
        Label9(UltTitUsado).Visible = True
        
        If H > 0 Then
            Load btBUY(H) 'es cero cuando no esta activado el bluetooth
            btBUY(H).Left = btBUY(H - 1).Left
        End If
        
        btBUY(H).Caption = "Grabar CD de audio" + vbCrLf + "Inserte CD vacio antes"
        btBUY(H).Tag = "CD AUDIO" 'se ignora no hay busqueda es automático
        btBUY(H).Top = Label9(UltTitUsado).Top + Label9(UltTitUsado).Height  'btBUY(H - 1).Top + btBUY(H - 1).Height + 60
        btBUY(H).Visible = True
        
        'que se cargue despintado!!
        SelBT btBUY(H), False
        If H > 0 Then
            btBUY(H).TabIndex = btBUY(H - 1).TabIndex + 1
        Else
            btBUY(0).TabIndex = 0
            'no hay nada mas por eso el setfocus
            'ya se pinta alli
            btBuy_GotFocus 0
            'si lanzo el evento setfocus no funciona!
        End If
        
        H = H + 1
        Load btBUY(H)
        btBUY(H).Caption = "Grabar CD de MP3" + vbCrLf + "Inserte CD vacio antes"
        btBUY(H).Tag = "CD MP3DATA" 'se ignora no hay busqueda es automático
        btBUY(H).Top = btBUY(H - 1).Top + btBUY(H - 1).Height + 60
        btBUY(H).Left = btBUY(H - 1).Left
        SelBT btBUY(H), False
        btBUY(H).Visible = True
        btBUY(H).TabIndex = btBUY(H - 1).TabIndex + 1
        
        'mm90
        'agregar la grabacion de DVD
        H = H + 1
        Load btBUY(H)
        btBUY(H).Caption = "Grabar DVD" + vbCrLf + "Inserte DVD vacio antes"
        btBUY(H).Tag = "CD DVD" 'se ignora no hay busqueda es automático
        btBUY(H).Top = btBUY(H - 1).Top + btBUY(H - 1).Height + 60
        btBUY(H).Left = btBUY(H - 1).Left
        SelBT btBUY(H), False
        btBUY(H).Visible = True
        btBUY(H).TabIndex = btBUY(H - 1).TabIndex + 1
        
    End If
    
    tERR.Anotar "xsak"
    
    AcomodarIndicadores
End Sub

Private Sub UnloadBtBuy()
    Dim H As Long
    For H = 1 To btBUY.Count - 1
        Unload btBUY(H)
    Next H
    
    For H = 1 To Label9.Count - 1
        Unload Label9(H)
    Next H
End Sub

Private Function LoadLista()
    
    'ver si se puede mostrar todo. Si quiedara muy chiquito ponemos algun mensaje
    Dim MinH As Long 'minimo de alto que muestro
    
    tERR.Anotar "daaq", Carrito.GetFileCant
    
    If Carrito.GetFileCant > 0 Then
        Dim H As Long
        For H = 1 To Carrito.GetFileCant
            ShowElem2 H, Carrito.GetFileCant
        Next H
    Else
        Label8.Visible = True 'tengo lugar para los precios
    End If
    
    If teX1(teX1.UBound).Top + teX1(teX1.UBound).Height + 60 > Label8.Top Then
        Label8.Visible = False
    Else
        Label8.Visible = True 'tengo lugar para los precios
    End If
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
        If K.sabseee("3pm") = Supsabseee Then
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
        CD1(I).Top = Label3.Top + Label3.Height + 160
        teX1(I).Top = CD1(I).Top + CD1(I).Height
        
        CD1(I).Left = Label3.Left + 120
        teX1(I).Left = CD1(I).Left
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
    tERR.AppendLog tERR.ErrToTXT(Err), "cpCC4"
    Resume Next
End Function

Private Function ShowElem2(I As Long, TotShow As Long) 'este es mas chico y de arriba hacia abajo
    On Local Error GoTo MER
    
    'segun la cantidad de elementos a mostrar se muestras mas grandes o mas chicos
    Dim TotH As Long
    TotH = Me.Height - (Label3.Top + Label3.Height) - tDown.Height
    
    If TotShow <= 8 Then
        CD1(0).Height = 1000
        CD1(0).Width = 1200
        CD1(0).Left = AnchoCol - CD1(0).Width - 15
        teX1(0).Font.Size = 12
        teX1(0).Font.Bold = True
    End If
    
    If TotShow >= 9 And TotShow <= 19 Then
        CD1(0).Height = TotH / (TotShow + 1)
        CD1(0).Width = CD1(0).Height * 1.2
        CD1(0).Left = AnchoCol - CD1(0).Width - 15
        teX1(0).Font.Size = 10
        teX1(0).Font.Bold = True
    End If
    
    If TotShow >= 20 Then
        CD1(0).Height = TotH / (TotShow + 5)
        CD1(0).Width = CD1(0).Height * 1.2
        teX1(0).Font.Size = 8
        teX1(0).Font.Bold = True
    End If
    
    CD1(0).Left = AnchoCol - CD1(0).Width - 15
    teX1(0).Height = CD1(0).Height
    teX1(0).Width = AnchoCol - CD1(0).Width - 90
    
    Load CD1(I)
    Load teX1(I)
    Load LN(I)
    
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
        If K.sabseee("3pm") = Supsabseee Then
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
        CD1(I).Top = CD1(0).Top
        teX1(I).Top = teX1(0).Top
        
        CD1(I).Left = CD1(0).Left
        teX1(I).Left = teX1(0).Left
    Else
        CD1(I).Top = CD1(I - 1).Top + CD1(I - 1).Height + 15
        teX1(I).Top = teX1(I - 1).Top + teX1(I - 1).Height + 15
        
        CD1(I).Left = CD1(I - 1).Left
        teX1(I).Left = teX1(I - 1).Left
        
    End If
    LN(I).Y1 = teX1(I).Top + teX1(I).Height
    LN(I).Y2 = LN(I).Y1
    LN(I).X1 = 0
    LN(I).X2 = Line1.X1
    
    LN(I).Visible = True
    CD1(I).Visible = True
    teX1(I).Visible = True
    
    tERR.Anotar "daav"
    
    Exit Function
MER:
    tERR.AppendLog tERR.ErrToTXT(Err), "cpCC4"
    Resume Next
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UnloadBtBuy
    'aviso que no hay mas dispositivos
    
    If TengoBluetooth Then
        'revisar si jode xxxxxxxxxxxxx
        BTM.ReiniciarColeccion
    End If
    
    Timer1.Interval = 0
End Sub

Private Sub Form_Resize()
    tERR.Anotar "eaah"
    'Me.PaintPicture frmIndex.picFondoDisco.Image, 0, 0, Me.Width, Me.Height
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
        
    Image1.Left = Me.Width - Image1.Width - 60
    Image1.Top = Me.Height - Image1.Height - 30 '30
    
    Label2.Top = 30
    Label2.Left = 30
    Label2.Width = Me.Width - Line1.X1
    
    AnchoCol = Me.Width / 3
    tERR.Anotar "eaai", AnchoCol
    
    Line1.Y1 = Label2.Top + Label2.Height + 130
    Line1.Y2 = Me.Height
    Line2.Y1 = Line1.Y1
    Line2.Y2 = Line1.Y2
    
    Line1.X1 = AnchoCol
    Line2.X1 = AnchoCol * 2
    Line1.X2 = AnchoCol
    Line2.X2 = AnchoCol * 2
    
    '- cont compra -
    Label3.Top = Label2.Top + Label2.Height + 30
    Label3.Left = (AnchoCol / 2 - Label9(0).Width / 2)
    
    CD1(0).Top = Label3.Top + Label3.Height + 30
    teX1(0).Top = CD1(0).Top
    CD1(0).Left = AnchoCol - CD1(0).Width - 30
    teX1(0).Left = 30
    teX1(0).Width = AnchoCol - CD1(0).Width - 90
    
    tERR.Anotar "eaaj", PachaMode
    '- dispos -
    Label9(0).Top = Label3.Top
    Label9(0).Left = AnchoCol + (AnchoCol / 2 - Label9(0).Width / 2)
    
    btBUY(0).Top = Label9(0).Top + Label9(0).Height + 90
    btBUY(0).Left = AnchoCol + (AnchoCol / 2 - btBUY(0).Width / 2)
    
    If PachaMode = 11000 Then
        'que quede igual!
        tDown.Top = frmIndex.picFondoPacha.Top + frmIndex.t1.Top
        tDown.Left = frmIndex.picFondoPacha.Left + frmIndex.t1.Left
        
        tUP.Top = frmIndex.picFondoPacha.Top + frmIndex.t3.Top
        tUP.Left = frmIndex.picFondoPacha.Left + frmIndex.t3.Left
        'este boton es más grande!
        'btOKPachaCart.Top = frmIndex.picFondoPacha.Top + frmIndex.btOKPacha.Top
        btOKPachaCart.Top = Me.Height - btOKPachaCart.Height + 60
        btOKPachaCart.Left = frmIndex.picFondoPacha.Left + frmIndex.btOKPacha.Left
        btOKPachaCart.Width = frmIndex.btOKPacha.Width
        'aqui tengo mas lugar y necesito más texto
        'btOKPachaCart.Height = frmIndex.btOKPacha.Height
        btOKPachaCart.Caption = "COMPRAR"
        
        tDown.Visible = True
        tUP.Visible = True
        btOKPachaCart.Visible = True
    End If
    
    'acomodar indicadores
    tERR.Anotar "eaak"
    UpdateData False
    
    tERR.Anotar "eaal"
    UpdateDrives
    
    tERR.Anotar "eaam"
    AcomodarIndicadores
    tERR.Anotar "eaan"
    LoadLista
    tERR.Anotar "eaao"
End Sub

Private Sub AcomodarIndicadores()
    
    tERR.Anotar "xsal", btBUY.Count - 1
    
    'el alto de los botones es 705
    'los botones aqui tienen ese alto tambien
    Dim BT1_Top As Long
    BT1_Top = frmIndex.frDiscos.Top + _
              frmIndex.picFondoDisco.Top + _
              frmIndex.picFondoDisco.Height - _
              ((3 * btANULA.Height) + _
              (2 * SeparacionTocuhDerecho))
    'no lo vinculo a los otros botones de frmindex por que no necesariamente están al costado
    'cuando no es modo pacha
    Label10.Top = BT1_Top - Label10.Height - 30
    btSalir.Top = BT1_Top
    btANULA.Top = btSalir.Top + btSalir.Height + SeparacionTocuhDerecho
    btReview.Top = btANULA.Top + btANULA.Height + SeparacionTocuhDerecho
    
    btSalir.Left = Me.Width - btSalir.Width + 90
    btANULA.Left = btSalir.Left
    btReview.Left = btSalir.Left
    Label10.Left = btSalir.Left
    
    Label1.Top = 30 'Image1.Top + Image1.Height + 30
    Label1.Width = 3900  'el ancho de una columna es 4000
    
    Dim TotIndic As Long 'total de indicadores
    TotIndic = 9
    'arrimar
    Label1.Left = Me.Width - Label1.Width - 45
    Label4.Left = Label1.Left
    Label5.Left = Label1.Left
    Label6.Left = Label1.Left
    Label7.Left = Label1.Left
    Label8.Left = 30
    
    tERR.Anotar "xsam"
    
    Label4.Width = Label1.Width
    Label5.Width = Label1.Width
    Label6.Width = Label1.Width
    Label7.Width = Label1.Width
    Label8.Width = Line1.X1 - 60
    
    Label4.Height = Label1.Height
    Label5.Height = Label1.Height
    Label6.Height = Label1.Height
    Label7.Height = Label1.Height
            
    Label4.Top = Label1.Top + Label1.Height - 15
    Label5.Top = Label4.Top + Label1.Height + 30
    
    Label6.Top = Label5.Top + Label1.Height - 15
    
    Label7.AutoSize = True
    Label7.Top = Label6.Top + Label6.Height + 30

    'lista de promos
    Label8.Caption = "PRECIOS" + vbCrLf + Promos
    Label8.Top = Me.Height - Label8.Height - 60
    
    tERR.Anotar "xsan"
End Sub

Private Sub tBT_Change()
    
    If tBT.tExt = "" Then Exit Sub
    
    tERR.Anotar "BT=" + tBT.tExt
    If ActivarERR Then tERR.AppendSinHist "BbbTtt:::" + tBT.tExt
    
    'SE QUEDO SIN LUGAR EL DISPOSITIVO
    'IMAGINO QUE PUEDE REPRESENTAR OTRAS COSAS TAMBIEN
    '33722,67:BT=4|Fallo Al comprobar el servicio
    '33722,67:BT=4|Fallo General
    
    Dim SP() As String
    SP = Split(tBT.tExt, "|")
    
    Select Case SP(0)
        Case "0"
            
        Case "1" 'sale drive
            'termino de buscar dispositivos
            tERR.Anotar "BTM_IF"
            
            UpdateDrives
            SW.ShowWait ""
            
        Case "2"
            'connection service status
            tERR.Anotar "BTM_CSR:", SP(1)
            'con algunos valores aqui se clava
            'por ejemplo
            'BTM_CSR:.Outgoing Connection Disconnect Indication!
            'me saca de pecho
            
        Case "3"
            'llego ok un archivo
            tERR.Anotar "BTM_SND_OK"
        Case "4"
            'llego mal el archivo
            tERR.Anotar "BTM_SND_BAD"
            tERR.AppendLog "FEBT-bt" 'falla en envio bluetooth
            'SE RECLAVA SI DESCONECTO BLUETOOTH DE PECHO
'            BT=4|General failed
'            BTM_SND_BAD
        Case "5"
            'encontro un dispositivo
            tERR.Anotar "BTM_DEV", SP(1), SP(2)
    End Select
    
    tBT.tExt = ""
End Sub

Private Sub Timer1_Timer()
    If TengoUSB Then UB.RefreshDriveList
    
    'negrada!!
    If TengoCD Then
        If CDR.GetStatus > 0 And CDR.GetStatus < 100 Then
            
            TecladoAnda = False
            'mm90 cambiar palabra cd por disco!
            SW.ShowWait "Grabando disco " + vbCrLf + _
                CDR.GetLastMsg + vbCrLf + _
                CStr(CDR.GetStatus) + " %", , CSng(CDR.GetStatus)
            
        End If
        
        If CDR.GetStatus = 100 Then
            SW.ShowWait "Terminando grabación ..."
            'descontale!
            tERR.Anotar "daby6-CD"
            VarCreditos -Carrito.CalculateTotalPrice
            'sumo al contador de creditos de carrito lo que se gasto
            SumarContadorCarrito Carrito.CalculateTotalPrice
            'indicar cuanta plata entro en esta fonola en concepto de compra de música
            Dim YU As Long, DTaa As String
            DTaa = CStr(Year(Date)) + STRceros(Month(Date), 2) + STRceros(Day(Date), 2) + STRceros(Hour(time), 2) + STRceros(Minute(time), 2)
            
            'grabar un registro de todo lo que se compro para control.
            Dim PrecioCU As Single 'precio de cada cancion
            PrecioCU = (Carrito.CalculateTotalPrice * (PrecioBase / TemasPorCredito))
            PrecioCU = Round(PrecioCU / Carrito.GetFileCantFull, 2)
            For YU = 1 To Carrito.GetFileCantFull
                'tERR.Anotar "A198|B" + Carrito.GetElementFull(YU)
                'grabar en un registro de aca
                'el de cd solo usa mp3s, si habia mas cosas en el carrito las rgistrara erroneamente
                'XXXX
                dwqu "C" + Carrito.GetElementFull(YU) + "*" + CStr(PrecioCU), dwQU_See, DTaa
            Next
            
            Carrito.ClearCart
            SW.ShowWait "Grabacion OK!", 3500
            CDR.SetStatus 0
            TecladoAnda = True
            Unload Me
        End If
        
        If CDR.GetStatus > 100 Then
            SW.ShowWait "Grabacion con falla: " + CStr(CDR.GetStatus) + vbCrLf + CDR.GetLastMsg, 3500
            CDR.SetStatus 0
            TecladoAnda = True
            Unload Me
        End If
    End If
End Sub

Private Sub tNADA_Change()
    If tNADA.tExt = "" Then Exit Sub
    tERR.Anotar "CarUSB", tNADA.tExt
    Dim SP() As String
    SP = Split(tNADA.tExt, "|")
    
    Select Case SP(0)
        Case "0" 'entro drive
            UpdateData False
            UpdateDrives
            
        Case "1" 'sale drive
            UpdateData False
            UpdateDrives
    End Select
    
    tNADA.tExt = ""
End Sub
