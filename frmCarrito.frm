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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox tBT 
      Height          =   435
      Left            =   4890
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Left            =   8160
      Top             =   6780
   End
   Begin VB.TextBox tNADA 
      Height          =   435
      Left            =   2040
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1155
   End
   Begin tbr3pm.tbrFullProc SW 
      Height          =   435
      Left            =   6990
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   767
   End
   Begin tbrFaroButton.fBoton btANULA 
      Height          =   705
      Left            =   180
      TabIndex        =   2
      Top             =   2430
      Width           =   3645
      _ExtentX        =   6429
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
      Height          =   885
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   1530
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   1561
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "NO HAY DISPOSITIVOS CONECTADOS"
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton btSalir 
      Height          =   705
      Left            =   180
      TabIndex        =   0
      Top             =   810
      Width           =   3645
      _ExtentX        =   6429
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
      Left            =   180
      TabIndex        =   3
      Top             =   3150
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   1244
      fFColor         =   16777215
      fBColor         =   16777215
      fCapt           =   "Eliminar parte de la compra"
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "MB libres en dispositivo"
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
      Height          =   300
      Left            =   180
      TabIndex        =   13
      Top             =   6870
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Canciones totales:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   180
      TabIndex        =   11
      Top             =   5460
      Width           =   4095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "MB libres en dispositivo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   180
      TabIndex        =   10
      Top             =   6570
      Width           =   4095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Costo carrito"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   180
      TabIndex        =   9
      Top             =   6300
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Credito:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   180
      TabIndex        =   8
      Top             =   6030
      Width           =   4095
   End
   Begin VB.Label teX1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Canciones elegidas para comprar: 99. Costo total $350.000. Credito disponible $ 380.000"
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
      Left            =   4290
      TabIndex        =   7
      Top             =   3390
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Image CD1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1425
      Index           =   0
      Left            =   4290
      Stretch         =   -1  'True
      Top             =   1950
      Visible         =   0   'False
      Width           =   2445
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
      Height          =   435
      Left            =   4140
      TabIndex        =   6
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciones:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   180
      TabIndex        =   5
      Top             =   5760
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Utilize los botones de desplazamiento para elegir las opciones. Confirme con el mismo boton de seleccion de discos y canciones"
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
      Height          =   645
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   9600
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   9720
      Picture         =   "frmCarrito.frx":0000
      Top             =   60
      Width           =   2205
   End
End
Attribute VB_Name = "frmCarrito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
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
    
    If TengoBluetooth And Index = 0 Then
    
        'quiere buscar bluetooth
        tERR.Anotar "BT_INQ_222"
        BTM.inquiereDev
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
                MT = Int(Rnd * 3) + 1
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
    
    tERR.Anotar "daab", Carrito.CalculateTotalPrice, CREDITOS
    
    'ver que alcance la plata que puso
    If Carrito.CalculateTotalPrice > CREDITOS Then
        SW.ShowWait "El crédito no es suficiente para la compra elegida", 3500
        SW.ShowWait ""
        Unload Me
        Exit Sub
    End If
    
    If TengoBluetooth Then
        tERR.Anotar "daab2", BTM.Count, UB.GetCantidadUSB
        'ver si hay dispositivos
        If (Index = 0) Or (UB.GetCantidadUSB = 0 And BTM.Count = 0) Then
            tERR.Anotar "daac2"
            SW.ShowWait "No hay dispositivos conectados!", 2500
            Exit Sub
        End If
    Else
        tERR.Anotar "daab2", UB.GetCantidadUSB
        'ver si hay dispositivos
        If (Index = 0) Or (UB.GetCantidadUSB = 0) Then
            tERR.Anotar "daac3"
            SW.ShowWait "No hay dispositivos conectados!", 2500
            Exit Sub
        End If
    End If
          
    UB.DevSel = Index 'hay solo uno, lo elijo

    'VER SI ALCANZA EL ESPACIO LIBRE
    tERR.Anotar "daae", btBuy(Index).Tag
    Dim JP() As String
    JP = Split(btBuy(Index).Tag)
    
    tERR.Anotar "daac2"
    
    '***************************************************************************************
    If JP(0) = "USB" Then
    
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
    '***************************************************************************************
    'medir la velocidad
    Dim Copiado As Single 'cantidad copiada
    Dim sTimeCopyINI As Single 'tiempo en que la copio
    Dim sTimeCopy As Single 'tiempo en que la copio
            
    Dim MBxSec As Single
    Dim Falta As Single 'segundos que faltan
    
    Dim totCart As Long
    
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
            RDS = K.LICENCIA("mLicencia3PMVtaMusica")
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
        
        Dim H As Long
        For H = 1 To Carrito.GetFileCantFull
            InFolder = fso.GetBaseName(fso.GetParentFolderName(Carrito.GetElementFull(H)))
            tERR.Anotar "dabx2", InFolder
            'EN LOS CELULARES O PENDRIVES PUEDE APARECER EL ERROR
            '-2147024784
            
            If InFolder = "" Then GoTo SIG444
            
            'poner en cero la espera
            tERR.Anotar "daby2-BT", Carrito.GetElementFull(H)
            
            BTM.SendFileBT Carrito.GetElementFull(H), JP(1)
            
            Dim KK As Single
            Dim SecPas As Long, lastSP As Long
            KK = Timer
            lastSP = 99
            Do
                DoEvents 'SIN ESTO NO ANDA el cancelar!!!!
                SecPas = CLng(CSng(Timer - KK))
                If lastSP <> SecPas Then
                    If H <= 1 Then
                        SW.ShowWait "Enviando por Bluetooth " + vbCrLf + _
                            "(recuerde ACEPTAR el envio en su celular)" + vbCrLf + _
                            fso.GetBaseName(Carrito.GetElementFull(H)), , (SecPas Mod 100)
                    Else
                        SW.ShowWait "Enviando por Bluetooth " + vbCrLf + _
                            "(recuerde ACEPTAR el envio en su celular)" + vbCrLf + _
                            fso.GetBaseName(Carrito.GetElementFull(H)) + vbCrLf + _
                            "(" + CStr(Round(MBxSec, 3)) + _
                            " MB/S falta aproximado: " + FaltaTXT(Falta - SecPas) + ")", , (SecPas Mod 100)
                    End If
                        
                    lastSP = SecPas
                End If
                
                If BTM.PushStatus = 2 Then
                    tERR.Anotar "dabz", SecPas, Round(MBxSec, 2)
                    BTM.PushStatus = 0 'lo dejo en cero
                    Exit Do
                End If
                'estar atento a si cancela el usuario
                If BTM.PushStatus = 3 Then
                    tERR.Anotar "dadb", SecPas, Round(MBxSec, 2)
                    SW.ShowWait "Usuario no acepto o fallo la conexion", 3
                    Exit Do
                End If
                ''NO SE PUEDE CANCELAR!
'                If BTM.PushStatus = 4 Then
'                    tERR.Anotar "dadc", SecPas, Round(MBxSec, 2)
'                    SW.ShowWait "Cancelado por el usuario", 3
'                    Exit Do
'                End If
                
            Loop
            SW.ShowWait ""
                        
            Copiado = Copiado + (FileLen(Carrito.GetElementFull(H)) / 1048576)
            sTimeCopy = Timer - sTimeCopyINI
            
            MBxSec = Round(Copiado / sTimeCopy, 6)
            Falta = CLng(CSng((totCart - Copiado) / MBxSec))
            
            'descontar el credito correspondiente a los que grabo
            'XXXXXXXXXXXXXX
            'No es un numero entero.... quilombo parecido al de las canciones
            'para nmo hacer lio saco todo lo que hay que sacar si se copio el primero ok
            
            'no lo saco al final por que si no se van a avivar y sacar el pendrive antes de
            'terminar y les va a costar cero!
            If H = 1 Then
                VarCreditos -Carrito.CalculateTotalPrice
                'sumo al contador de creditos de carrito lo que se gasto
                SumarContadorCreditos Carrito.CalculateTotalPrice
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
        SW.ShowWait "Proceso terminado con exito", 3300
        
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

Private Sub btBUY_Click(Index As Integer)
    ComprarCC CLng(Index)
End Sub

Private Sub btBuy_GotFocus(Index As Integer)
    SelBT btBuy(Index), True
End Sub

Private Sub btBuy_LostFocus(Index As Integer)
    SelBT btBuy(Index), False
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
    
    tERR.Anotar "daam", Chr(KeyCode), KeyCode
    
    Select Case KeyCode
        Case TeclaDER: SendKeys "{TAB}"
        Case TeclaIZQ: SendKeys "+{TAB}"
        Case TeclaPagAd: SendKeys "{TAB}"
        Case TeclaPagAt: SendKeys "+{TAB}"
        Case TeclaESC
            'ver si esta cancelando un bluetooth
            If BTM.PushStatus > 0 Then
                'NO SE PUEDE CANCELAR!
                'BTM.PushStatus = 4
            Else
                Unload Me
            End If
        Case TeclaCarrito: SendKeys "{ENTER}"
        Case TeclaCerrarSistema
            tERR.Anotar "YCS_FrmCart"
            YaCerrar3PM
        
        Case TeclaShowContador 'para uso mio!!!
            'elegir el que este elegido
            Dim H As Long
            For H = 1 To btBuy.Count - 1
                If btBuy(H).BackColor = ColSel Then  'esta elegido
                    UB.DevSel = H
                    Unload Me
                    frmCarritoDelete.Show 1
                End If
            Next H
            
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

Private Sub UpdateData(SoloCredit As Boolean)

    tERR.Anotar "daao", ShowCreditsMode, CREDITOS
    Select Case ShowCreditsMode
        Case 1 'modo creditos
            Label5.Caption = "Costo total: " + CStr(Carrito.CalculateTotalPrice)
            Label6.Caption = "Credito disponible: " + CStr(CREDITOS)
            
            If CREDITOS >= Carrito.CalculateTotalPrice Then
                Label6.ForeColor = vbGreen
            Else
                Label6.ForeColor = vbRed
            End If
            
        Case 0 'modo plata
            Label5.Caption = "Costo total: $ " + CStr(Carrito.CalculateTotalPrice * PrecioBase / TemasPorCredito)
            Label6.Caption = "Credito disponible: $ " + CStr(Round(CREDITOS * PrecioBase / TemasPorCredito, 2))
            
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
    If UB.GetCantidadUSB > 0 Then
        Dim H As Long
        For H = 1 To UB.GetCantidadUSB
            Label7.Caption = Label7.Caption + vbCrLf + "Espacio libre " + CStr(H) + ": " + CStr(UB.GetFreeMB(H)) + " MB"
            If (UB.GetFreeMB(H)) >= (Carrito.GetTotalMB) Then LeEntra = True
        Next H
    Else
        Label7.Caption = Label7.Caption + vbCrLf + "NO HAY DISPOSITIVOS"
        LeEntra = False
    End If
    If LeEntra Then
        Label7.ForeColor = vbGreen
    Else
        Label7.ForeColor = vbRed
    End If
    'lista de promos
    Label8.Caption = "PRECIOS" + vbCrLf + Promos
    
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
    
    Promos = TMP
End Function


'-------Agregado por el complemento traductor------------
Private Sub Form_Load()
    
    If TengoBluetooth Then
        BTM.UseEventMSG tBT.HWND
    End If
    
    Pintar_fBoton Me
    Me.AutoRedraw = True
    
    CD1(0).Top = Label3.Top + Label3.Height + 60
    teX1(0).Top = CD1(0).Top + CD1(0).Height
    'CD1(0).Left = Line1.X1 - CD1(0).Width
    teX1(0).Left = CD1(0).Left
    tERR.Anotar "daan", Carrito.GetFileCant, Carrito.GetFileCantFull
    
    UB.UseEventMSG tNADA.HWND
    
    UpdateData False
    
    UpdateDrives
    
    'CUANDO HAY ALGUN LECTOR DE MEMORIA YA SE CARGA COMO USB
    'ENTONCES APARECE COMO DISPOSITIVO DE CERO MB Y AL CONECTARLE ALGO
    'NO LANZA EVENTO YA QUE EL DISPOSITIVO YA EXISTIA, SOLO CAMBIA SU TAMAÑO EN MB
    Timer1.Interval = 1000
    
'    Dim RDS As TypeLic
'    RDS = K.LICENCIA("mLicencia3PMVtaMusica")
'    If RDS < DMinima Then
'        btSalir.Enabled = False
'        btBUY.Enabled = False'
'        btANULA.Enabled = False
'    End If
    
End Sub

Private Sub UpdateDrives()
    tERR.Anotar "xsaa"
    UnloadBtBuy
    
    tERR.Anotar "xsab", UB.GetCantidadUSB
    If UB.GetCantidadUSB > 0 Then
        btBuy(0).Visible = False
        
        'SI SOLO HAY UN DISPOSITIVO ALCANZA CON QUE DIGA COMPRAR
        If UB.GetCantidadUSB = 1 Then
            tERR.Anotar "xsac", UB.GetFreeMB(1)
            UB.RefreshValues 1
            Load btBuy(1)
            btBuy(1).Top = btBuy(0).Top
            btBuy(1).Caption = "Comprar ahora" + vbCrLf + CStr(UB.GetFreeMB(1)) + " MB libres en dispositivo"
            btBuy(1).Visible = True
            btBuy(1).Tag = "USB"
            btBuy(1).TabIndex = btBuy(0).TabIndex + 1
            'que se cargue despintado!!
            SelBT btBuy(1), False
        End If
        
        'SI HAY MAS DE UNO QUE SEPA QUIEN ES QUIEN
        tERR.Anotar "xsad", UB.GetCantidadUSB
        If UB.GetCantidadUSB > 1 Then
            Dim H As Long
            For H = 1 To UB.GetCantidadUSB
                Load btBuy(H)
                'xxxxxxxxxxxxxxxxxxxx error 68 disp no disponible
                UB.RefreshValues H
                If H = 1 Then
                    btBuy(H).Top = btBuy(H - 1).Top
                Else
                    btBuy(H).Top = btBuy(H - 1).Top + btBuy(H - 1).Height
                End If
                tERR.Anotar "xsae", UB.GetNameUSB(H)
                
                btBuy(H).Caption = "Comprar por USB: " + vbCrLf + _
                    UB.GetNameUSB(H) + " (" + UB.GetLetterUSB(H) + ":\)" + vbCrLf + _
                    "Tiene " + CStr(UB.GetFreeMB(H)) + " MB libres"
                
                btBuy(H).Tag = "USB"
                
                btBuy(H).Visible = True
                btBuy(H).TabIndex = btBuy(H - 1).TabIndex + 1
                
                'que se cargue despintado!!
                SelBT btBuy(H), False
            Next H
        End If
        
        
    End If
    
    tERR.Anotar "xsaf", TengoBluetooth
    If TengoBluetooth Then
        tERR.Anotar "xsag", BTM.Count
        If BTM.Count > 0 Then
        
            btBuy(0).Visible = False
            
            H = btBuy.Count
            tERR.Anotar "xsah", H
            
            Dim CBT As TbrBtDevice
            tERR.Anotar "xsah2"
            
            For Each CBT In BTM
                tERR.Anotar "xsah3", H
                
                Load btBuy(H)
                
                If H = 1 Then
                    btBuy(H).Top = btBuy(H - 1).Top
                Else
                    btBuy(H).Top = btBuy(H - 1).Top + btBuy(H - 1).Height
                End If
                tERR.Anotar "xsah4", btBuy(H).Top
                
                tERR.Anotar "xsai", CBT.name, CBT.getAddress
                btBuy(H).Caption = "Comprar en Bluetooth: " + vbCrLf + _
                    CBT.name + " (" + CBT.getAddress + ")"
                    
                btBuy(H).Tag = "BT " + CBT.getAddress 'PARA PODER USARLO
                
                btBuy(H).Visible = True
                btBuy(H).TabIndex = btBuy(H - 1).TabIndex + 1
                'que se cargue despintado!!
                SelBT btBuy(H), False
                
                H = H + 1
            Next
            tERR.Anotar "xsaj"
        End If
    End If
    
    tERR.Anotar "xsak"
    
    If UB.GetCantidadUSB = 0 And BTM.Count = 0 Then
        If TengoBluetooth Then
            btBuy(0).Caption = "NO HAY DISPOSITIVOS" + vbCrLf + "Inserte USB o presione aqui para buscar por bluetooth"
        Else
            btBuy(0).Caption = "NO HAY DISPOSITIVOS" + vbCrLf + "Inserte USB ahora"
        End If
        btBuy(0).Visible = True
    End If
    
    AcomodarIndicadores
End Sub

Private Sub UnloadBtBuy()
    Dim H As Long
    For H = 1 To btBuy.Count - 1
        Unload btBuy(H)
    Next H
End Sub

Private Function LoadLista()
    
    'ver si se puede mostrar todo. Si quiedara muy chiquito ponemos algun mensaje
    Dim MinH As Long 'minimo de alto que muestro
    
    tERR.Anotar "daaq", Carrito.GetFileCant
    
    If Carrito.GetFileCant > 0 Then
        Dim H As Long
        For H = 1 To Carrito.GetFileCant
            ShowElem H
        Next H
    Else
        
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
        If K.LICENCIA("3pm") = HSuperLicencia Then
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UnloadBtBuy
    'aviso que no hay mas dispositivos
    BTM.ReiniciarColeccion
    Timer1.Interval = 0
End Sub

Private Sub Form_Resize()
    
    'Me.PaintPicture frmIndex.picFondoDisco.Image, 0, 0, Me.Width, Me.Height
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
    
    Image1.Left = Me.Width - Image1.Width - 60
    Image1.Top = 30
    Label3.Left = btSalir.Left + btSalir.Width + 30
    Label3.Width = Me.Width - Label3.Left - 60 - Image1.Width
    Label2.Width = Me.Width - Label2.Left - 60 - Image1.Width
    'Line1.Y2 = Me.Height
    
    'acomodar indicadores
    AcomodarIndicadores
    
    LoadLista

End Sub

Private Sub AcomodarIndicadores()
    
    tERR.Anotar "xsal", btBuy.Count - 1
    
    btANULA.Top = btBuy(btBuy.Count - 1).Top + btBuy(btBuy.Count - 1).Height
    btReview.Top = btANULA.Top + btANULA.Height
    Label1.Top = btReview.Top + btReview.Height + 60
    
    Dim TotIndic As Long 'total de indicadores
    TotIndic = 9
    'arrimar
    Label1.Left = 0
    Label4.Left = Label1.Left
    Label5.Left = Label1.Left
    Label6.Left = Label1.Left
    Label7.Left = Label1.Left
    Label8.Left = Label1.Left
    
    tERR.Anotar "xsam"
    
    'Label1.Width = Line1.X1 - 30
    Label4.Width = Label1.Width
    Label5.Width = Label1.Width
    Label6.Width = Label1.Width
    Label7.Width = Label1.Width
    Label8.Width = Label1.Width
    
    'Label1.Height = (Me.Height - Line2.Y1) / TotIndic
    Label4.Height = Label1.Height
    Label5.Height = Label1.Height
    Label6.Height = Label1.Height
    Label7.Height = Label1.Height
    Label8.Height = Me.Height - Label1.Top
    
    
    Label4.Top = Label1.Top + Label1.Height - 15
    Label5.Top = Label4.Top + Label1.Height + 490
    
    Label6.Top = Label5.Top + Label1.Height - 15
    
    Label7.AutoSize = True
    Label7.Top = Label6.Top + Label6.Height + 490
    
    Label8.Top = Label7.Top + Label7.Height + 490
    tERR.Anotar "xsan"
End Sub

Private Sub tBT_Change()
    
    If tBT.Text = "" Then Exit Sub
    
    tERR.Anotar "BT=" + tBT.Text
    
    Dim SP() As String
    SP = Split(tBT.Text, "|")
    
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
            
            'SE RECLAVA SI DESCONECTO BLUETOOTH DE PECHO
'            BT=4|General failed
'            BTM_SND_BAD
        Case "5"
            'encontro un dispositivo
            tERR.Anotar "BTM_DEV", SP(1), SP(2)
    End Select
    
    tBT.Text = ""
End Sub

Private Sub Timer1_Timer()
    UB.RefreshDriveList
End Sub

Private Sub tNADA_Change()
    If tNADA.Text = "" Then Exit Sub
    tERR.Anotar "CarUSB", tNADA.Text
    Dim SP() As String
    SP = Split(tNADA.Text, "|")
    
    Select Case SP(0)
        Case "0" 'entro drive
            UpdateData False
            UpdateDrives
            
        Case "1" 'sale drive
            UpdateData False
            UpdateDrives
    End Select
    
    tNADA.Text = ""
End Sub
