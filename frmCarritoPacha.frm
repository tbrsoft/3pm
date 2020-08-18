VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmCarritoPacha 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox tUSBPacha 
      Height          =   435
      Left            =   5730
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.TextBox tBTPacha 
      Height          =   435
      Left            =   5760
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   2565
   End
   Begin tbr3pm.tbrFullProc SW 
      Height          =   555
      Left            =   7110
      TabIndex        =   0
      Top             =   3630
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   979
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   705
      Left            =   2910
      TabIndex        =   1
      Top             =   3600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1244
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Buscar Bluetooth"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton2 
      Height          =   705
      Left            =   2910
      TabIndex        =   2
      Top             =   4320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1244
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "SALIR"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton btBuy 
      Height          =   1125
      Index           =   0
      Left            =   390
      TabIndex        =   3
      Top             =   1200
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   1984
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "No hay dispositivos"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Defina dispositivo a grabar."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8385
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño de la compra"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   30
      TabIndex        =   5
      Top             =   390
      Width           =   8385
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Otras opciones"
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
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   8655
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   60
      Picture         =   "frmCarritoPacha.frx":0000
      Top             =   3960
      Width           =   2205
   End
End
Attribute VB_Name = "frmCarritoPacha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CantDEV As Long 'cantidad de dispositivos totales

Private Sub btBuy_Click(Index As Integer)
        If TengoBluetooth And Index = 0 Then
    
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
                MT = Int(Rnd * 6) + 1
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
End Sub

Private Sub btBuy_GotFocus(Index As Integer)
    SelBT btBuy(Index), True
End Sub

Private Sub btBuy_LostFocus(Index As Integer)
    SelBT btBuy(Index), False
End Sub

Private Sub fBoton1_Click()
    
    SW.ShowWait "Detectando dispositivos ...."
    tERR.Anotar "daax"
    'ListarUSB
    
    SW.ShowWait ""
End Sub

Private Sub fBoton1_GotFocus()
    SelBT fBoton1, True
End Sub

Private Sub fBoton1_LostFocus()
    SelBT fBoton1, False
End Sub

Private Sub fBoton2_Click()
    Unload Me
End Sub

Private Sub fBoton2_GotFocus()
    SelBT fBoton2, True
End Sub

Private Sub fBoton2_LostFocus()
    SelBT fBoton2, False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    tERR.Anotar "daaz", KeyCode, Chr(KeyCode)
    Select Case KeyCode
        Case TeclaOK
            If TeclaOK <> 13 And TeclaOK <> 108 Then
                SendKeys "{ENTER}"
            End If
        Case TeclaDER: SendKeys "{TAB}"
        Case TeclaIZQ: SendKeys "+{TAB}"
        Case TeclaPagAd: SendKeys "{TAB}"
        Case TeclaPagAt: SendKeys "+{TAB}"
        Case TeclaESC: Unload Me
        Case TeclaCarrito: SendKeys "{ENTER}"
    End Select
End Sub

Private Sub Form_Load()

    If TengoBluetooth Then
        BTM.UseEventMSG tBTPacha.HWND
    End If
    
    Me.AutoRedraw = True
    Pintar_fBoton Me
    
    UB.UseEventMSG tUSBPacha.HWND
    
    Label2.Caption = "Tamaño de la compra: " + CStr(Carrito.GetTotalMB) + " MB"
    tERR.Anotar "daba"
    'ListarUSB
End Sub

Private Sub Acomodar()
    Label1.Left = 60
    Label1.Width = Me.Width - 120
    
    Label2.Left = Label1.Left
    Label2.Width = Label1.Width
    
    Label4.Left = Label1.Left
    Label4.Width = Label1.Width
    
    btBuy(0).Left = Me.Width / 2 - btBuy(0).Width / 2
    
    Dim H As Long
    If btBuy.Count > 1 Then
        For H = 1 To btBuy.Count - 1
            btBuy(H).Top = btBuy(H - 1).Top + btBuy(H - 1).Height + 60
            btBuy(H).Left = Me.Width / 2 - btBuy(H).Width / 2
            btBuy(H).Visible = True
        Next H
        
        Label4.Top = btBuy(btBuy.Count - 1).Top + btBuy(btBuy.Count - 1).Height + 60
    Else
        Label4.Top = btBuy(0).Top + btBuy(0).Height + 60
    End If
    
    fBoton1.Top = Label4.Top + Label4.Height + 60
    fBoton2.Top = fBoton1.Top + fBoton1.Height + 60
    
    fBoton1.Left = Me.Width / 2 - fBoton1.Width / 2
    fBoton2.Left = Me.Width / 2 - fBoton1.Width / 2
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Carrito.ClearCart
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
    Acomodar
End Sub

Private Sub tUSBPacha_Change()
    If tUSBPacha.Text = "" Then Exit Sub
    tERR.Anotar "CarUSB", tUSBPacha.Text
    Dim SP() As String
    SP = Split(tUSBPacha.Text, "|")
    
    Select Case SP(0)
        Case "0" 'entro drive
            'UpdateData False
            UpdateDrives
            
        Case "1" 'sale drive
            'UpdateData False
            UpdateDrives
    End Select
    
    tUSBPacha.Text = ""
End Sub

'XXXXXXXXXXXXXX
'me falta hacer que funcione como en el otro
'xxxxxxxxxxxxxx
Private Sub UpdateDrives()
    tERR.Anotar "xsaa"
    UnloadBtBuy
    
    tERR.Anotar "xsab", UB.GetCantidadUSB
    If UB.GetCantidadUSB > 0 Then
        btBuy(0).Visible = False
        
        'SI HAY MAS DE UNO QUE SEPA QUIEN ES QUIEN
        tERR.Anotar "xsad", UB.GetCantidadUSB
        If UB.GetCantidadUSB >= 1 Then
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
                
                tERR.Anotar "xsai", CBT.Name, CBT.getAddress
                btBuy(H).Caption = "Comprar en Bluetooth: " + vbCrLf + _
                    CBT.Name + " (" + CBT.getAddress + ")"
                    
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
    
    If UB.GetCantidadUSB = 0 Then
        If TengoBluetooth Then
            If BTM.Count = 0 Then
                btBuy(0).Caption = "NO HAY DISPOSITIVOS" + vbCrLf + "Inserte USB o presione aqui para buscar bluetooth"
            End If
        Else
            btBuy(0).Caption = "NO HAY DISPOSITIVOS" + vbCrLf + "Inserte USB ahora"
        End If
        btBuy(0).Visible = True
    End If
    
    Acomodar
End Sub

Private Sub UnloadBtBuy()
    Dim H As Long
    For H = 1 To btBuy.Count - 1
        Unload btBuy(H)
    Next H
End Sub
