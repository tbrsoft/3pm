VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmCarritoBuy 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin tbr3pm.tbrFullProc SW 
      Height          =   555
      Left            =   7170
      TabIndex        =   7
      Top             =   3810
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   979
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   705
      Left            =   2970
      TabIndex        =   1
      Top             =   3780
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1244
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Volver a detectar dispositivos compatibles"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton2 
      Height          =   705
      Left            =   2970
      TabIndex        =   2
      Top             =   4500
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1244
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir sin vaciar carrito"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fbDEV 
      Height          =   1125
      Index           =   0
      Left            =   450
      TabIndex        =   0
      Top             =   1380
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
   Begin tbrFaroButton.fBoton fBoton3 
      Height          =   705
      Left            =   2970
      TabIndex        =   8
      Top             =   5220
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1244
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir eliminando la compra"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   120
      Picture         =   "frmCarritoBuy.frx":0000
      Top             =   4140
      Width           =   2205
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
      Left            =   60
      TabIndex        =   6
      Top             =   3420
      Width           =   8655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dispositivos encontrados. Seleccione dispositivo."
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
      Left            =   90
      TabIndex        =   5
      Top             =   1020
      Width           =   8655
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
      Left            =   90
      TabIndex        =   4
      Top             =   570
      Width           =   8385
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
      Left            =   60
      TabIndex        =   3
      Top             =   180
      Width           =   8385
   End
End
Attribute VB_Name = "frmCarritoBuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CantDEV As Long 'cantidad de dispositivos totales

Private Sub fbDEV_Click(Index As Integer)
    
    tERR.Anotar "daaw", Index, fbDEV(Index).Caption
    
    'VER SI TIENE QUE IR A BORRAR COSAS O A GRABAR!!!!
    UB.DevSel = Index + 1 'el boton con index 0 corresponde al dispositivo 1
    Select Case fbDEV(Index).Tag
        Case "NO" 'ni borrando hace lugar
            SW.ShowWait "El tamaño de la compra supera el tamaño TOTAL del dispositivo" + vbCrLf + _
                "El carrito se vaciara para que reformule su compra", 6500
            Carrito.ClearCart
            Unload Me
        Case "BORRA" 'si borra puede llegar
            SW.ShowWait "El tamaño de la compra supera el tamaño disponible del dispositivo" + vbCrLf + _
                "Elija (o no) de la lista siguiente lo que desee eliminar para " + _
                "hacer lugar suficiente", 6500
                
            Me.Visible = False
            frmCarritoDelete.Show 1
            Unload Me
            Exit Sub
        Case Else
            UB.DevSel = Index + 1 'marcarlo para leer
            SW.ShowWait "Leyendo dispositivo ..."
            frmCarritoInDev.ShowDEV fbDEV(Index).Tag
            Unload Me
    End Select
    
    Unload Me
End Sub

Private Sub fbDEV_GotFocus(Index As Integer)
    SelBT fbDEV(Index), True
End Sub

Private Sub fbDEV_LostFocus(Index As Integer)
    SelBT fbDEV(Index), False
End Sub

Private Sub fBoton1_Click()
    
    SW.ShowWait "Detectando dispositivos ...."
    tERR.Anotar "daax"
    ListarUSB
    
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

Private Sub fBoton3_Click()
    KeyCode = 0
    Carrito.ClearCart
    Unload Me
End Sub

Private Sub fBoton3_GotFocus()
    SelBT fBoton3, True
End Sub

Private Sub fBoton3_LostFocus()
    SelBT fBoton3, False
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
    Pintar_fBoton Me
    Label2.Caption = "Tamaño de la compra: " + CStr(Carrito.GetTotalMB) + " MB"
    tERR.Anotar "daba"
    ListarUSB
End Sub

Private Sub ListarUSB()

    On Local Error GoTo MER
    CantDEV = 0
    
    CantDEV = CantDEV + UB.GetCantidadUSB 'despues se deberían ir sumando
    'CantDEV = CantDEV + dispositivos bluetooth
    'CantDEV = CantDEV + dispositivos infrarojos
    tERR.Anotar "dabb", CantDEV
    If CantDEV = 0 Then
        fbDEV(0).Caption = "No hay dispositivos"
    Else 'AL MENOS HAY UNO
        Dim H As Long
        For H = 1 To UB.GetCantidadUSB
            'otros tipos de dispositivos
            If H > 1 Then
                Load fbDEV(H - 1)
                fbDEV(H - 1).TabIndex = H - 1
            End If
            'para poder abrir despues lo que corresponde
            fbDEV(H - 1).Tag = UB.GetLetterUSB(H)
            'para que el tipo elija comodo
            fbDEV(H - 1).Caption = "Dispositivo: " + UB.GetNameUSB(H) + " (" + UB.GetLetterUSB(H) + ":\)" + vbCrLf + _
                CStr(UB.GetFreeMB(H)) + " MB libres / " + CStr(UB.GetTotalMB(H)) + " MB"
            tERR.Anotar "dabc", fbDEV(H - 1).Caption
            'si no alcanza el tamaño desactivarlo !!
            If Carrito.GetTotalMB > UB.GetFreeMB(H) Then
                tERR.Anotar "dabd", Carrito.GetTotalMB, UB.GetFreeMB(H)
                If Carrito.GetTotalMB > UB.GetTotalMB(H) Then
                    fbDEV(H - 1).Caption = "ESPACIO INALCANZABLE EN EL DISPOSITIVO" + vbCrLf + _
                        fbDEV(H - 1).Caption
                    fbDEV(H - 1).Tag = "NO"
                Else
                    fbDEV(H - 1).Caption = "ESPACIO INSUFICIENTE DEBERA ELIMINAR" + vbCrLf + _
                        fbDEV(H - 1).Caption
                    fbDEV(H - 1).Tag = "BORRA"
                End If
            End If
        Next H
    End If
    tERR.Anotar "dabe"
    Acomodar

    Exit Sub
MER:
    tERR.AppendLog tERR.ErrToTXT(Err), "cpCC5"
    Resume Next

End Sub

Private Sub Acomodar()
    Label1.Left = 60
    Label1.Width = Me.Width - 120
    
    Label2.Left = Label1.Left
    Label2.Width = Label1.Width
    
    Label3.Left = Label1.Left
    Label3.Width = Label1.Width
    
    Label4.Left = Label1.Left
    Label4.Width = Label1.Width
    
    fbDEV(0).Left = Me.Width / 2 - fbDEV(0).Width / 2
    
    Dim H As Long
    If fbDEV.Count > 1 Then
        For H = 1 To fbDEV.Count - 1
            fbDEV(H).Top = fbDEV(H - 1).Top + fbDEV(H - 1).Height + 60
            fbDEV(H).Left = Me.Width / 2 - fbDEV(H).Width / 2
            fbDEV(H).Visible = True
        Next H
        
        Label4.Top = fbDEV(fbDEV.Count - 1).Top + fbDEV(fbDEV.Count - 1).Height + 60
    Else
        Label4.Top = fbDEV(0).Top + fbDEV(0).Height + 60
    End If
    
    fBoton1.Top = Label4.Top + Label4.Height + 60
    fBoton2.Top = fBoton1.Top + fBoton1.Height + 60
    fBoton3.Top = fBoton2.Top + fBoton2.Height + 60
    
    fBoton1.Left = Me.Width / 2 - fBoton1.Width / 2
    fBoton2.Left = Me.Width / 2 - fBoton1.Width / 2
    fBoton3.Left = Me.Width / 2 - fBoton1.Width / 2
    
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
    Acomodar
End Sub
