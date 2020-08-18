VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmConfigCartPrecios 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Precios de ventas de multimedia"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frSELL 
      BackColor       =   &H00000000&
      Caption         =   "Precios"
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
      Height          =   8355
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   9315
      Begin VB.ListBox lstPROMOSA 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   600
         IntegralHeight  =   0   'False
         Left            =   60
         TabIndex        =   12
         Top             =   2250
         Width           =   5655
      End
      Begin VB.ComboBox cmbTipoFile 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmConfigCartPrecios.frx":0000
         Left            =   1650
         List            =   "frmConfigCartPrecios.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   780
         Width           =   2985
      End
      Begin VB.TextBox txtCantFileCarrito 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   810
         Width           =   600
      End
      Begin VB.VScrollBar vsCantFileCarrito 
         Height          =   330
         LargeChange     =   10
         Left            =   6540
         Max             =   0
         Min             =   100
         TabIndex        =   9
         Top             =   810
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtPrecioTotalCarrito 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8310
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   825
         Width           =   600
      End
      Begin VB.VScrollBar vsPrecioTotalCarrito 
         Height          =   330
         LargeChange     =   10
         Left            =   8940
         Max             =   0
         Min             =   400
         TabIndex        =   7
         Top             =   825
         Value           =   1
         Width           =   330
      End
      Begin VB.ListBox lstPromosV 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   600
         IntegralHeight  =   0   'False
         Left            =   90
         TabIndex        =   6
         Top             =   3150
         Width           =   5655
      End
      Begin VB.ListBox lstPromosR 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   600
         IntegralHeight  =   0   'False
         Left            =   90
         TabIndex        =   5
         Top             =   4050
         Width           =   5655
      End
      Begin VB.ListBox lstPromosW 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   600
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   4
         Top             =   5820
         Width           =   5655
      End
      Begin VB.ListBox lstPromosJ 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   600
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   3
         Top             =   6750
         Width           =   5655
      End
      Begin VB.ListBox lstPromosI 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   600
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   2
         Top             =   7650
         Width           =   5655
      End
      Begin VB.ListBox lstPromos3 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   600
         IntegralHeight  =   0   'False
         Left            =   90
         TabIndex        =   1
         Top             =   4920
         Width           =   5655
      End
      Begin tbrFaroButton.fBoton fBoton2 
         Height          =   540
         Left            =   7620
         TabIndex        =   13
         Top             =   1890
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   953
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Agregar esta promoción"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton3 
         Height          =   330
         Left            =   5820
         TabIndex        =   14
         Top             =   2340
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "borrar elegida"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton5 
         Height          =   330
         Left            =   5790
         TabIndex        =   15
         Top             =   3270
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "borrar elegida"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton6 
         Height          =   330
         Left            =   5820
         TabIndex        =   16
         Top             =   4170
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "borrar elegida"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton7 
         Height          =   330
         Left            =   5850
         TabIndex        =   17
         Top             =   5940
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "borrar elegida"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton8 
         Height          =   330
         Left            =   5850
         TabIndex        =   18
         Top             =   6870
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "borrar elegida"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton10 
         Height          =   330
         Left            =   5850
         TabIndex        =   19
         Top             =   7770
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "borrar elegida"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton11 
         Height          =   330
         Left            =   5820
         TabIndex        =   20
         Top             =   5040
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "borrar elegida"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton1 
         Height          =   660
         Left            =   7950
         TabIndex        =   34
         Top             =   4140
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1164
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Grabar todo"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton4 
         Height          =   660
         Left            =   7950
         TabIndex        =   35
         Top             =   5160
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1164
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Salir sin grabar"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Precio total en créditos"
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
         Height          =   435
         Index           =   23
         Left            =   6600
         TabIndex        =   33
         Top             =   780
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de archivos"
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
         Height          =   495
         Index           =   22
         Left            =   4380
         TabIndex        =   32
         Top             =   780
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de archivo"
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
         Height          =   255
         Index           =   21
         Left            =   90
         TabIndex        =   31
         Top             =   870
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Puede agregar la cantidad de promociones que desee. Asegúrese que sean coherentes."
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
         Height          =   255
         Index           =   20
         Left            =   90
         TabIndex        =   30
         Top             =   480
         Width           =   9765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Se toma como referencia los créditos por señal de monedero especificado en la sección 'Créditos'"
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
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   10815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIOS DE AUDIO"
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
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   28
         Top             =   2010
         Width           =   2445
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIOS DE VIDEO"
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
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   27
         Top             =   2970
         Width           =   2025
      End
      Begin VB.Label lblDetPrec 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   150
         TabIndex        =   26
         Top             =   1260
         Width           =   9105
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIOS RINGTONES"
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
         Height          =   225
         Index           =   4
         Left            =   90
         TabIndex        =   25
         Top             =   3870
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIOS WALLPAPERS"
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
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Top             =   5640
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIOS JUEGOS JAVA"
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
         Height          =   225
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   6570
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIOS Imagenes ISO/Nero"
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
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   7470
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIOS Videos 3GP celulares"
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
         Height          =   225
         Index           =   8
         Left            =   90
         TabIndex        =   21
         Top             =   4740
         Width           =   2025
      End
   End
End
Attribute VB_Name = "frmConfigCartPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fBoton1_Click()
    Carrito.SavePrices GPF("promocart")
    Unload Me
End Sub

'mm91
Private Sub fBoton11_Click()

    If lstPromos3.ListIndex = -1 Then Exit Sub
    
    If lstPromos3.ListCount = 1 Then
        MsgBox "No puede dejar sin promociones!"
    Else
        KillPromo lstPromos3
    End If

End Sub

Private Sub fBoton4_Click()
    Unload Me
End Sub

Private Sub fBoton10_Click() 'mm91
    If lstPromosI.ListIndex = -1 Then Exit Sub
    
    'si solo queda una no se puede eliminar
    If lstPromosI.ListCount = 1 Then
        MsgBox "No puede dejar sin promociones!"
    Else
        KillPromo lstPromosI
    End If
End Sub

Private Sub fBoton2_Click()
    If vsCantFileCarrito.Value = 0 Then
        MsgBox "No puede especificar el precio para cero selecciones!"
        Exit Sub
    End If
    
    If cmbTipoFile.ListIndex = 0 Then 'ES AUDIO
        Carrito.SetPricesAudio vsCantFileCarrito.Value, vsPrecioTotalCarrito.Value
    End If
    
    If cmbTipoFile.ListIndex = 1 Then 'ES VIDEO
        Carrito.SetPricesVideo vsCantFileCarrito.Value, vsPrecioTotalCarrito.Value
    End If
    
    If cmbTipoFile.ListIndex = 2 Then 'ES ringtones
        Carrito.SetPricesRingtones vsCantFileCarrito.Value, vsPrecioTotalCarrito.Value
    End If
    
    If cmbTipoFile.ListIndex = 3 Then 'ES wallpapers
        Carrito.SetPricesWallpapers vsCantFileCarrito.Value, vsPrecioTotalCarrito.Value
    End If
    
    If cmbTipoFile.ListIndex = 4 Then 'ES java
        Carrito.SetPricesJava vsCantFileCarrito.Value, vsPrecioTotalCarrito.Value
    End If
    
    If cmbTipoFile.ListIndex = 5 Then 'ES imagen ISO 'mm91
        Carrito.SetPricesIso vsCantFileCarrito.Value, vsPrecioTotalCarrito.Value
    End If
    
    If cmbTipoFile.ListIndex = 6 Then 'ES video movil 'mm91
        Carrito.SetPrices3GP vsCantFileCarrito.Value, vsPrecioTotalCarrito.Value
    End If
    
    VerPromos
End Sub

Private Sub fBoton3_Click()
    
    If lstPROMOSA.ListIndex = -1 Then Exit Sub
    
    'si solo queda una no se puede eliminar
    If lstPROMOSA.ListCount = 1 Then
        MsgBox "No puede dejar sin promociones!"
    Else
        KillPromo lstPROMOSA
    End If
End Sub

Private Sub fBoton5_Click()
    If lstPromosV.ListIndex = -1 Then Exit Sub
    
    If lstPromosV.ListCount = 1 Then
        MsgBox "No puede dejar sin promociones!"
    Else
        KillPromo lstPromosV
    End If
End Sub

Private Sub KillPromo(S As String)

    'leer que es lo que se quiere eliminar
    Dim SP() As String, TP As Long, Cant As Long
    SP = Split(S)
    TP = CLng(SP(0))
    Cant = CLng(SP(2))
    
    If Cant = 1 Then
        MsgBox "No puede borrar el precio de un solo archivo!" + vbCrLf + "Es obligatorio."
        Exit Sub
    End If
    
    If TP = 1 Then ' ES MUSICA
        Carrito.KillPricesAudioBase Cant
    End If
    
    If TP = 2 Then ' ES video
        Carrito.KillPricesVideoBase Cant
    End If
    
    If TP = 3 Then ' ES ringtone
        Carrito.KillPricesRingtonesBase Cant
    End If
    
    If TP = 4 Then ' ES wallpaper
        Carrito.KillPricesWallpapersBase Cant
    End If
    
    If TP = 5 Then ' ES java
        Carrito.KillPricesJAVABase Cant
    End If
    
    If TP = 6 Then ' ES iso/nero    'mm91
        Carrito.KillPricesISOBase Cant
    End If
    
    If TP = 7 Then ' ES videos 3gp  'mm91
        Carrito.KillPrices3GPBase Cant
    End If
    
    
    VerPromos
End Sub

Private Sub fBoton6_Click()
    If lstPromosR.ListIndex = -1 Then Exit Sub
    
    If lstPromosR.ListCount = 1 Then
        MsgBox "No puede dejar sin promociones!"
    Else
        KillPromo lstPromosR
    End If
End Sub

Private Sub fBoton7_Click()
    If lstPromosW.ListIndex = -1 Then Exit Sub
    
    If lstPromosW.ListCount = 1 Then
        MsgBox "No puede dejar sin promociones!"
    Else
        KillPromo lstPromosW
    End If
End Sub

Private Sub fBoton8_Click()
    If lstPromosJ.ListIndex = -1 Then Exit Sub
    
    If lstPromosJ.ListCount = 1 Then
        MsgBox "No puede dejar sin promociones!"
    Else
        KillPromo lstPromosJ
    End If
End Sub

Private Sub VerPromos()
    'ver las promociones ya grabadas
    Dim H As Long
    
    Dim SN As Single
    
    lstPROMOSA.Clear
    For H = 1 To Carrito.GetTotalPricesAudio
        If Carrito.GetPricesAudioBase(H) > 0 Then
            SN = Round(Carrito.GetPricesAudioBase(H) * PrecioBase / TemasPorCredito, 2)
            
            If H = 1 Then
                lstPROMOSA.AddItem "1 - " + CStr(H) + " fichero de AUDIO " + _
                    " por $ " + CStr(SN)
                    
            Else
                lstPROMOSA.AddItem "1 - " + CStr(H) + " ficheros de AUDIO " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)"
            End If
        End If
    Next H
    
    lstPromosV.Clear
    For H = 1 To Carrito.GetTotalPricesVideo
        If Carrito.GetPricesVideoBase(H) > 0 Then
            SN = Round(Carrito.GetPricesVideoBase(H) * PrecioBase / TemasPorCredito, 2)
            If H = 1 Then
                lstPromosV.AddItem "2 - " + CStr(H) + " fichero de VIDEO " + _
                    " por $ " + _
                    CStr(SN)
            Else
                lstPromosV.AddItem "2 - " + CStr(H) + " ficheros de VIDEO " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)"
            End If
        End If
    Next H
    
    lstPromosR.Clear
    For H = 1 To Carrito.GetTotalPricesRingtones
        If Carrito.GetPricesRingtonesBase(H) > 0 Then
            SN = Round(Carrito.GetPricesRingtonesBase(H) * PrecioBase / TemasPorCredito, 2)
            
            If H = 1 Then
                lstPromosR.AddItem "3 - " + CStr(H) + " ringtone " + _
                    " por $ " + CStr(SN)
                    
            Else
                lstPromosR.AddItem "3 - " + CStr(H) + " ringtones " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)"
            End If
        End If
    Next H
    
    lstPromosW.Clear
    For H = 1 To Carrito.GetTotalPricesWallpapers
        If Carrito.GetPricesWallpapersBase(H) > 0 Then
            SN = Round(Carrito.GetPricesWallpapersBase(H) * PrecioBase / TemasPorCredito, 2)
            
            If H = 1 Then
                lstPromosW.AddItem "4 - " + CStr(H) + " wallpaper " + _
                    " por $ " + CStr(SN)
                    
            Else
                lstPromosW.AddItem "4 - " + CStr(H) + " wallpapers " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)"
            End If
        End If
    Next H
    
    lstPromosJ.Clear
    For H = 1 To Carrito.GetTotalPricesJAVA
        If Carrito.GetPricesJAVABase(H) > 0 Then
            SN = Round(Carrito.GetPricesJAVABase(H) * PrecioBase / TemasPorCredito, 2)
            
            If H = 1 Then
                lstPromosJ.AddItem "5 - " + CStr(H) + " juego java " + _
                    " por $ " + CStr(SN)
                    
            Else
                lstPromosJ.AddItem "5 - " + CStr(H) + " juegos java " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)"
            End If
        End If
    Next H
    
    lstPromosI.Clear 'mm91
    For H = 1 To Carrito.GetTotalPricesISO
        If Carrito.GetPricesISOBase(H) > 0 Then
            SN = Round(Carrito.GetPricesISOBase(H) * PrecioBase / TemasPorCredito, 2)
            
            If H = 1 Then
                lstPromosI.AddItem "6 - " + CStr(H) + " imagen iso/nero " + _
                    " por $ " + CStr(SN)
                    
            Else
                lstPromosI.AddItem "6 - " + CStr(H) + " imagenes iso/nero " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)"
            End If
        End If
    Next H
    
    lstPromos3.Clear 'mm91
    For H = 1 To Carrito.GetTotalPrices3GP
        If Carrito.GetPrices3GPBase(H) > 0 Then
            SN = Round(Carrito.GetPrices3GPBase(H) * PrecioBase / TemasPorCredito, 2)
            
            If H = 1 Then
                lstPromos3.AddItem "7 - " + CStr(H) + " video 3GP " + _
                    " por $ " + CStr(SN)
                    
            Else
                lstPromos3.AddItem "7 - " + CStr(H) + " videos 3GP " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)"
            End If
        End If
    Next H
    
End Sub

Private Sub Form_Load()
    'xxxxxx poner todas estas en todos los lugares que correspondan!!!!
    
    VerPromos
    cmbTipoFile.Clear
    cmbTipoFile.AddItem "AUDIO"
    cmbTipoFile.AddItem "VIDEO"
    cmbTipoFile.AddItem "Ringtones"
    cmbTipoFile.AddItem "Wallpapers"
    cmbTipoFile.AddItem "Juegos Java" 'juegos o aplicaciones 'METER DE PECHO tbrMobileCash
    cmbTipoFile.AddItem "Imagenes ISO" 'incluye tambien las de nero 'mm91
    cmbTipoFile.AddItem "Videos 3GP" 'mm91
    
    cmbTipoFile.ListIndex = 0
    
    vsCantFileCarrito.Value = 0
    vsPrecioTotalCarrito.Value = 0
End Sub

Private Sub vsCantFileCarrito_Change()
    txtCantFileCarrito.tExt = vsCantFileCarrito.Value
    UpDesc
End Sub

Private Sub vsPrecioTotalCarrito_Change()
    txtPrecioTotalCarrito.tExt = vsPrecioTotalCarrito.Value
    UpDesc
End Sub

Private Sub UpDesc()
    Dim PP As Single
    PP = Round(vsPrecioTotalCarrito.Value * PrecioBase / TemasPorCredito, 2)
    
    Dim TP As String
    If cmbTipoFile.ListIndex = 0 Then TP = "Audio"
    If cmbTipoFile.ListIndex = 1 Then TP = "Video"
    If cmbTipoFile.ListIndex = 2 Then TP = "Ringtones"
    If cmbTipoFile.ListIndex = 3 Then TP = "Wallpapers"
    If cmbTipoFile.ListIndex = 4 Then TP = "Juegos JAVA"
    If cmbTipoFile.ListIndex = 5 Then TP = "Imagenes ISO/Nero" 'mm91
    If cmbTipoFile.ListIndex = 6 Then TP = "Videos 3GP" 'mm91
    
    If vsCantFileCarrito.Value = 0 Or vsPrecioTotalCarrito.Value = 0 Then
        lblDetPrec.Caption = "Paquete no válido"
        Exit Sub
    End If
    
    If vsCantFileCarrito.Value = 1 Then
        lblDetPrec.Caption = "Paquete de " + CStr(vsCantFileCarrito.Value) + _
            " fichero de " + TP + ": $ " + CStr(PP)
    Else
        lblDetPrec.Caption = "Paquete de " + CStr(vsCantFileCarrito.Value) + _
            " ficheros de " + TP + ": $ " + _
            CStr(PP) + " cada uno $ " + CStr(Round(PP / vsCantFileCarrito.Value, 2))
    End If
End Sub

Private Sub cmbTipoFile_Click()
    UpDesc
End Sub

