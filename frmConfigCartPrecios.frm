VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmConfigCartPrecios 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Precios de ventas de multimedia"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   11550
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
      Height          =   5955
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11445
      Begin VB.ListBox lstPromosI_DVD 
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
         Left            =   5730
         TabIndex        =   39
         Top             =   3120
         Width           =   5200
      End
      Begin VB.ListBox lstPromosT 
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
         Left            =   5790
         TabIndex        =   34
         Top             =   4770
         Width           =   5200
      End
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
         Left            =   30
         TabIndex        =   12
         Top             =   2010
         Width           =   5200
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
         Left            =   5280
         List            =   "frmConfigCartPrecios.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   180
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
         Left            =   5280
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   660
         Width           =   600
      End
      Begin VB.VScrollBar vsCantFileCarrito 
         Height          =   330
         LargeChange     =   10
         Left            =   5940
         Max             =   0
         Min             =   100
         TabIndex        =   9
         Top             =   660
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
         Left            =   5310
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1185
         Width           =   600
      End
      Begin VB.VScrollBar vsPrecioTotalCarrito 
         Height          =   330
         LargeChange     =   10
         Left            =   5940
         Max             =   0
         Min             =   400
         TabIndex        =   7
         Top             =   1185
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
         Left            =   30
         TabIndex        =   6
         Top             =   2820
         Width           =   5200
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
         Left            =   60
         TabIndex        =   5
         Top             =   3600
         Width           =   5200
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
         Left            =   30
         TabIndex        =   4
         Top             =   5190
         Width           =   5200
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
         Left            =   5730
         TabIndex        =   3
         Top             =   2340
         Width           =   5200
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
         Left            =   5760
         TabIndex        =   2
         Top             =   3990
         Width           =   5200
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
         Left            =   30
         TabIndex        =   1
         Top             =   4410
         Width           =   5200
      End
      Begin tbrFaroButton.fBoton fBoton2 
         Height          =   1080
         Left            =   9750
         TabIndex        =   13
         Top             =   360
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   1905
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
         Left            =   5250
         TabIndex        =   14
         Top             =   2070
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "x"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton5 
         Height          =   330
         Left            =   5250
         TabIndex        =   15
         Top             =   2880
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "x"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton6 
         Height          =   330
         Left            =   5280
         TabIndex        =   16
         Top             =   3660
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "x"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton7 
         Height          =   330
         Left            =   5250
         TabIndex        =   17
         Top             =   5250
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "x"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton8 
         Height          =   330
         Left            =   10950
         TabIndex        =   18
         Top             =   2370
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "x"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton10 
         Height          =   330
         Left            =   10980
         TabIndex        =   19
         Top             =   3960
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "x"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton11 
         Height          =   330
         Left            =   5280
         TabIndex        =   20
         Top             =   4440
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "x"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton9 
         Height          =   330
         Left            =   10950
         TabIndex        =   35
         Top             =   4740
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "x"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton12 
         Height          =   330
         Left            =   10950
         TabIndex        =   40
         Top             =   3120
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   582
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "x"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIOS Imagenes ISO/Nero DE CD"
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
         Index           =   3
         Left            =   5760
         TabIndex        =   41
         Top             =   3750
         Width           =   3495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   60
         X2              =   11460
         Y1              =   1710
         Y2              =   1710
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIOS Temas para mobil"
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
         Index           =   2
         Left            =   5760
         TabIndex        =   36
         Top             =   4590
         Width           =   3495
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
         Left            =   3600
         TabIndex        =   33
         Top             =   1140
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
         Left            =   3780
         TabIndex        =   32
         Top             =   630
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
         Left            =   3780
         TabIndex        =   31
         Top             =   270
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
         Height          =   645
         Index           =   20
         Left            =   90
         TabIndex        =   30
         Top             =   900
         Width           =   3345
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
         Height          =   645
         Index           =   19
         Left            =   120
         TabIndex        =   29
         Top             =   180
         Width           =   3315
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
         Top             =   1830
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
         Left            =   60
         TabIndex        =   27
         Top             =   2640
         Width           =   2025
      End
      Begin VB.Label lblDetPrec 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   915
         Left            =   6360
         TabIndex        =   26
         Top             =   600
         Width           =   3165
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
         Left            =   60
         TabIndex        =   25
         Top             =   3420
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
         Left            =   30
         TabIndex        =   24
         Top             =   5010
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
         Left            =   5730
         TabIndex        =   23
         Top             =   2160
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIOS Imagenes ISO/Nero DE DVD"
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
         Left            =   5730
         TabIndex        =   22
         Top             =   2940
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
         Left            =   30
         TabIndex        =   21
         Top             =   4230
         Width           =   2025
      End
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   390
      Left            =   30
      TabIndex        =   37
      Top             =   6060
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   688
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar todo"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton4 
      Height          =   390
      Left            =   2910
      TabIndex        =   38
      Top             =   6060
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   688
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir sin grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
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

Private Sub fBoton12_Click() 'mp01
    If lstPromosI_DVD.ListIndex = -1 Then Exit Sub
    
    'si solo queda una no se puede eliminar
    If lstPromosI_DVD.ListCount = 1 Then
        MsgBox "No puede dejar sin promociones!"
    Else
        KillPromo lstPromosI_DVD
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
    
    If cmbTipoFile.ListIndex = 7 Then 'ES video movil 'mm91
        Carrito.SetPricesThemes vsCantFileCarrito.Value, vsPrecioTotalCarrito.Value
    End If
    
    If cmbTipoFile.ListIndex = 8 Then 'iso dvd mp01
        Carrito.SetPricesIsoDVD vsCantFileCarrito.Value, vsPrecioTotalCarrito.Value
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
    
    If TP = 6 Then ' ES iso/nero
        Carrito.KillPricesISOBase Cant
    End If
    
    If TP = 7 Then ' ES videos 3gp
        Carrito.KillPrices3GPBase Cant
    End If
    
    If TP = 8 Then ' temas para celular
        Carrito.KillPricesThemesBase Cant
    End If
    
    If TP = 9 Then ' iso dvd mp01
        Carrito.KillPricesISODVDBase Cant
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
                'mp01
                lstPromosI.AddItem "6 - " + CStr(H) + " imagen iso/nero CD" + _
                    " por $ " + CStr(SN)
                    
            Else
                'mp01
                lstPromosI.AddItem "6 - " + CStr(H) + " imagenes iso/nero CD" + _
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
    
    lstPromosT.Clear 'mm91
    For H = 1 To Carrito.GetTotalPricesThemes
        If Carrito.GetPricesThemesBase(H) > 0 Then
            SN = Round(Carrito.GetPricesThemesBase(H) * PrecioBase / TemasPorCredito, 2)
            
            If H = 1 Then
                lstPromosT.AddItem "8 - 1 tema para movil " + _
                    " por $ " + CStr(SN)
                    
            Else
                lstPromosT.AddItem "8 - " + CStr(H) + " temas para movil " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)"
            End If
        End If
    Next H
    
    'mp01
    lstPromosI_DVD.Clear
    For H = 1 To Carrito.GetTotalPricesISODVD
        If Carrito.GetPricesISODVDBase(H) > 0 Then
            SN = Round(Carrito.GetPricesISODVDBase(H) * PrecioBase / TemasPorCredito, 2)
            
            If H = 1 Then
                lstPromosI_DVD.AddItem "9 - " + CStr(H) + " imagen iso/nero DVD" + _
                    " por $ " + CStr(SN)
                    
            Else
                lstPromosI_DVD.AddItem "9 - " + CStr(H) + " imagen iso/nero DVD" + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)"
            End If
        End If
    Next H
    
End Sub

Private Sub fBoton9_Click()
    If lstPromosT.ListIndex = -1 Then Exit Sub
    
    'si solo queda una no se puede eliminar
    If lstPromosT.ListCount = 1 Then
        MsgBox "No puede dejar sin promociones!"
    Else
        KillPromo lstPromosT
    End If
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
    cmbTipoFile.AddItem "Imagenes ISO/NRG CD" 'incluye tambien las de nero 'mp01
    cmbTipoFile.AddItem "Videos 3GP"
    cmbTipoFile.AddItem "Temas para movil"
    cmbTipoFile.AddItem "Imagenes ISO/NRG DVD" 'mp01
    
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
    If cmbTipoFile.ListIndex = 5 Then TP = "Imagenes ISO/Nero CD"
    If cmbTipoFile.ListIndex = 6 Then TP = "Videos 3GP"
    If cmbTipoFile.ListIndex = 7 Then TP = "Temas para movil"
    If cmbTipoFile.ListIndex = 8 Then TP = "Imagenes ISO/Nero DVD" 'mp01
    
    If vsCantFileCarrito.Value = 0 Or vsPrecioTotalCarrito.Value = 0 Then
        lblDetPrec.Caption = "Paquete no definido" 'mp01
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

