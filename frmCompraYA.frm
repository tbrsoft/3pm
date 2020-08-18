VERSION 5.00
Begin VB.Form frmCompraYA 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compra inmediata!"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "frmCompraYA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDIFF 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1155
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "frmCompraYA.frx":014A
      Top             =   4980
      Width           =   8445
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3450
      TabIndex        =   6
      Top             =   6150
      Width           =   1785
   End
   Begin VB.TextBox LBL 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3195
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmCompraYA.frx":0179
      Top             =   90
      Width           =   8445
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   5730
      TabIndex        =   9
      Top             =   3600
      UseMnemonic     =   0   'False
      Width           =   2805
   End
   Begin VB.Label lblUPLIC 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Licencias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   5730
      TabIndex        =   8
      Top             =   3930
      UseMnemonic     =   0   'False
      Width           =   2805
   End
   Begin VB.Label lblPrecioSuperLIC 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Licencias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   2910
      TabIndex        =   5
      Top             =   3930
      UseMnemonic     =   0   'False
      Width           =   2805
   End
   Begin VB.Label lblPrecioLIC 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Licencias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   60
      TabIndex        =   4
      Top             =   3930
      UseMnemonic     =   0   'False
      Width           =   2805
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SUPERLICENCIAS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   2910
      TabIndex        =   3
      Top             =   3600
      UseMnemonic     =   0   'False
      Width           =   2805
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Licencias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   3600
      UseMnemonic     =   0   'False
      Width           =   2805
   End
   Begin VB.Label lblDisco 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precios de 3PM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   3300
      UseMnemonic     =   0   'False
      Width           =   8475
   End
End
Attribute VB_Name = "frmCompraYA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LBL = "Para comprar ahora su licencia de 3PM debera hacer una transferencia por WESTERN UNION " + _
    "por el monto que corresponda a:" + vbCrLf + vbCrLf + _
    "Destinatario: Andrés Vázquez Flexes" + vbCrLf + _
    "Domicilio: Calle 4 n° 16" + vbCrLf + _
    "Ciudad: Mendiolaza" + vbCrLf + _
    "Estado/Provincia: Córdoba" + vbCrLf + _
    "Pais: Argentina" + vbCrLf + _
    "Telefono: 54-3543-485045" + vbCrLf + _
    "Celular: 54-9-351-4022170" + vbCrLf + _
    "DNI: 26.453.653 (Documento Nacional de Identidad)" + vbCrLf + vbCrLf + _
    "Luego envie un email a info@tbrsoft y a tbrsoft@cpcipc.org con el código de la " + _
    "transferencia y el codigo de su/s equipo/s. Recibira inmediatamente via email " + _
    "la clave para habilitar su equipo. Si adquiere más de una licencia quedará " + _
    "habilitado para solicitar una por una las claves para los " + _
    "distintos equipos sin fecha de vencimiento alguna." + vbCrLf + _
    "Si desea además recibir el CD de instalación y el manual de uso en su domicilio" + _
    " (esto no es necesario) consulte por el costo de envio desde Argentina hasta su domicilio."
    
    lblPrecioLIC = "1 Licencia = U$S 75" + vbCrLf + _
        "2 Licencias = U$S 115" + vbCrLf + _
        "3 Licencias = U$S 155" + vbCrLf + _
        "5 Licencias = U$S 200"
    
    lblPrecioSuperLIC = "1 Licencia = U$S 145" + vbCrLf + _
        "2 Licencias = U$S 245" + vbCrLf + _
        "3 Licencias = U$S 345" + vbCrLf + _
        "5 Licencias = U$S 400"
    
    lblUPLIC = "El costo para actualizar una licencia común a SUPERLICENCIA es de U$S 100"
        
    txtDIFF = "Diferencias entre las licencias y las superlicencias:" + vbCrLf + _
    "La funcionalidad en los dos casos es exactamente igual. El beneficio de la SUPERLICENCIA es que" + _
    " permite modificar las imágenes y textos del software para quitar los logos y textos de tbrSoft y" + _
    " colocar imágenes y textos personalizados. Esto será importante solo si usted es operador de " + _
    "fonolas y desea imponer su propia marca." + vbCrLf + _
    "Se usa para vender o rentar equipos armados. La instalación seguira usando los logos originales" + _
    " y cada equipo deberá modificarse manualmente. El proceso dura solo 5 minutos"
        
End Sub

