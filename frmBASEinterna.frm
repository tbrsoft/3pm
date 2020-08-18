VERSION 5.00
Begin VB.Form frmBASEinterna 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Base interna"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInterprete 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   450
      TabIndex        =   10
      Top             =   1740
      Width           =   3735
   End
   Begin VB.TextBox txtNombreDisco 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   510
      TabIndex        =   9
      Top             =   2400
      Width           =   3675
   End
   Begin VB.TextBox txtIdDisco 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      TabIndex        =   8
      Top             =   1050
      Width           =   1275
   End
   Begin VB.ComboBox cmbAno 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmBASEinterna.frx":0000
      Left            =   1740
      List            =   "frmBASEinterna.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1080
      Width           =   885
   End
   Begin VB.ComboBox cmbEstilo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmBASEinterna.frx":0004
      Left            =   2700
      List            =   "frmBASEinterna.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   1485
   End
   Begin VB.TextBox txtCarpeta 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3090
      Width           =   3705
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   345
      Left            =   2310
      TabIndex        =   4
      Top             =   3630
      Width           =   1425
   End
   Begin VB.CommandButton cmdModif 
      Caption         =   "Modificar"
      Height          =   345
      Left            =   2310
      TabIndex        =   3
      Top             =   3990
      Width           =   1425
   End
   Begin VB.CommandButton cmdCarpeta 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2220
      TabIndex        =   2
      Top             =   2820
      Width           =   405
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Borrar todo"
      Height          =   345
      Left            =   870
      TabIndex        =   1
      Top             =   3630
      Width           =   1425
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   345
      Left            =   870
      TabIndex        =   0
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image TapaCD 
      Height          =   3300
      Left            =   5130
      Picture         =   "frmBASEinterna.frx":0008
      Stretch         =   -1  'True
      Top             =   300
      Width           =   3465
   End
   Begin VB.Label lblDisco 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "La renga - Despedazado por mil partes uy la nave del olvido"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   4890
      TabIndex        =   18
      Top             =   3630
      Width           =   3885
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Intérprete"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   420
      TabIndex        =   17
      Top             =   1500
      Width           =   1185
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Disco"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   16
      Top             =   2190
      Width           =   1515
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "IdDisco"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   15
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1740
      TabIndex        =   14
      Top             =   810
      Width           =   705
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Estilo"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   3150
      TabIndex        =   13
      Top             =   810
      Width           =   705
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "En Carpeta"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   510
      TabIndex        =   12
      Top             =   2850
      Width           =   1245
   End
   Begin VB.Label lblGDE 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Modificar Valores de Disco"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   525
      Index           =   8
      Left            =   750
      TabIndex        =   11
      Top             =   210
      Width           =   3135
   End
End
Attribute VB_Name = "frmBASEinterna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
