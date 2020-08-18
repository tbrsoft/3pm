VERSION 5.00
Begin VB.Form frmConfigLPT 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Puerto"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLPTPORT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   2
      Left            =   2880
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "378"
      Top             =   510
      Width           =   1350
   End
   Begin VB.TextBox txtLPTPORT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   1
      Left            =   1500
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "378"
      Top             =   510
      Width           =   1350
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   435
      Left            =   2190
      TabIndex        =   3
      Top             =   1050
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   435
      Left            =   1230
      TabIndex        =   2
      Top             =   1050
      Width           =   855
   End
   Begin VB.TextBox txtLPTPORT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "378"
      Top             =   510
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Direcciones del puerto paralelo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   52
      Left            =   600
      TabIndex        =   1
      Top             =   150
      Width           =   3435
   End
End
Attribute VB_Name = "frmConfigLPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ChangeConfig "LptPort0", txtLPTPORT(0).tExt
    ChangeConfig "LptPort1", txtLPTPORT(1).tExt
    ChangeConfig "LptPort2", txtLPTPORT(2).tExt
    
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtLPTPORT(0).tExt = LeerConfig("LptPort0", "378")
    txtLPTPORT(1).tExt = LeerConfig("LptPort1", "379")
    txtLPTPORT(2).tExt = LeerConfig("LptPort2", "37A")
End Sub
