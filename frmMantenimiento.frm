VERSION 5.00
Begin VB.Form frmMantenimiento 
   BackColor       =   &H00404080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      Caption         =   "Recordarme de ingresar aqui..."
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
      Height          =   2115
      Left            =   3630
      TabIndex        =   4
      Top             =   900
      Width           =   4305
      Begin VB.OptionButton Option4 
         BackColor       =   &H00404080&
         Caption         =   "Nunca, ingresare cuando lo desee."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   8
         Top             =   1620
         Width           =   3500
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00404080&
         Caption         =   "Cada 30 días"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   7
         Top             =   1230
         Width           =   3500
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404080&
         Caption         =   "Cada 15 días"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   6
         Top             =   810
         Width           =   3500
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404080&
         Caption         =   "Una vez por semana"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   390
         Width           =   3500
      End
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00404080&
      Caption         =   "Limpiar el archivo de log (log.txt)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   350
      Left            =   3420
      TabIndex        =   2
      Top             =   3990
      Width           =   5775
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00404080&
      Caption         =   "Revisar tamaño de las tapas de los discos (tapa.jpg)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   350
      Left            =   3420
      TabIndex        =   1
      Top             =   3690
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Realizar mantenimiento ahora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3300
      TabIndex        =   0
      Top             =   4770
      Width           =   5085
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "3PM mantenimiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   60
      TabIndex        =   3
      Top             =   90
      Width           =   11895
   End
End
Attribute VB_Name = "frmMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
