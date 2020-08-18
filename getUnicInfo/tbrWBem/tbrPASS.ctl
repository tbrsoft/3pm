VERSION 5.00
Begin VB.UserControl tbrPASS 
   BackColor       =   &H000040C0&
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   ScaleHeight     =   1860
   ScaleWidth      =   5805
   Begin VB.Label TTT 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1920
      TabIndex        =   7
      Top             =   180
      Width           =   2595
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1920
      TabIndex        =   6
      Top             =   510
      Width           =   2595
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   2595
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "WBmem Procesor"
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
      Height          =   225
      Left            =   150
      TabIndex        =   4
      Top             =   210
      Width           =   1755
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "GetInfoProcesor"
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
      Height          =   225
      Left            =   150
      TabIndex        =   3
      Top             =   540
      Width           =   1755
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "BiosDate Mother"
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
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   870
      Width           =   1755
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "FINAL"
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
      Height          =   225
      Left            =   2550
      TabIndex        =   1
      Top             =   1230
      Width           =   675
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   300
      Left            =   150
      TabIndex        =   0
      Top             =   1410
      Width           =   5565
   End
End
Attribute VB_Name = "tbrPASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Initialize()
    Label1 = GetInfoProc
    Label2 = GetBiosNumber
    TTT = GetProcID
    Label7 = TTT + "." + Hex(Label1) + "." + Hex(Label2)
    
End Sub
