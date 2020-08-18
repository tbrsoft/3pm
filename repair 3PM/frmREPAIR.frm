VERSION 5.00
Begin VB.Form frmREPAIR 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reparar 3PM"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5820
   Icon            =   "frmREPAIR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "CERRAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1560
      TabIndex        =   4
      Top             =   2640
      Width           =   2745
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reparar 3PM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1590
      TabIndex        =   1
      Top             =   1440
      Width           =   2745
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   4
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   1950
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   3
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   1950
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   2
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   1950
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   1
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   1950
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   0
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   1950
      Width           =   225
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmREPAIR.frx":0442
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1125
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   5655
   End
   Begin VB.Label lblPBAR 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   2160
      Width           =   15
   End
   Begin VB.Label lblBAR 
      BackColor       =   &H000000C0&
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   2160
      Width           =   5685
   End
End
Attribute VB_Name = "frmREPAIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WinFolder As String
Dim SysFolder As String
Dim FSO As New Scripting.FileSystemObject

Private Sub Command1_Click()
    'borrar todos los archivos que no se instalaron y que forman parte de 3PM
    
    'primero RMLVF.DLL (indicador de licencia)
    tbrDelete SysFolder + "\rmlvf.dll", 40, 0
    tbrDelete SysFolder + "\rmlvf.tlb", 55, 1
    'primero nnr.dll de win y de system (indicador de creditos)
    tbrDelete WinFolder + "\nnr.dll", 70, 2
    tbrDelete SysFolder + "\nnr.dll", 85, 3
    'archivos de usus demo
    tbrDelete WinFolder + "\slx98.dll", 100, 4
    MsgBox "Se ha terminado la reparacion"
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    SysFolder = FSO.GetSpecialFolder(SystemFolder)
    WinFolder = FSO.GetSpecialFolder(WindowsFolder)
    
End Sub

Public Function tbrDelete(Arch As String, PorcPasado As Long, IndiceBola As Integer) As Boolean
    If FSO.FileExists(Arch) Then
        FSO.DeleteFile Arch, True
        OKdelete(IndiceBola).BackColor = vbGreen
        tbrDelete = True
    Else
        OKdelete(IndiceBola).BackColor = vbRed
        tbrDelete = False
    End If
    lblPBAR.Width = lblBAR.Width * PorcPasado / 100
End Function
