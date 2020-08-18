VERSION 5.00
Begin VB.Form frmREPAIR 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reparar 3PM"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5820
   Icon            =   "frmREPAIR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   2745
   End
   Begin VB.Label Label3 
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
      Height          =   1125
      Left            =   120
      TabIndex        =   4
      Top             =   150
      Width           =   5655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   90
      TabIndex        =   3
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label lblBAR 
      BackColor       =   &H000000C0&
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   2160
      Width           =   5685
   End
   Begin VB.Label Label1 
      Caption         =   "Sin tareas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1950
      Width           =   5655
   End
End
Attribute VB_Name = "frmREPAIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WinFolder As String
Dim SystemFolder As String
Dim FSO As New Scripting.FileSystemObject

Private Sub Form_Load()

    SYSfolder = FSO.GetSpecialFolder(SystemFolder)
    WinFolder = FSO.GetSpecialFolder(WindowsFolder)
    
End Sub
