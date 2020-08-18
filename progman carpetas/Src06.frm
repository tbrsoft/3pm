VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cargar 3PM al inicio"
   ClientHeight    =   3975
   ClientLeft      =   1080
   ClientTop       =   1425
   ClientWidth     =   2625
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Src06.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3975
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Crear inicio de 3PM"
      Height          =   435
      Left            =   210
      TabIndex        =   1
      Top             =   2220
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   480
      Picture         =   "Src06.frx":014A
      Top             =   150
      Width           =   1650
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Una vez terminado reinicie el sistema. Esto se hara esta única vez"
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
      Height          =   1125
      Left            =   210
      TabIndex        =   2
      Top             =   2790
      Width           =   2205
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   2160
      TabIndex        =   0
      Top             =   1110
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Dim WinFolder As String
Dim FSO As New Scripting.FileSystemObject

Private Sub Form_Load()
    WinFolder = FSO.GetSpecialFolder(WindowsFolder)
End Sub

Private Sub Command1_Click()
    Dim FileNameSHORT As String
    'transformar el acceso a ShortPath
    Dim TMP As String * 255
    Dim lenShort As Long
    lenShort = GetShortPathName(App.Path + "\3pm.exe", TMP, 255)
    FileNameSHORT = Left$(TMP, lenShort)
    
    CreateProgManGroup Me, "Inicio", WinFolder
    CreateProgManItem Me, FileNameSHORT, "3PM"
    'para esto ultimo el ejecutable debe estar en la carpeta de 3PM
    
    MsgBox "Se crearon las carpetas corectamente. Reinicie el equipo"
    End
End Sub

Private Sub CreateProgManGroup(x As Form, GroupName$, GroupPath$)
    Dim i%, z%                'Declare required working variables
    Screen.MousePointer = 11  'hourglass mousepointer while working
    On Error Resume Next      'Not good to have program crash :-)
    ' Set LinkTopic & LinkMode parameters
    x.Label1.LinkTopic = "ProgMan|Progman"
    x.Label1.LinkMode = 2
    For i% = 1 To 10          ' Give the DDE process time to take place
      z% = DoEvents()
    Next
    x.Label1.LinkTimeout = 100
    ' Actually create the group now
    x.Label1.LinkExecute "[CreateGroup(" + GroupName$ + Chr$(44) + GroupPath$ + ")]"
    ' Reset label properties and mousepointer
    x.Label1.LinkTimeout = 50
    x.Label1.LinkMode = 0
    Screen.MousePointer = 0
End Sub

Private Sub CreateProgManItem(x As Form, CmdLine$, IconTitle$)
    Dim i%, z%                'Declare required working variables
    Screen.MousePointer = 11  'hourglass mousepointer while working
    On Error Resume Next      'Not good to have program crash :-)
    ' Set LinkTopic & LinkMode parameters
    x.Label1.LinkTopic = "ProgMan|Progman"
    x.Label1.LinkMode = 2
    For i% = 1 To 10          ' Give the DDE process time to take place
      z% = DoEvents()
    Next
    x.Label1.LinkTimeout = 100
    x.Label1.LinkExecute "[AddItem(" + CmdLine$ + Chr$(44) + IconTitle$ + Chr$(44) + ",,)]"
    ' Reset label properties and mousepointer
    x.Label1.LinkTimeout = 50
    x.Label1.LinkMode = 0
    Screen.MousePointer = 0
End Sub

