VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reemplazo de codigo 3PM"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6660
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Generar Codigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   3105
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   6240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0442
      Top             =   540
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Presione el boton GENERAR CODIGO para obtener una identificacion del equipo que esta usando"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   705
      Left            =   420
      TabIndex        =   3
      Top             =   180
      Width           =   5685
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   3780
      TabIndex        =   2
      Top             =   990
      Width           =   2385
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub GetMem1 Lib "msvbvm60.dll" (ByVal MemAddress As Long, var As Byte)
Dim SumaChar As Long

Private Function GetBIOSDate() As String
  SumaChar = 0
  Dim p As Byte, MemAddr As Long, sBios As String
  Dim i As Integer
  'start of bios serial number ?&HFE0C0
  MemAddr = &HFE000
  For i = 0 To 331
      Call GetMem1(MemAddr + i, p)
      'get printable characters
      If p > 31 And p <= 128 Then
      sBios = sBios & Chr$(p)
      SumaChar = SumaChar + p
    End If
  Next i
  GetBIOSDate = sBios
End Function

Private Sub Command1_Click()
    Text1 = GetBIOSDate
    Label1 = SumaChar
End Sub

