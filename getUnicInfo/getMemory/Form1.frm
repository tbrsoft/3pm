VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reemplazo de codigo 3PM"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6465
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6465
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
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   750
      Width           =   3105
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   1185
      Left            =   420
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   5715
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   525
      Left            =   3600
      TabIndex        =   3
      Top             =   750
      Width           =   2505
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
      Height          =   495
      Left            =   420
      TabIndex        =   2
      Top             =   180
      Width           =   5685
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub GetMem1 Lib "msvbvm60.dll" (ByVal MemAddress As Long, var As Byte)

Private Function GetBIOSDate() As String
  Dim p As Byte, MemAddr As Long, sBios As String
  Dim i As Integer
  'start of bios serial number ?&HFE0C0
  MemAddr = &HFE000
  For i = 0 To 331
      Call GetMem1(MemAddr + i, p)
      'get printable characters
      If p > 31 And p <= 128 Then
      sBios = sBios & Chr$(p)
    End If
  Next i
  GetBIOSDate = sBios
End Function

Private Function SumaCHRtxt(TXT As String) As Long
    
    'sumar el valor CHR de los caracteres de un texto
    Dim Caracter As String
    Dim TMP As Long
    
    For J = 1 To Len(TXT)
      Caracter = Mid(TXT, J, 1)
      TMP = TMP + Asc(Caracter)
    Next J
    SumaCHRtxt = TMP
End Function
Private Sub Command1_Click()
    Text1 = GetBIOSDate
    Label3 = SumaCHRtxt(GetBIOSDate)
End Sub

