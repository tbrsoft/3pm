VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reemplazo de codigo 3PM"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6465
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4620
      TabIndex        =   9
      Text            =   "1040384"
      Top             =   1050
      Width           =   1005
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   3480
      TabIndex        =   7
      Text            =   "&HFE000"
      Top             =   1050
      Width           =   1005
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   4740
      TabIndex        =   6
      Text            =   "331"
      Top             =   660
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   3510
      TabIndex        =   4
      Text            =   "0"
      Top             =   660
      Width           =   555
   End
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
      Top             =   1830
      Width           =   3105
   End
   Begin VB.TextBox Text1 
      Height          =   3465
      Left            =   450
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2400
      Width           =   5715
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "En la posicion de memoria"
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
      Height          =   225
      Index           =   2
      Left            =   750
      TabIndex        =   8
      Top             =   1080
      Width           =   2715
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Desde la posicion de memoria            hasta"
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
      Height          =   225
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   690
      Width           =   4965
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
      Left            =   3630
      TabIndex        =   3
      Top             =   1830
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
      Index           =   0
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
  sBios = ""
  p = 0
  'start of bios serial number ?&HFE0C0
  MemAddr = CLng(Text5)
  'For i = 0 To 331
  For i = CLng(Tex2) To CLng(Text3)
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

Private Sub Text4_Change()
    Text5 = HEXtoLONG(Text4)
End Sub

Private Function HEXtoLONG(N As String)
    'recibe el hex en str y devuelve un numero en str
    
    Dim Letra As String
    Dim C As Long
    Dim NumeroActual As Long
    Dim ACUM ' As Double
    For C = 1 To Len(N)
        Letra = Mid(N, C, 1)
        Select Case Letra
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                NumeroActual = Val(Letra)
            Case "A"
                NumeroActual = 10
            Case "B"
                NumeroActual = 11
            Case "C"
                NumeroActual = 12
            Case "D"
                NumeroActual = 13
            Case "E"
                NumeroActual = 14
            Case "F"
                NumeroActual = 15
        End Select
        Dim ToSum ' As Double
        ToSum = NumeroActual * (16 ^ (Len(N) - C))
        ACUM = ACUM + ToSum
        'Label10 = Label10 + "LETRA: " + Letra + "=" + CStr(ToSum) + vbCrLf
        
    Next
    
    HEXtoLONG = CStr(ACUM)
End Function
