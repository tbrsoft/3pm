VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Testeando teclado de tbrSoft"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11295
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "cero conts"
      Height          =   375
      Left            =   8100
      TabIndex        =   8
      Top             =   60
      Width           =   1965
   End
   Begin VB.CommandButton Command5 
      Caption         =   "LIC calculo azar"
      Height          =   375
      Left            =   6030
      TabIndex        =   7
      Top             =   30
      Width           =   1965
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LIC numero placa"
      Height          =   375
      Left            =   3990
      TabIndex        =   6
      Top             =   30
      Width           =   1965
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Apagar T2"
      Height          =   315
      Left            =   2610
      TabIndex        =   5
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Prender T2"
      Height          =   315
      Left            =   1290
      TabIndex        =   4
      Top             =   60
      Width           =   1275
   End
   Begin VB.ListBox List2 
      Columns         =   5
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      ItemData        =   "Form1.frx":0442
      Left            =   90
      List            =   "Form1.frx":0449
      TabIndex        =   3
      Top             =   510
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Prender"
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   6450
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.ListBox List1 
      Columns         =   4
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      Left            =   1500
      TabIndex        =   1
      Top             =   540
      Width           =   9255
   End
   Begin VB.Timer Timer1 
      Left            =   10230
      Top             =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S3 As New tbrSKS3.clsTbrSKS3

Private Sub Command1_Click()
    S3.ToTimer2 True
End Sub

Private Sub Command2_Click()
    S3.HwndMsg = Text1.hWnd
    S3.ReIniCounters
    S3.Prender
End Sub

Private Sub Command3_Click()
    S3.ToTimer2 False
End Sub

Private Sub Command4_Click()
    Dim SF As String
    Dim FSo As New Scripting.FileSystemObject
    SF = FSo.GetSpecialFolder(SystemFolder)
    If Right(SF, 1) <> "\" Then SF = SF + "\"
    
    S3.GetnPlaca SF + "prec.dll"
    
End Sub

Private Sub Command5_Click()
    Dim J As Long
    Randomize
    J = CLng(Rnd * 3)
    S3.AddCont J
End Sub

Private Sub Command6_Click()
    S3.ReIniContLuis
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    
    List1.Left = List2.Left + List2.Width + 50
    List1.Width = Me.Width - List2.Width - 400 - List2.Left
    List1.Height = Me.Height - List1.Top - 600
    List2.Left = 30
    List2.Top = List1.Top
    List2.Height = List1.Height
End Sub

Private Sub List1_DblClick()
    List1.Clear
End Sub

Private Sub Text1_Change()
    
    Dim txtS3 As String
    txtS3 = Text1.Text

    If txtS3 = "" Then Exit Sub
    
    List1.AddItem CStr(Timer) + "  " + txtS3
    List1.ListIndex = List1.ListCount - 1
    
    Dim P As String
    P = txtS3
    
    Dim SP() As String
    SP = Split(P, ":")
    
    If SP(0) = "sD" Then
        'actualziar los contadores
        Dim I As Long: List2.Clear
        Dim GC As Long
        For I = 1 To 23
            GC = S3.GetCounter(I)
            List2.AddItem String(2 - Len(CStr(I)), "0") + CStr(I) + ": " + _
                String(4 - Len(GC), "0") + CStr(GC)
        Next I
       
    End If
    
    'vaciarlos !!!
    Text1.Text = ""
    
End Sub
