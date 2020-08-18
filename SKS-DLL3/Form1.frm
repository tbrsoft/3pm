VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Testeando teclado de tbrSoft"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "Detener!!!"
      Height          =   315
      Left            =   7530
      TabIndex        =   13
      Top             =   450
      Width           =   825
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   5790
      TabIndex        =   12
      Text            =   "100"
      Top             =   450
      Width           =   825
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6660
      TabIndex        =   11
      Text            =   "0,3"
      Top             =   480
      Width           =   825
   End
   Begin VB.TextBox Text2 
      Height          =   4425
      Left            =   7950
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1530
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "F2"
      Height          =   315
      Left            =   5190
      TabIndex        =   9
      Top             =   450
      Width           =   585
   End
   Begin VB.CommandButton Command6 
      Caption         =   "cero conts"
      Height          =   315
      Left            =   2700
      TabIndex        =   8
      Top             =   780
      Width           =   1965
   End
   Begin VB.CommandButton Command5 
      Caption         =   "LIC calculo azar"
      Height          =   315
      Left            =   3210
      TabIndex        =   7
      Top             =   450
      Width           =   1965
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LIC numero placa"
      Height          =   315
      Left            =   1230
      TabIndex        =   6
      Top             =   450
      Width           =   1965
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Apagar T2"
      Height          =   315
      Left            =   1410
      TabIndex        =   5
      Top             =   780
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Prender T2"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   780
      Width           =   1275
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00404000&
      Columns         =   2
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
      Height          =   1260
      IntegralHeight  =   0   'False
      ItemData        =   "Form1.frx":0442
      Left            =   90
      List            =   "Form1.frx":0449
      TabIndex        =   3
      Top             =   1170
      Width           =   3165
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Prender"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7140
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2940
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List1 
      BackColor       =   &H0080FFFF&
      Columns         =   3
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   3300
      TabIndex        =   1
      Top             =   1170
      Width           =   7065
   End
   Begin VB.Timer Timer1 
      Left            =   8850
      Top             =   930
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Detener As Boolean

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



Private Sub Command7_Click()
    'Form2.Show 1
    
    Detener = False
    On Local Error Resume Next
    
    Dim J As Long
    For J = 1 To CLng(Text3)
        S3.AddCont (J Mod 4)
        Text2.Text = Text2.Text + "CONT:" + CStr(J) + " " + S3.GetResLicSTR + vbCrLf
        Text2.SelStart = Len(Text2) - 1
        Text2.Refresh
        
        Label1.Caption = Round(S3.GetPorcLic, 2)
        Label1.Refresh
        
        esperar CSng(Text4.Text)
        
        If Detener = True Then
            MsgBox "me detuve"
            Exit Sub
        End If
    Next J
    
End Sub

Private Sub Command8_Click()
    Detener = True
End Sub

Private Sub Form_Load()
    Dim FSO As New Scripting.FileSystemObject
    SF = FSO.GetSpecialFolder(SystemFolder)
    If Right(SF, 1) <> "\" Then SF = SF + "\"
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    

    List1.Width = Me.Width - List2.Width - 200 - List2.Left - Text2.Width
    List1.Height = Me.Height - List1.Top - 600
    List2.Left = 30
    List2.Top = List1.Top
    List2.Height = List1.Height
    List1.Left = List2.Left + List2.Width + 50
    
    Text2.Left = List1.Left + List1.Width
    Text2.Top = List1.Top
    Text2.Height = List1.Height
    
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

Private Sub esperar(N As Single)
    N = Timer + N
    Do While Timer < N
        DoEvents
    Loop
End Sub
