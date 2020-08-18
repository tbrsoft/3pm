VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Form2"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   LinkTopic       =   "Form2"
   ScaleHeight     =   4545
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Detener!!!"
      Height          =   525
      Left            =   180
      TabIndex        =   7
      Top             =   3120
      Width           =   1965
   End
   Begin VB.TextBox Text4 
      Height          =   465
      Left            =   3150
      TabIndex        =   6
      Text            =   "0,3"
      Top             =   600
      Width           =   825
   End
   Begin VB.TextBox Text3 
      Height          =   465
      Left            =   2250
      TabIndex        =   5
      Text            =   "100"
      Top             =   570
      Width           =   825
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Left            =   180
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.TextBox Text2 
      Height          =   4425
      Left            =   4620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   60
      Width           =   5445
   End
   Begin VB.CommandButton Command7 
      Caption         =   "PRUEBA FULL"
      Height          =   885
      Left            =   210
      TabIndex        =   1
      Top             =   330
      Width           =   1965
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   180
      TabIndex        =   2
      Top             =   2160
      Width           =   4065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   180
      TabIndex        =   0
      Top             =   2610
      Width           =   2835
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Detener As Boolean

Private Sub Command7_Click()
    
    S3.HwndMsg = Text1.hWnd
    esperar 1
    
    S3.Prender
    esperar 1
    
    'pongo todo en cero
    S3.ReIniContLuis
    esperar 1
    
    Label2.Caption = S3.GetnPlaca(SF + "prec.dll")
    
    Detener = False
    
    Dim J As Long
    For J = 1 To CLng(Text3)
        S3.AddCont J Mod 4
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

Private Sub esperar(N As Single)
    N = Timer + N
    Do While Timer < N
        DoEvents
    Loop
End Sub

