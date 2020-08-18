VERSION 5.00
Begin VB.UserControl txtRolling 
   BackColor       =   &H00000080&
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   315
   ScaleWidth      =   4800
   Begin VB.Timer Reloj 
      Left            =   2340
      Top             =   150
   End
   Begin VB.Label lblROLL2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "txtRolling2"
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
      Height          =   195
      Left            =   1140
      TabIndex        =   1
      Top             =   60
      Width           =   1050
   End
   Begin VB.Label lblROLL1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "txtRolling"
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
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   930
   End
End
Attribute VB_Name = "txtRolling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_txtToRoll As String

Public Property Let txtToRoll(newTXT As String)
    m_txtToRoll = newTXT
End Property

Public Sub ToRoll()
    lblROLL1 = m_txtToRoll
    lblROLL2 = m_txtToRoll
    UserControl_Resize
    Reloj.Interval = 250 'dos veces por segundo
End Sub

Public Sub ToStop()
    Reloj.Interval = 0
End Sub

Public Sub Clear()
    lblROLL1 = ""
    lblROLL2 = ""
End Sub

Private Sub Reloj_Timer()
    lblROLL1.Left = lblROLL1.Left - 75
    lblROLL2.Left = lblROLL2.Left - 75
    If lblROLL1.Left < 0 Then lblROLL2.Left = lblROLL1.Left + lblROLL1.Width
    If lblROLL2.Left < 0 Then lblROLL1.Left = lblROLL2.Left + lblROLL2.Width
End Sub

Private Sub UserControl_Resize()
    'reubicar los controles
    'el uno visible...
    'no debo definir he y wi ya que son autosize
    lblROLL1.Top = 0
    lblROLL1.Left = 0
    ' y el dos justo por detras
    lblROLL2.Top = lblROLL1.Top
    lblROLL2.Left = lblROLL1.Left + lblROLL1.Width
End Sub
