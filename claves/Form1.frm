VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CLaves III Edicion"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   6180
      Width           =   8200
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3960
      Width           =   8200
   End
   Begin VB.TextBox tAsig 
      Height          =   6705
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1320
      Width           =   1515
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1620
      TabIndex        =   3
      Top             =   840
      Width           =   6345
   End
   Begin VB.TextBox lstClaves 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1770
      Width           =   8200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar Clave"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8190
      TabIndex        =   1
      Top             =   1350
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   9855
   End
   Begin VB.Label Label3 
      Caption         =   "SuperLicencia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1620
      TabIndex        =   9
      Top             =   5790
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "FULL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   3600
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "GRATUITA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1650
      TabIndex        =   7
      Top             =   1410
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim K As New clsKEYS

Private Sub Command1_Click()
    'lstcLAVES.Clear
    Dim A As Long
    'For A = 1 To 5
    '    lstClaves = lstClaves + vbCrLf + "Sin Cargar(" + CStr(A) + "): " + K.CLAVE(aSinCargar, A, Text1)
    'Next
    'For A = 1 To 5
    '    lstClaves = lstClaves + vbCrLf + "Erronea(" + CStr(A) + "): " + K.CLAVE(BErronea, A, Text1)
    'Next
    For A = 1 To 50
        lstClaves = lstClaves + "Gratuita(" + CStr(A) + "): " + K.CLAVE(CGratuita, A, Text1) + vbCrLf
    Next
    tAsig = K.Asignaciones 'del 50 de gratuita
    For A = 1 To 50
        lstClaves = lstClaves + vbCrLf + "Minima(" + CStr(A) + "): " + K.CLAVE(DMinima, A, Text1)
    Next
    For A = 1 To 50
        lstClaves = lstClaves + vbCrLf + "Comun(" + CStr(A) + "): " + K.CLAVE(EComun, A, Text1)
    Next
    For A = 1 To 50
        lstClaves = lstClaves + vbCrLf + "Premium(" + CStr(A) + "): " + K.CLAVE(FPremium, A, Text1)
    Next
    For A = 1 To 50
        Text3 = Text3 + "Full(" + CStr(A) + "): " + K.CLAVE(GFull, A, Text1) + vbCrLf
    Next
    For A = 1 To 50
        Text4 = Text4 + "SuperLicencia(" + CStr(A) + "): " + K.CLAVE(HSuperLicencia, A, Text1) + vbCrLf
    Next
    
    lstClaves = lstClaves + vbCrLf
    
End Sub

Private Sub Form_Load()
    K.ClaveDLL = "ashjdklahsJKLHASL65456456456"
    Text1 = K.UniquePC
End Sub

Private Sub Text1_Change()
    'al cambiar que se vea a que numero de la clave anterior corresponde
    'Text6 = K.UniquePCOLD
    'siempre mayusculas!!!!!!!!!!
    'ya tuve quilombo con tomas porque copiaba de una PC a otra el codifp
    'no tenia internet en la pcx de la fonola
    Text6 = K.GetOldFromNew(UCase(Text1.Text))
End Sub
