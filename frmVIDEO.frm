VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmVIDEO 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picKAR_V 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   1710
      ScaleHeight     =   2985
      ScaleWidth      =   3555
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   3555
      Begin tbrFaroButton.fBoton lblWAIT_V 
         Height          =   495
         Left            =   2160
         TabIndex        =   9
         Top             =   870
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         fFColor         =   6553600
         fBColor         =   16761024
         fCapt           =   ""
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   16777215
      End
      Begin tbrFaroButton.fBoton fPRC2_V 
         Height          =   250
         Left            =   420
         TabIndex        =   8
         Top             =   1530
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   450
         fFColor         =   6553600
         fBColor         =   8388608
         fCapt           =   ""
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   16777215
      End
      Begin tbrFaroButton.fBoton fPRC1_V 
         Height          =   250
         Left            =   450
         TabIndex        =   7
         Top             =   1530
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         fFColor         =   6553600
         fBColor         =   16761024
         fCapt           =   ""
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   16777215
      End
      Begin VB.Label lblTimeK_V 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   30
         TabIndex        =   4
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label LF1_V 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   600
         Left            =   30
         TabIndex        =   3
         Top             =   1110
         Width           =   705
      End
      Begin VB.Shape shKAR_V 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C0C0&
         Height          =   255
         Left            =   2100
         Shape           =   3  'Circle
         Top             =   180
         Width           =   255
      End
      Begin VB.Label lblTimeK2_V 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   60
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label LF2_V 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   600
         Left            =   60
         TabIndex        =   6
         Top             =   1140
         Width           =   705
      End
   End
   Begin VB.PictureBox picVideo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   4320
      ScaleHeight     =   975
      ScaleWidth      =   1455
      TabIndex        =   1
      Top             =   540
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picBigImg 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   2160
      ScaleHeight     =   975
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   570
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmVIDEO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Sub Form_Load()
    picKAR_V.AutoRedraw = True
End Sub

Private Sub picVideo_Resize()
    picKAR_V.Width = picVideo.Width
    picKAR_V.Height = picVideo.Height
    picKAR_V.Top = picVideo.Top
    picKAR_V.Left = picVideo.Left
End Sub

