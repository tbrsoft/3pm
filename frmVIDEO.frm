VERSION 5.00
Begin VB.Form frmVIDEO 
   BackColor       =   &H00400000&
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
   Begin VB.PictureBox picVideo 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
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
      BackColor       =   &H00C0FFFF&
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

