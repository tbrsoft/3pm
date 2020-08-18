VERSION 5.00
Begin VB.Form frmCrearSKIN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tCS 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2415
      Left            =   270
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmCerarSKIN.frx":0000
      Top             =   180
      Width           =   5835
   End
End
Attribute VB_Name = "frmCrearSKIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If fso.FileExists(AP + "crearskin.txt") Then
        Dim a As TextStream
        Set a = fso.OpenTextFile(AP + "crearskin.txt", ForReading, False)
        tCS.Text = a.ReadAll
    Else
        tCS.Text = TR.Trad("No se encuentra el archivo explicativo%99%")
    End If
End Sub
