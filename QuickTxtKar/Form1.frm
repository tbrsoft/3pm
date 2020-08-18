VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0013
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Texto As New LeerKar
Private Sub List1_Click()
Select Case List1.ListIndex
 Case 0
  Texto.OpenDevice "1.kar"
  Text1.Text = Texto.TextLyric

 Case 1
  Texto.OpenDevice "2.kar"
  Text1.Text = Texto.TextLyric

 Case 2
  Texto.OpenDevice "3.kar"
  Text1.Text = Texto.TextLyric

 Case 3
  Texto.OpenDevice "4.kar"
  Text1.Text = Texto.TextLyric

 Case 4
  Texto.OpenDevice "5.kar"
  Text1.Text = Texto.TextLyric

End Select
End Sub
