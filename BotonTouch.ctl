VERSION 5.00
Begin VB.UserControl BotonTouch 
   BackColor       =   &H00000000&
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1620
   LockControls    =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   1620
   Begin VB.Image imgTouch 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   60
      Top             =   45
      Width           =   1500
   End
End
Attribute VB_Name = "BotonTouch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Imagen As Image
Private m_ImagenDown As Image
Event CLICK()

'las propiedades de imágenes se toman si o si de un control Image
Public Property Set Imagen(imgNew As Image)
    Set m_Imagen = imgNew
    imgTouch.Picture = imgNew.Picture
End Property

Public Property Set ImagenDown(imgNew As Image)
    Set m_ImagenDown = imgNew
End Property

Private Sub imgTouch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgTouch.Picture = m_ImagenDown.Picture
    RaiseEvent CLICK
End Sub

Private Sub imgTouch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgTouch.Picture = m_Imagen.Picture
End Sub

Public Sub Flash()
    'destello para simular pulsacion
    imgTouch.Picture = m_ImagenDown.Picture
    imgTouch.Refresh
    imgTouch.Picture = LoadPicture
    imgTouch.Refresh
    imgTouch.Picture = m_ImagenDown.Picture
    imgTouch.Refresh
    imgTouch.Picture = LoadPicture
    imgTouch.Refresh
    imgTouch.Picture = m_Imagen.Picture
End Sub
Private Sub UserControl_Resize()
    'imgTouch.Visible = False
    Dim Margen As Long
    Margen = 10
    imgTouch.Stretch = True
    imgTouch.Width = UserControl.Width - Margen * 2
    imgTouch.Height = UserControl.Height - Margen * 2
    imgTouch.Left = Margen
    imgTouch.Top = Margen
    'imgTouch.Visible = True
End Sub
