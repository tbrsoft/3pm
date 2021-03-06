VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.UserControl tbrFullProc 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin tbrFaroButton.fBoton fBoton4 
      Height          =   405
      Left            =   720
      TabIndex        =   3
      Top             =   3180
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   12632256
      fCapt           =   "Boton"
      fEnabled        =   -1  'True
      fFontN          =   "Trebuchet MS"
      fFontS          =   14
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton3 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   2820
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   503
      fFColor         =   16777215
      fBColor         =   16777215
      fCapt           =   ""
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   8
      fECol           =   0
   End
   Begin tbrFaroButton.fBoton fBoton2 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   2820
      Visible         =   0   'False
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   503
      fFColor         =   16777215
      fBColor         =   12632256
      fCapt           =   ""
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   8
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   2115
      Left            =   720
      TabIndex        =   0
      Top             =   690
      Visible         =   0   'False
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   3731
      fFColor         =   16777215
      fBColor         =   12632256
      fCapt           =   "Detectando dispositivos ...."
      fEnabled        =   -1  'True
      fFontN          =   "Trebuchet MS"
      fFontS          =   26
      fECol           =   5452834
   End
End
Attribute VB_Name = "tbrFullProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub ShowWait(T As String, _
    Optional xMiliSegSalir As Long = 0, _
    Optional sPorc As Single = 0, _
    Optional Info2 As String)
    'ver que el credito alcance y que el tama�o disponible tambien
    
    'si le pone "" es para sacarlo. Para casos que no es por tiempo si no para esperar otro proceso
    If T = "" Then
        Extender.Visible = False
        Exit Sub
    End If
    
    fBoton4.Font.Size = 14
    fBoton4.Font.Bold = True
    
    Extender.Top = 0
    Extender.Left = 0
    'PutMe_X_Y 0, 0
    
    UserControl.Width = Parent.Width
    UserControl.Height = Parent.Height
    
    fBoton1.Top = 0
    fBoton1.Left = 0
    fBoton1.Width = UserControl.Width
    fBoton1.Height = UserControl.Height
    fBoton1.Caption = T
    fBoton1.Visible = True
    
    fBoton2.Width = fBoton1.Width / 2
    fBoton2.Left = fBoton1.Width / 2 - fBoton2.Width / 2
    fBoton2.Top = fBoton1.Top + fBoton1.Height - (4 * fBoton2.Height)
    fBoton2.Visible = True
    fBoton2.ZOrder
    
    fBoton4.Width = fBoton2.Width
    fBoton4.Left = fBoton2.Left
    fBoton4.Top = fBoton2.Top + fBoton2.Height + 30
    fBoton4.Caption = Replace(Info2, vbCrLf, " / ") 'set08 no entran 2 renglones
    fBoton4.Caption = Replace(Info2, vbCr, " / ") 'set08 no entran 2 renglones
    fBoton4.Caption = Replace(Info2, vbLf, " / ") 'set08 no entran 2 renglones
    fBoton4.Visible = (Info2 <> "")
    
    Extender.Visible = True
    Extender.ZOrder
    'PutMeTop
    
    If sPorc > 0 Then ShPorc sPorc
    
    UserControl.Refresh

    If xMiliSegSalir <> 0 Then
        
        'si es negativo son los mismo milisegundos pero no muestro el segundero feo
        Dim ShowSecBack As Boolean
        ShowSecBack = (xMiliSegSalir > 0)
        xMiliSegSalir = Abs(xMiliSegSalir)
        
        Dim H As Single
        H = Timer
        Dim SFalta As Long, LastS As Long
        LastS = 10
        Do
            DoEvents 'agregado set 08 !!!!
            SFalta = CLng((H + (xMiliSegSalir / 1000)) - Timer)
            If LastS <> SFalta Then
                Extender.ZOrder
                'PutMeTop
                If ShowSecBack Then
                    fBoton1.Caption = T + vbCrLf + "(" + CStr(SFalta + 1) + ")"
                End If
                ShPorc ((SFalta * 1000) / xMiliSegSalir) * 100
                'Extender.Refresh
                UserControl.Refresh
            End If
            LastS = SFalta
            If Timer > (H + (xMiliSegSalir / 1000)) Then Exit Do
        Loop
        Extender.Visible = False
    End If
    
End Sub

Private Sub ShPorc(X As Single)
    
    X = Abs(X) '!!!!???? me llego negativo !!!???
'    If x = 0 Then
'        fBoton2.Visible = False
'        fBoton3.Visible = False
'        Exit Sub
'    End If

    fBoton3.Top = fBoton2.Top
    fBoton3.Left = fBoton2.Left
    fBoton3.Width = CLng(CSng((X / 100) * fBoton2.Width))
    
    fBoton3.Visible = True
    
    fBoton2.ZOrder
    fBoton3.ZOrder
    
    'Extender.Refresh
    UserControl.Refresh
End Sub

