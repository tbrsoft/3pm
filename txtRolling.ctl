VERSION 5.00
Begin VB.UserControl txtRolling 
   BackColor       =   &H00000000&
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   1665
   ScaleWidth      =   4800
   Begin VB.Label lblROLL1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "txtRolling"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A9C8C9&
      Height          =   630
      Left            =   930
      TabIndex        =   0
      Top             =   510
      Width           =   2865
   End
End
Attribute VB_Name = "txtRolling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents TimerMR As tbrTimer.clsTimer 'ModoRoll4
Attribute TimerMR.VB_VarHelpID = -1

'estetica de los dos textos
Private mFont As New StdFont
Private mForeColor1() As OLE_COLOR 'uno distinto para cada elemento

Private DirColor As Integer
'es 1 si va de forecolor a backcolor (camino de ida)
'es 2 si va de backcolor a forecolor (camino de vuelta)

Private Textos() As String 'cola de datos a mostrar cero es siempre el actual
Private nTextoActual As Long

Dim tRGB As New tbrRGB
Private mInterval As Long
Private mVarColor As Byte

Private mMaxlargoRenglon As Long 'si esta en cero es libre

Public Property Let MaxlargoRenglon(MLR As Long)
    mMaxlargoRenglon = MLR
End Property

Public Property Get MaxlargoRenglon() As Long
    MaxlargoRenglon = mMaxlargoRenglon
End Property

Public Sub TextoACola(sTXT As String, lColor As OLE_COLOR)

    'ver que no sean renglones muy largos!
    sTXT = CortarTextos(sTXT)

    'ver si el el primero de los primeros
    Dim mUB As Long
    mUB = UBound(Textos) + 1
    If mUB = 1 And Textos(0) = "" Then
        'voy a escribir en el primero!
        Textos(0) = sTXT
    Else
        ReDim Preserve Textos(mUB): Textos(mUB) = sTXT
        'meto color de pecho
        ReDim Preserve mForeColor1(mUB): mForeColor1(mUB) = lColor
    End If
End Sub

Public Sub SetVarColor(NewVar As Byte)
    mVarColor = NewVar
End Sub

Public Sub SetInterval(NewInterval As Long)
    mInterval = NewInterval
End Sub

Public Sub INI()
    
    'dejar cargados los colores!
    'lblROLL1.Visible = False 'solo para medir!
    Set lblROLL1.Font = mFont
    lblROLL1 = Textos(0)
    lblROLL1.ForeColor = mForeColor1(0)
    
    'que elija el primero con tamaño y todo
    nTextoActual = UBound(Textos) 'ntextoActual queda en cero
    lblROLL1.Caption = TextoQueSigue
    
    lblROLL1.Top = UserControl.Height / 2 - lblROLL1.Height / 2
    lblROLL1.Left = UserControl.Width / 2 - lblROLL1.Width / 2

    DirColor = 1
    'arranca siempre mal!
    CalzarTexto
    
    TimerMR.Interval = mInterval
    TimerMR.Enabled = True
End Sub

Public Sub STOP_Roll()
    TimerMR.Enabled = False
End Sub

Public Sub Continue_Roll()
    TimerMR.Enabled = True
End Sub

Private Sub TimerMR_Timer()
    
    Dim NewColor As Long
    
    If DirColor = 1 Then
        'acerco el color y veo si llegue
        NewColor = tRGB.AcercarColores(lblROLL1.ForeColor, UserControl.BackColor, mVarColor)
        lblROLL1.ForeColor = NewColor
        If NewColor = UserControl.BackColor Then
            'llego a esconderse!
            lblROLL1.Caption = TextoQueSigue
            CalzarTexto 'acomodarlo para que se vea siempre centrado
            DirColor = 2
        End If
    Else 'esta empezando a mostrar
        'acerco el color y veo si llegue
        NewColor = tRGB.AcercarColores(lblROLL1.ForeColor, mForeColor1(nTextoActual), mVarColor)
        lblROLL1.ForeColor = NewColor
        If NewColor = mForeColor1(nTextoActual) Then
            'llego a esconderse!
            DirColor = 1
        End If
    End If
    
    'lblROLL1.Refresh
End Sub

Private Function TextoQueSigue() As String
    nTextoActual = nTextoActual + 1
    If nTextoActual > UBound(Textos) Then nTextoActual = 0
    TextoQueSigue = Textos(nTextoActual)
End Function

Private Sub CalzarTexto()
    'el label rool ya se escribio pero se debe agrandar o achicar segun corresponda!
    
    'si lo pasa en ancho o alto lo voy achicando
    If lblROLL1.Height > UserControl.Height Or lblROLL1.Width > UserControl.Width Then
        Do While mFont.Size > 5 'me aseguro que no genere un fakin error
            mFont.Size = mFont.Size - 1
            'tiene autosize, se acomoda solo!
            If lblROLL1.Height < UserControl.Height And lblROLL1.Width < UserControl.Width Then
                Exit Do 'llegue a un punto joia
            End If
        Loop
    Else 'es mas chico, lo puedo agrandar
        Do While mFont.Size < 90 'me aseguro que no genere un fakin error
            mFont.Size = mFont.Size + 1
            'tiene autosize, se acomoda solo!
            If lblROLL1.Height > UserControl.Height Or lblROLL1.Width > UserControl.Width Then
                mFont.Size = mFont.Size - 1 'vuelvo al ultimo punto joia
                Exit Do
            End If
        Loop
    End If
    
    lblROLL1.Top = UserControl.Height / 2 - lblROLL1.Height / 2
    lblROLL1.Left = UserControl.Width / 2 - lblROLL1.Width / 2
    
End Sub

Public Sub SetFont1(SF1 As StdFont)
    Set mFont = SF1
    Set lblROLL1.Font = mFont
End Sub

Public Sub SetForeColor1(FC1 As OLE_COLOR, I As Long)
    mForeColor1(I) = FC1
    'lblROLL1.ForeColor = mForeColor1
End Sub

Public Sub Clear()
    lblROLL1 = ""
End Sub

Private Sub UserControl_Click()
    INI
End Sub

Private Sub UserControl_Initialize()
    'valores predeterminados
    mFont.Bold = True
    mFont.Name = "Tahoma"
    mFont.Size = 8
    ReDim mForeColor1(0)
    mForeColor1(0) = &HA9C8C9
    mVarColor = 10
    mInterval = 70
    Set TimerMR = New tbrTimer.clsTimer
    TimerMR.Interval = 0: TimerMR.Enabled = False
    
    ClearTextos
    
'    TextoACola "primera"
'    TextoACola "prueba que puede ser muuuuy pero muuuuuuuuuy larga - larga"
'    TextoACola "de tbrRoll" + vbCrLf + "en" + vbCrLf + "muchos" + vbCrLf + "renglones" + vbCrLf + "renglones" + vbCrLf + "renglones" + vbCrLf + "renglones"
'    TextoACola "se ven " + vbCrLf + "2 renglones?"
'
End Sub

Public Sub ClearTextos()
    'si esta leyendo lo paro!!!
    TimerMR.Enabled = False
    ReDim Textos(0)
End Sub

'en casos de que la matriz es fija por que muestra especificamente X textos se puede hacer
Public Sub ReplaceIndex(I As Long, newText As String)

    'ver que no sean renglones muy largos!
    newText = CortarTextos(newText)

    If I > UBound(Textos) Then ReDim Preserve Textos(I)
    Textos(I) = newText
End Sub

Private Function CortarTextos(t1 As String) As String
    'es encesario ver que todos los renglones tengan un maximo X de caracteres si fuera necesario
    'segun la ubicación del objeto
    
    If t1 = "" Then
        CortarTextos = t1
        Exit Function
    End If
    
    Dim TMP As String
    
    TMP = t1
    'si es cero no queria nada
    If mMaxlargoRenglon = 0 Then Exit Function
    
    Dim MatrizFinal() As String, C As Long
    C = 0
    Dim SP() As String
    SP = Split(t1, vbCrLf)
    Dim A As Long
    For A = 0 To UBound(SP)
        If Len(SP(A)) > mMaxlargoRenglon Then
            'buscar un espacio para cortar por palabras
            Dim B As Long
            For B = mMaxlargoRenglon To 1 Step -1
                If Mid(SP(A), B, 1) = " " Then
                                    
                    'el renglon esta ok
                    ReDim Preserve MatrizFinal(C)
                    MatrizFinal(C) = Mid(SP(A), 1, B)
                    C = C + 1
                    
                    ReDim Preserve MatrizFinal(C)
                    MatrizFinal(C) = Mid(SP(A), B + 1, Len(SP(A)) - B)
                    C = C + 1
                    Exit For
                    'ver que lo que queda no se pase!
                    'XXXX
                End If
            Next B
        Else
            'el renglon esta ok
            ReDim Preserve MatrizFinal(C)
            MatrizFinal(C) = SP(A)
            C = C + 1
        End If
    Next A
    TMP = ""
    For A = 0 To UBound(MatrizFinal)
        TMP = TMP + MatrizFinal(A)
        If A < UBound(MatrizFinal) Then TMP = TMP + vbCrLf
    Next A
    
    CortarTextos = TMP
    
End Function

Private Sub UserControl_Resize()
    'durante la ejecuciion puede cambiar de tamaño y dejar el texto regalado!
    CalzarTexto
End Sub
