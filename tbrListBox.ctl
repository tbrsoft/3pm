VERSION 5.00
Begin VB.UserControl tbrListBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000040&
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4890
   ScaleHeight     =   2100
   ScaleWidth      =   4890
   Begin VB.PictureBox frTbrListBox 
      AutoRedraw      =   -1  'True
      Height          =   1635
      Left            =   30
      Picture         =   "tbrListBox.ctx":0000
      ScaleHeight     =   1575
      ScaleWidth      =   4695
      TabIndex        =   2
      Top             =   360
      Width           =   4755
      Begin VB.Label lblPLUS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   4230
         TabIndex        =   5
         Top             =   90
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblFULL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "full - invisible"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   1170
         Visible         =   0   'False
         Width           =   4245
      End
      Begin VB.Label lblTXT 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   " tbrListBox"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   105
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblIndice 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   105
         Visible         =   0   'False
         Width           =   195
      End
   End
   Begin VB.PictureBox picFondoTitle 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   " tbrListBox"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   30
         TabIndex        =   1
         Top             =   60
         Width           =   4785
      End
   End
End
Attribute VB_Name = "tbrListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private mCaption As String
Private mCaptionHide As String
Private mListIndex As Long
Private mListCount As Long
Private mBackColor As OLE_COLOR
Private mBackColorItems As OLE_COLOR
Private mForeColorItems As OLE_COLOR
Private mBackColorItemsSel As OLE_COLOR
Private mForeColorItemsSel As OLE_COLOR
Private mAutoHeight As Long 'que el alto del listBox sea el de su ultimo item
'lo puse long para futuras opciones como un maximo o minimo
Private mMaxHeight As Long 'maximo alto si es automático!
'si es cero no da bola
Const mBackColorDEF = &H40&
Const mBackColorItemsDEF = &H80C0FF

Const mForeColorItemsDEF = &H0&
Const mBackColorItemsSelDEF = &H80C0FF
Const mForeColorItemsSelDEF = &H0&

Const mTituloDEF = "Titulo"
Private mTitulo As String
Private mPlusVisible As Boolean


'TRANSPARENCIA
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Private Const AC_SRC_OVER = &H0

Private Declare Function AlphaBlend Lib "msimg32.dll" _
    (ByVal hdcDest As Long, ByVal xOriginDest As Long, _
    ByVal yOriginDest As Long, ByVal WidthDest As Long, _
    ByVal HeightDest As Long, ByVal hdcSrc As Long, _
    ByVal xOriginSrc As Long, ByVal yOriginSrc As Long, _
    ByVal WidthSrc As Long, ByVal HeightSrc As Long, _
    ByVal BLENDFUNCT As Long) As Long
    
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" _
    (Destination As Any, Source As Any, ByVal Length As Long)

Dim Blend As BLENDFUNCTION
Dim Blendlong As Long

'EVENTOS
Public Event ClickITEM(IndiceItem As Long)

Public Property Get Caption()
    Caption = mCaption
End Property

Public Property Get CaptionHide()
    CaptionHide = mCaptionHide
End Property

Public Property Get Titulo() As String
    Titulo = mTitulo
End Property

Public Property Let Titulo(NewTitulo As String)
    mTitulo = NewTitulo
    lblTitulo = NewTitulo
End Property

Public Sub AddItem(NuevoElemento As String, Optional PLUS As String = "", Optional NewElemHide As String = "")
    On Error GoTo ErrF1
    
    tERR.Anotar "aaop", NuevoElemento, PLUS
    'so no define el Hide, queda lo mismo...
    If NewElemHide = "" Then NewElemHide = NuevoElemento
    
    'agragra al ultimo el registro
    'el plus es uno a la derecha con alineacion derecha
    mListCount = mListCount + 1
    Load lblIndice(mListCount)
    Load lblTXT(mListCount)
    Load lblPLUS(mListCount)
    'agregar tambien uno oculto que tiene el valor real
    'de esta forma muestro lo que quiero y tengo un buen dato adentro
    tERR.Anotar "aaop2", mListCount
    Load lblFULL(mListCount)
    
    'ubicarlos
    lblIndice(mListCount).Top = lblIndice(mListCount - 1).Top + lblIndice(mListCount - 1).Height - 3
    lblTXT(mListCount).Top = lblIndice(mListCount).Top
    lblIndice(mListCount).Left = 5 'lblIndice(mListCount - 1).Left
    lblTXT(mListCount).Left = lblTXT(mListCount - 1).Left
    lblPLUS(mListCount).Top = lblIndice(mListCount).Top
    
    tERR.Anotar "aaop3"
    'como esta en autosize el ancho es  muy poco. Lo agrando al tamaño _
    que quiero y despues a lo que estaba
    Dim ExCap As String
    ExCap = lblPLUS(mListCount).Caption
    
    lblPLUS(mListCount) = "000.00"
    lblPLUS(mListCount).Left = (UserControl.Width / 15) - _
        (frTbrListBox.Left + lblPLUS(mListCount).Width + 6)
    lblPLUS(mListCount).Caption = ExCap
    
    
    'bien ancho para que el ajuste de linea no corte!
    lblTXT(mListCount).Width = frTbrListBox.Width + 20 ' - lblTXT(mListCount).Left
    
    tERR.Anotar "aaoq"
    'cargarlos
    If mListCount - 1 < 10 Then
        lblIndice(mListCount).Caption = "0" + CStr(mListCount - 1)
    Else
        lblIndice(mListCount).Caption = mListCount - 1
    End If
    lblTXT(mListCount).Caption = " " + NuevoElemento 'el espacio es para que quede bien
    lblFULL(mListCount).Caption = " " + NewElemHide 'el espacio es para que quede bien
    lblPLUS(mListCount).Caption = PLUS
    'mostrarlos
    lblPLUS(mListCount).BackColor = mBackColorItems
    lblIndice(mListCount).Visible = True
    lblTXT(mListCount).Visible = True
    'mostrarlo segun se pida!
    lblPLUS(mListCount).Visible = mPlusVisible
    lblPLUS(mListCount).ZOrder
    tERR.Anotar "aaor"
    'dejarlo elejido!!!
    mListIndex = mListCount
    'SelIndice mListCount
    
    If mAutoHeight Then
        Dim nH As Long
        nH = (picFondoTitle.Top + picFondoTitle.Height + _
            lblTXT(mListCount).Top + lblTXT(mListCount).Height + 10) * 15
        'ver si tiene topo de maximo!
        If mMaxHeight > 0 And nH > mMaxHeight Then
            'no hago nada
        Else
            UserControl.Height = nH
        End If
        
    End If
    Exit Sub
ErrF1:
    tERR.AppendLog tERR.ErrToTXT(Err), ".zabf"
    Resume Next
    
End Sub

Public Property Let PlusVisible(Vis As Boolean)
    mPlusVisible = Vis
End Property

Public Sub RemoveItem(Optional IND As Long = -1)
    On Error GoTo ErrF1
    
    tERR.Anotar "aaos"
    If IND = -1 Then
        'borrar el elegido
        IND = ListIndex
    End If
    'esconder todo
    lblIndice(IND).Visible = False
    lblTXT(IND).Visible = False
    lblPLUS(IND).Visible = False
    Dim ToRef As Long 'refencia para subirse en la lista
    For A = IND + 1 To mListCount
        'solo correrlo a la ubicacion del proximo visible encima del elegido
        ToRef = 1
        Do
            If lblTXT(A - ToRef).Visible = True Or A = ToRef Then
                tERR.Anotar "aaot"
                Exit Do
            End If
            ToRef = ToRef + 1
        Loop
        lblIndice(A).Top = lblIndice(A - ToRef).Top + lblIndice(A - ToRef).Height + 15
        lblTXT(A).Top = lblTXT(A - ToRef).Top + lblTXT(A - ToRef).Height + 15
        lblPLUS(A).Top = lblPLUS(A - ToRef).Top + lblPLUS(A - ToRef).Height + 15
    Next
    tERR.Anotar "aaou"
    'debe seleccionar alguno!!
    SelNext
    
    Exit Sub
ErrF1:
    tERR.AppendLog tERR.ErrToTXT(Err), ".zabg"
    Resume Next
End Sub


Public Sub inHabilitarItem(Item As Long)
    
    tERR.Anotar "aaov"
    ' hace que un item este enable=false y lo pasa de largo
    ' en ese caso lo tacha y no se puede posicionar en el
    
    lblTXT(Item).ForeColor = &H808080
    lblIndice(Item).ForeColor = &H808080
    
    SelNext 'elige el primero que no este ni inhabilitado ni quitado
    
End Sub

Public Function SelNext()
    On Error GoTo ErrF1
    
    tERR.Anotar "aaow"
    Dim CONT As Long
    'el contador me dice si esta dando vueltas al pedo y no hay nada que elegir
    Dim C As Long
    'no parar al final!!ww
    'ver cual es el primero visible!!
    If mListIndex >= mListCount Then
        C = 1
    Else
        C = mListIndex + 1
    End If
    
    Do
        If lblTXT(C).Visible And lblTXT(C).ForeColor <> &H808080 Then Exit Do
        C = C + 1
        'si llego al final que empieze del principio
        If C > mListCount Then C = 1
        CONT = CONT + 1
        'si dio varias vueltas es que no hay mas discos!!
        If CONT = 100 Then Exit Function
    Loop
    
    ListIndex = C 'ya se carga alla mListIndex
    
    Exit Function
ErrF1:
    tERR.AppendLog tERR.ErrToTXT(Err), ".zabh"
    Resume Next
End Function

Public Function SelFirst()
    On Error GoTo ErrF1
    
    tERR.Anotar "aaox"
    Dim CONT As Long
    'el contador me dice si esta dando vueltas al pedo y no hay nada que elegir
    If mListCount = 0 Then Exit Function
    'ver cual es el primero visible!!
    Dim C As Long
    C = 1
    Do
        If lblTXT(C).Visible And lblTXT(C).ForeColor <> &H808080 Then Exit Do
        C = C + 1
        CONT = CONT + 1
        'si dio varias vueltas es que no hay mas discos!!
        If CONT = 100 Then Exit Function
        If C > mListCount Then Exit Function
    Loop
    ListIndex = C 'ya se carga alla mListIndex
    
    Exit Function
ErrF1:
    tERR.AppendLog tERR.ErrToTXT(Err), ".zabi"
    Resume Next
End Function
Public Function SelPrevious()
    On Error GoTo ErrF1
    
    tERR.Anotar "aaoy"
    Dim CONT As Long
    'el contador me dice si esta dando vueltas al pedo y no hay nada que elegir
    Dim C As Long
    'no parar al final!!
    'ver cual es el primero visible!!
    If mListIndex <= 1 Then
        C = mListCount
    Else
        C = mListIndex - 1
    End If
    
    Do
        If lblTXT(C).Visible And lblTXT(C).ForeColor <> &H808080 Then Exit Do
        C = C - 1
        If C = 0 Then C = mListCount
        CONT = CONT + 1
        'si dio varias vueltas es que no hay mas discos!!
        If CONT = 100 Then Exit Function
    Loop
    
    ListIndex = C 'ya se carga alla mListIndex
    
    Exit Function
ErrF1:
    tERR.AppendLog tERR.ErrToTXT(Err), ".zabj"
    Resume Next
End Function

Public Function SelLast()
    On Error GoTo ErrF1
    
    tERR.Anotar "aaoz"
    Dim CONT As Long
    'el contador me dice si esta dando vueltas al pedo y no hay nada que elegir
    If mListCount = 0 Then Exit Function
    Dim C As Long
    C = mListCount
    Do
        If lblTXT(C).Visible And lblTXT(C).ForeColor <> &H808080 Then Exit Do
        C = C - 1
        'si llego al final y no hay nada salir
        '!!!!
        If C = 0 Then C = mListCount
        'si da varias vueltas es que no hay nada!!!
        CONT = CONT + 1
        'si dio varias vueltas es que no hay mas discos!!
        If CONT = 100 Then Exit Function
    Loop
    ListIndex = C 'ya se carga alla mListIndex
    
    Exit Function
ErrF1:
    tERR.AppendLog tERR.ErrToTXT(Err), ".zabk"
    Resume Next
End Function

Public Property Get ListIndex() As Long
    ListIndex = mListIndex
End Property
Public Property Let ListIndex(NewIndice As Long)
    RaiseEvent ClickITEM(NewIndice)
    SelIndice NewIndice 'aqui ya se carga el mListIndex
End Property

Private Sub SelIndice(IND As Long)
    On Error GoTo ErrF1
    
    tERR.Anotar "aapa"
    'frTbrListBox.Visible = False
    
    'deseleccionar solo el ultimo
    On Error Resume Next
       
    lblIndice(mListIndex).BackStyle = 0
    lblTXT(mListIndex).BackStyle = 0
    lblPLUS(mListIndex).BackStyle = 1
    
    lblIndice(mListIndex).BackColor = mBackColorItems
    lblTXT(mListIndex).BackColor = mBackColorItems
    lblPLUS(mListIndex).BackColor = mBackColorItems
    
    'lblIndice(mListIndex).ForeColor = mForeColorItems
    'lblTXT(mListIndex).ForeColor = mForeColorItems
    lblPLUS(mListIndex).ForeColor = mForeColorItems
       
    'lblIndice(mListIndex).Font.Underline = False
    'lblTXT(mListIndex).Font.Underline = False
    'lblplus(mListIndex).Font.Underline = False
    
    'lblIndice(mListIndex).Font.Italic = False
    'lblTXT(mListIndex).Font.Italic = False
    'lblplus(mListIndex).Font.Italic = False
    
    'lblIndice(mListIndex).Font.Bold = False
    'lblTXT(mListIndex).Font.Bold = False
    'lblplus(mListIndex).Font.Bold = False
    
    'seleccionar el elegido
    
    lblIndice(IND).BackStyle = 1
    lblTXT(IND).BackStyle = 1
    lblPLUS(IND).BackStyle = 1
    
    lblIndice(IND).BackColor = mBackColorItemsSel
    lblTXT(IND).BackColor = mBackColorItemsSel
    lblPLUS(IND).BackColor = mBackColorItemsSel
    
    lblIndice(IND).ForeColor = mForeColorItemsSel
    lblTXT(IND).ForeColor = mForeColorItemsSel
    lblPLUS(IND).ForeColor = mForeColorItemsSel
        
    'lblIndice(IND).Font.Underline = True
    'lblTXT(IND).Font.Underline = True
    'lblplus(IND).Font.Underline = True
    
    'lblIndice(IND).Font.Italic = True
    'lblTXT(IND).Font.Italic = True
    'lblplus(IND).Font.Italic = True
    
    'lblIndice(IND).Font.Bold = True
    'lblTXT(IND).Font.Bold = True
    'lblplus(IND).Font.Bold = True
    
    'cargar la propiedad caption
    mCaption = lblTXT(IND).Caption
    mCaptionHide = lblFULL(IND).Caption
    mListIndex = IND
    'asegurarme que se vea!!!!
    '.....
    'If lblTXT(mListIndex).Top > frTbrListBox.Height - lblTXT(mListIndex).Height - 25 Then
    '    MostrarItem mListIndex, False
    'End If
    'If lblTXT(mListIndex).Top < lblTXT(mListIndex).Height + 25 Then
    '    MostrarItem mListIndex, True
    'End If
   'frTbrListBox.Visible = True
   
   Exit Sub
ErrF1:
    tERR.AppendLog tERR.ErrToTXT(Err), ".zabl"
    Resume Next
End Sub

Private Sub MostrarItem(nItem As Long, Optional ComoPrimero As Boolean = True)
    tERR.Anotar "banp"
    Dim Subir As Long
    If ComoPrimero Then
        Subir = -(lblTXT(nItem).Top - 140)
    Else
        Subir = -(lblTXT(nItem).Top - (frTbrListBox.Height - lblTXT(nItem).Height - 45))
    End If
    Dim F As Control
    For Each F In UserControl.Controls
        If F.Name = "lblIndice" Or F.Name = "lblTXT" Then
            F.Top = F.Top + Subir
        End If
    Next
End Sub

Private Sub lblIndice_Click(Index As Integer)
    ListIndex = Index
End Sub

Private Sub lblTXT_Click(Index As Integer)
    ListIndex = Index
End Sub

Private Sub UserControl_Initialize()
    
    mListIndex = -1
    mListCount = 0
    lblIndice(0).Top = -lblIndice(0).Height
    lblTXT(0).Top = -lblTXT(0).Height
    lblPLUS(0).Top = -lblPLUS(0).Height
    'lo pongo a todo por si esta modo invisible por atras se mida igual
    'a que si mido dentro del ftrbt
    UserControl.ScaleMode = vbPixels
    frTbrListBox.ScaleMode = vbPixels
    
End Sub

Private Sub UserControl_InitProperties()
    tERR.Anotar "aapb"
    'esto sucede cuando se inicaliaza el control
    'las propiedades deben tomar su valor predeterminado
    'sucede cuando se agraga una instancia en algun formulario. Es lo primero que pasa
    mBackColor = mBackColorDEF 'aqui ya se actualiza mBackColor
    mBackColorItems = mBackColorItemsDEF 'aqui ya se actualiza mBackColor
    mForeColorItems = mForeColorItemsDEF 'aqui ya se actualiza mBackColor
    mBackColorItemsSel = mBackColorItemsSelDEF  'aqui ya se actualiza mBackColor
    mForeColorItemsSel = mForeColorItemsSelDEF  'aqui ya se actualiza mBackColor
    mTitulo = mTituloDEF
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'sucede cuando se carga el formulario
    mBackColor = PropBag.ReadProperty("BackColor", mBackColorDEF)
    mBackColorItems = PropBag.ReadProperty("BackColorItems", mBackColorItemsDEF)
    mForeColorItems = PropBag.ReadProperty("ForeColorItems", mForeColorItemsDEF)
    mBackColorItemsSel = PropBag.ReadProperty("BackColorItemsSel", mBackColorItemsSelDEF)
    mForeColorItemsSel = PropBag.ReadProperty("ForeColorItemsSel", mForeColorItemsSelDEF)
    mTitulo = PropBag.ReadProperty("Titulo", mTituloDEF)
    BackColor = mBackColor 'aqui ya se actualiza mBackColor
    BackColorItems = mBackColorItems 'aqui ya se actualiza mBackColor
    ForeColorItems = mForeColorItems 'aqui ya se actualiza mBackColor
    BackColorItemsSel = mBackColorItemsSel 'aqui ya se actualiza mBackColor
    ForeColorItemsSel = mForeColorItemsSel  'aqui ya se actualiza mBackColor
    Titulo = mTitulo
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'sucede cuando el programnador graba el formulario en donde esta mi control
    PropBag.WriteProperty "BackColor", mBackColor, mBackColorDEF
    PropBag.WriteProperty "BackColorItems", mBackColorItems, mBackColorItemsDEF
    PropBag.WriteProperty "ForeColorItems", mForeColorItems, mForeColorItemsDEF
    PropBag.WriteProperty "BackColorItemsSel", mBackColorItemsSel, mBackColorItemsSelDEF
    PropBag.WriteProperty "ForeColorItemsSel", mForeColorItemsSel, mForeColorItemsSelDEF
    PropBag.WriteProperty "Titulo", mTitulo, mTituloDEF
End Sub

Public Sub QuitarTitulo()
    frTbrListBox.Top = 0
    frTbrListBox.Height = UserControl.Height
    picFondoTitle.Visible = False
End Sub

Private Sub UserControl_Resize()
    On Error GoTo ErrF1
    
    tERR.Anotar "aapc"
    'acomodar todo
    frTbrListBox.Left = 2
    'frTbrListBox.Width = UserControl.Width - 3
    lblTitulo.Width = UserControl.Width
    picFondoTitle.Width = UserControl.Width
    tERR.Anotar "aapc2"
    lblTitulo.Left = 0
    frTbrListBox.Height = UserControl.Height - frTbrListBox.Top - 3
    tERR.Anotar "aapc3"
    'cambiar el tamaño de todos los otros
    Dim F As Control
    For Each F In UserControl.Controls
        If F.Name = "lblTXT" Then
            F.Width = frTbrListBox.Width + 90
        End If
        If F.Name = "lblPLUS" Then
            'como esta en autosize el ancho es  muy poco. Lo agrando al tamaño _
            que quiero y despues a lo que estaba
            Dim ExCap As String
            ExCap = lblPLUS(mListCount).Caption
            
            lblPLUS(mListCount) = "000.00"
            
            F.Left = (UserControl.Width / 15) - _
                (frTbrListBox.Left + lblPLUS(mListCount).Width + 6)
            
            lblPLUS(mListCount).Caption = ExCap
        End If
    Next
    
    Exit Sub
ErrF1:
    tERR.AppendLog tERR.ErrToTXT(Err), "tbrLSTBOX.zabl"
    Resume Next
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property
Public Property Let BackColor(NewColor As OLE_COLOR)
    mBackColor = NewColor
    UserControl.BackColor = mBackColor
    frTbrListBox.BackColor = mBackColor
End Property

Public Property Get BackColorItems() As OLE_COLOR
    BackColorItems = mBackColorItems
End Property
Public Property Let BackColorItems(NewColor As OLE_COLOR)
    mBackColorItems = NewColor
    picFondoTitle.BackColor = mBackColor
    'cargar a todos los elementos
    Dim F As Control
    For Each F In UserControl.Controls
        If F.Name = "lblIndice" Or F.Name = "lblTXT" Then F.BackColor = mBackColorItems
    Next
End Property
Public Function IsEnableItem() As Boolean
    If lblTXT(mListIndex).ForeColor = &H808080 Then
        IsEnableItem = False
    Else
        IsEnableItem = True
    End If
End Function
Public Property Get ForeColorItems() As OLE_COLOR
    ForeColorItems = mForeColorItems
End Property
Public Property Let ForeColorItems(NewColor As OLE_COLOR)
    mForeColorItems = NewColor
    'poner el titulo tambien NOOO
    'lblTitulo.ForeColor = mForeColorItemsSel
    
    'cargar a todos los elementos
    Dim F As Control
    For Each F In UserControl.Controls
        If F.Name = "lblIndice" Or F.Name = "lblTXT" Then F.ForeColor = mForeColorItems
    Next
End Property

Public Property Get BackColorItemsSel() As OLE_COLOR
    BackColorItemsSel = mBackColorItemsSel
End Property
Public Property Let BackColorItemsSel(NewColor As OLE_COLOR)
    mBackColorItemsSel = NewColor
    'poner el titulo tambien NOOO
    'lblTitulo.BackColor = mBackColorItemsSel
End Property

Public Property Get ForeColorItemsSel() As OLE_COLOR
    ForeColorItemsSel = mForeColorItemsSel
End Property
Public Property Let ForeColorItemsSel(NewColor As OLE_COLOR)
    mForeColorItemsSel = NewColor
End Property

Public Sub UnSelectAll()
    tERR.Anotar "aapd"
    ''deseleccionar todos
    Dim F As Control
    For Each F In UserControl.Controls
        If F.Name = "lblIndice" Or F.Name = "lblTXT" Or F.Name = "lblPLUS" Then
            F.BackColor = mBackColorItems
            F.ForeColor = mForeColorItems
        End If
    Next
End Sub

Public Sub Clear()
    tERR.Anotar "aape"
    Dim F As Control
    'descargatr TODO!!
    For Each F In UserControl.Controls
        If F.Name = "lblIndice" Or F.Name = "lblTXT" _
            Or F.Name = "lblPLUS" Or F.Name = "lblFULL" Then
            If F.Index > 0 Then Unload F
        End If
    Next
    'Avisarle!!!!!
    tERR.Anotar "aapf"
    mListIndex = -1
    mListCount = 0
End Sub

Public Sub ClearImage()
    frTbrListBox.Picture = LoadPicture
End Sub

Public Sub LoadImage(sImage As String, Optional Ajustar As Boolean = False, _
    Optional P2ConColor As PictureBox)
    
    'SE PUEDE ACOMODAR LA MISMA IMAGEN!
    If sImage = "same" Then GoTo Sigue
    If Dir(sImage) = "" Then Exit Sub
    frTbrListBox.Picture = LoadPicture(sImage)
    
Sigue:
    If Ajustar Then
        frTbrListBox.PaintPicture frTbrListBox.Picture, 0, 0, _
            frTbrListBox.ScaleWidth, frTbrListBox.ScaleHeight
        
        
        frTbrListBox.Width = UserControl.Width
        'acomodar los lblPLUS
        UserControl_Resize
        'si era la misma imgen no se aclara contra nada!
        If sImage = "same" Then Exit Sub
        
        P2ConColor.Width = frTbrListBox.Width
        P2ConColor.Height = frTbrListBox.Height
        P2ConColor.BackColor = vbWhite
        P2ConColor.ScaleMode = vbPixels
        frTbrListBox.ScaleMode = vbPixels
        
        Blend.BlendOp = AC_SRC_OVER
        Blend.BlendFlags = 0
        'cantidad de transparecina 255 es full!
        Blend.SourceConstantAlpha = 100
        Blend.AlphaFormat = 0
        
        RtlMoveMemory Blendlong, Blend, 4
        
        AlphaBlend frTbrListBox.hdc, 0, 0, _
            frTbrListBox.ScaleWidth, _
            frTbrListBox.ScaleHeight, _
            P2ConColor.hdc, 0, 0, _
            P2ConColor.ScaleWidth, _
            P2ConColor.ScaleHeight, _
            Blendlong
            
        frTbrListBox.Refresh
        
    End If
    
End Sub

Public Sub QuitarNumero()
    Dim A As Long
    For A = 1 To mListCount
        lblIndice(A).Visible = False
        lblTXT(A).Left = 0
    Next A
End Sub

Public Sub FondoInvisile()
    QuitarTitulo
    Dim A As Long
    For A = 0 To mListCount
        Set lblIndice(A).Container = Me
        Set lblTXT(A).Container = Me
        Set lblPLUS(A).Container = Me
        
        'lblIndice(A).Top = lblIndice(A).Top + picFondoTitle.Top + picFondoTitle.Height
        'lblTXT(A).Top = lblIndice(A).Top
        'lblPLUS(A).Top = lblIndice(A).Top
        
    Next A
    'sacar el fondo pntado
    frTbrListBox.Visible = False
    
End Sub

Public Sub SetListFont(NewF As StdFont)
    Dim A As Long
    For A = 0 To mListCount
        Set lblIndice(A).Font = NewF
        Set lblTXT(A).Font = NewF
        Set lblPLUS(A).Font = NewF
    Next A
    'y el titulo!
    'Set lblTITULO.Font = NewF
    '**** si hago esto al cambiar lbltitulo cambia NewF y por lo tanto todos los demas!!
    lblTitulo.Font.Name = lblIndice(0).Font.Name
    lblTitulo.Font.Size = lblIndice(0).Font.Size + 4
    lblTitulo.Font.Bold = lblIndice(0).Font.Bold
    
    
End Sub

Public Function GetListFont() As StdFont
    Set GetListFont = lblIndice(0).Font
End Function

Public Sub ReEscribirLista()
    'lo uso cuando cambia la fuente y como esta autosize sse debe reajustar y correr los TOP de las lineas
    Dim A As Long
    For A = 1 To mListCount
        lblIndice(A).Top = lblIndice(A - 1).Top + lblIndice(A - 1).Height - 3
        lblTXT(A).Top = lblIndice(A).Top
        lblPLUS(A).Top = lblIndice(A).Top
    Next A
End Sub

Public Property Get AutoHeight() As Long
    AutoHeight = mAutoHeight
End Property

Public Property Let AutoHeight(NewAH As Long)
    mAutoHeight = NewAH
End Property

Public Property Get MaxHeight2() As Long
    MaxHeight2 = mMaxHeight
End Property

Public Property Let MaxHeight2(NewH As Long)
    mMaxHeight = NewH
End Property

