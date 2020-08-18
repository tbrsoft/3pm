VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "CLaves III Edicion"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkHideLargo 
      BackColor       =   &H00000000&
      Caption         =   "Ocultar codigo rec largo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6780
      TabIndex        =   25
      Top             =   6690
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00004080&
      Caption         =   "Todos los clientes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   500
      Left            =   7590
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6150
      Value           =   -1  'True
      Width           =   1635
   End
   Begin VB.OptionButton opSoloCLI 
      BackColor       =   &H00004080&
      Caption         =   "Solo cliente elegido"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   500
      Left            =   7590
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5640
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Ordenar por"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   9270
      TabIndex        =   18
      Top             =   5610
      Width           =   2535
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Codigo recibido corto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   2235
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Variacion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   2235
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   2235
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Clave"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   2235
      End
   End
   Begin VB.TextBox txtToCopy 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Text            =   "Form1.frx":0000
      Top             =   4380
      Width           =   4275
   End
   Begin VB.TextBox txtTipoClave 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7560
      TabIndex        =   16
      Text            =   "Tipo de Clave"
      Top             =   3360
      Width           =   4275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11340
      TabIndex        =   15
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar nueva entrega de clave"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7590
      TabIndex        =   14
      Top             =   3720
      Width           =   4275
   End
   Begin VB.ListBox lstYaEntregadas 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3600
      IntegralHeight  =   0   'False
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   6930
      Width           =   11775
   End
   Begin VB.TextBox txtCodPcLargo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7560
      TabIndex        =   11
      Text            =   "Código PC (formato largo original)"
      Top             =   1710
      Width           =   4275
   End
   Begin VB.TextBox txtVariacion 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7560
      TabIndex        =   10
      Text            =   "Variacion de la clave"
      Top             =   2700
      Width           =   4275
   End
   Begin VB.TextBox txtClaveEntregada 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7560
      TabIndex        =   9
      Text            =   "Clave entregada"
      Top             =   2370
      Width           =   4275
   End
   Begin VB.TextBox txtCodPcCorto 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7560
      TabIndex        =   8
      Text            =   "Código PC (formato corto)"
      Top             =   2040
      Width           =   4275
   End
   Begin VB.TextBox txtOBS 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7560
      TabIndex        =   7
      Text            =   "Observaciones varias"
      Top             =   3030
      Width           =   4275
   End
   Begin MSComCtl2.DTPicker FECHA 
      Height          =   315
      Left            =   7560
      TabIndex        =   6
      Top             =   1380
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   24444929
      CurrentDate     =   38149
   End
   Begin VB.ComboBox cmbCLIENTES 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Form1.frx":0008
      Left            =   7560
      List            =   "Form1.frx":000A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1050
      Width           =   3765
   End
   Begin VB.ListBox lstClaves 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5610
      IntegralHeight  =   0   'False
      Left            =   900
      TabIndex        =   4
      Top             =   1050
      Width           =   6645
   End
   Begin VB.TextBox tAsig 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   60
      Width           =   855
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
      Left            =   900
      TabIndex        =   2
      Top             =   480
      Width           =   9015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar Clave"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   9960
      TabIndex        =   1
      Top             =   60
      Width           =   1875
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   900
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   9015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Claves ya entregadas a otros clientes!!! (SORTED = TRUE)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   30
      TabIndex        =   13
      Top             =   6720
      Width           =   6225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim K As New clsKEYS
Dim FiSiOb As New Scripting.FileSystemObject
Dim FileClis As String 'lista de clientes
Dim FileClaves As String 'claves entregadas

Private Sub chkHideLargo_Click()
    If Option1(0) Then crgLstClaves opSoloCLI, "clave", chkHideLargo
    If Option1(1) Then crgLstClaves opSoloCLI, "cliente", chkHideLargo
    If Option1(2) Then crgLstClaves opSoloCLI, "variacion", chkHideLargo
    If Option1(3) Then crgLstClaves opSoloCLI, "codigopc", chkHideLargo
End Sub

Private Sub Command2_Click()
    'ver que datos se van a cargar
    Dim LineaEntera As String
    LineaEntera = txtClaveEntregada + ":" + _
        cmbCLIENTES + ":" + _
        CStr(FECHA) + ":" + _
        txtTipoClave + ":" + _
        txtVariacion + ":" + _
        txtCodPcCorto + ":" + _
        txtCodPcLargo + ":" + _
        txtOBS
    'para copiar
    txtToCopy = LineaEntera
    
    If MsgBox("Esta seguro que desea agregar:" + vbCrLf + _
        LineaEntera, vbYesNo) = vbNo Then Exit Sub
        
    'agregar un nuevo cliente!!
    Dim TE As TextStream
    'lo crea si no existe!!
    Set TE = FiSiOb.OpenTextFile(FileClaves, ForAppending, True)
        TE.WriteLine LineaEntera
    TE.Close
    Set TE = Nothing
    lstYaEntregadas.AddItem LineaEntera
End Sub

Private Sub Form_Load()
    K.ClaveDLL = "ashjdklahsJKLHASL65456456456"
    Text1 = K.UniquePC
    FECHA = Date
    FileClis = App.Path + "\fclis.txt"
    FileClaves = App.Path + "\fclaves.txt"
    
    'cargar la lista de clientes
    Dim TE As TextStream
    If FiSiOb.FileExists(FileClis) Then
        Set TE = FiSiOb.OpenTextFile(FileClis, ForReading, False)
        Do While Not TE.AtEndOfStream
            cmbCLIENTES.AddItem TE.ReadLine
        Loop
        TE.Close
        Set TE = Nothing
    Else
        cmbCLIENTES.Clear
    End If
    If cmbCLIENTES.ListCount > 0 Then cmbCLIENTES.ListIndex = 0
    
    crgLstClaves False, "clave", chkHideLargo
    
End Sub

Public Sub crgLstClaves(SoloCli As Boolean, PrimeroQue As String, HideLargo As Boolean)
    lstYaEntregadas.Clear
    'solocli es si solo se ven las licencias de un cliente
    
    'primero que dira por que se ordena
    'clave=clave que le entregue
    'cliente=
    'variacion
    'codigopc=codigo pc recibido corto
    
    Dim Partes() As String
    Dim ThisLine As String
    'asi se graba!!!!!!!!!!!!!!!
    'LineaEntera = txtClaveEntregada + ":" + _
        cmbCLIENTES + ":" + _
        CStr(FECHA) + ":" + _
        txtTipoClave + ":" + _
        txtVariacion + ":" + _
        txtCodPcCorto + ":" + _
        txtCodPcLargo + ":" + _
        txtOBS
    'cargar la lista de CLAVES
    If FiSiOb.FileExists(FileClaves) Then
        Set TE = FiSiOb.OpenTextFile(FileClaves, ForReading, False)
        Do While Not TE.AtEndOfStream
            ThisLine = TE.ReadLine
            Partes = Split(ThisLine, ":")
            'si es solo un cliente
            If SoloCli Then
                If Partes(1) <> cmbCLIENTES Then GoTo SIG
            End If
            Select Case PrimeroQue
                Case "clave"
                    'no hya nada que cambiar!!!
                    If HideLargo Then
                        ThisLine = Partes(0) + ":" + Partes(1) + ":" + Partes(2) + ":" + Partes(3) + ":" + Partes(4) + ":" + Partes(5) + ":" + Partes(7)
                    Else
                        ThisLine = ThisLine
                    End If
                Case "cliente"
                    If HideLargo Then
                        ThisLine = Partes(1) + ":" + Partes(0) + ":" + Partes(2) + ":" + Partes(3) + ":" + Partes(4) + ":" + Partes(5) + ":" + Partes(6) + ":" + Partes(7)
                    Else
                        ThisLine = Partes(1) + ":" + Partes(0) + ":" + Partes(2) + ":" + Partes(3) + ":" + Partes(4) + ":" + Partes(5) + ":" + Partes(7)
                    End If
                Case "variacion"
                    If HideLargo Then
                        ThisLine = Partes(4) + ":" + Partes(0) + ":" + Partes(1) + ":" + Partes(2) + ":" + Partes(3) + ":" + Partes(5) + ":" + Partes(7)
                    Else
                        ThisLine = Partes(4) + ":" + Partes(0) + ":" + Partes(1) + ":" + Partes(2) + ":" + Partes(3) + ":" + Partes(5) + ":" + Partes(6) + ":" + Partes(7)
                    End If
                Case "codigopc"
                    If HideLargo Then
                        ThisLine = Partes(5) + ":" + Partes(0) + ":" + Partes(1) + ":" + Partes(2) + ":" + Partes(3) + ":" + Partes(4) + ":" + Partes(7)
                    Else
                        ThisLine = Partes(5) + ":" + Partes(0) + ":" + Partes(1) + ":" + Partes(2) + ":" + Partes(3) + ":" + Partes(4) + ":" + Partes(6) + ":" + Partes(7)
                    End If
            End Select
            'thisline ya esta corregido a lo pedido
            lstYaEntregadas.AddItem ThisLine
SIG:
        Loop
        TE.Close
        Set TE = Nothing
    Else
        lstYaEntregadas.Clear
    End If
End Sub

Private Sub Command1_Click()
    lstClaves.Clear
    Dim A As Long
    'For A = 1 To 5
    '    lstClaves = lstClaves + vbCrLf + "Sin Cargar(" + CStr(A) + "): " + K.CLAVE(aSinCargar, A, Text1)
    'Next
    'For A = 1 To 5
    '    lstClaves = lstClaves + vbCrLf + "Erronea(" + CStr(A) + "): " + K.CLAVE(BErronea, A, Text1)
    'Next
    For A = 1 To 50
        lstClaves.AddItem "Grat:" + CStr(A) + ": " + K.CLAVE(CGratuita, A, Text1)
    Next
    tAsig = K.Asignaciones  'del 50 de gratuita
    
    For A = 1 To 50
        lstClaves.AddItem "Full:" + CStr(A) + ": " + K.CLAVE(GFull, A, Text1)
    Next
    For A = 1 To 50
        lstClaves.AddItem "SupL:" + CStr(A) + ": " + K.CLAVE(HSuperLicencia, A, Text1)
    Next
    
    For A = 1 To 50
        lstClaves.AddItem "Mini:" + CStr(A) + ": " + K.CLAVE(DMinima, A, Text1)
    Next
    For A = 1 To 50
        lstClaves.AddItem "Comu:" + CStr(A) + ": " + K.CLAVE(EComun, A, Text1)
    Next
    For A = 1 To 50
        lstClaves.AddItem "Prem:" + CStr(A) + ": " + K.CLAVE(FPremium, A, Text1)
    Next
End Sub

Private Sub Command3_Click()
    Dim NewCli As String
    NewCli = InputBox("Ingrese el nombre del nuevo cliente")
    If NewCli = "" Then
        MsgBox "Nada se cargo"
        Exit Sub
    End If
    'agregar un nuevo cliente!!
    Dim TE As TextStream
    'lo crea si no existe!!
    Set TE = FiSiOb.OpenTextFile(FileClis, ForAppending, True)
    TE.WriteLine NewCli
    TE.Close
    Set TE = Nothing
    cmbCLIENTES.AddItem NewCli
    cmbCLIENTES = NewCli
End Sub

Private Sub Form_Resize()
    lstYaEntregadas.Width = Me.Width - lstYaEntregadas.Left
End Sub

Private Sub lstClaves_Click()
    'llevar a donde corresponde:
        'el tipo de clave
        'la variacio
        'la clave
    Dim Partes() As String
    Partes = Split(lstClaves, ":")
    txtTipoClave = Partes(0)
    txtVariacion = Partes(1)
    txtClaveEntregada = Trim(Partes(2))
        
End Sub

Private Sub opSoloCLI_Click()
    If Option1(0) Then crgLstClaves opSoloCLI, "clave", chkHideLargo
    If Option1(1) Then crgLstClaves opSoloCLI, "cliente", chkHideLargo
    If Option1(2) Then crgLstClaves opSoloCLI, "variacion", chkHideLargo
    If Option1(3) Then crgLstClaves opSoloCLI, "codigopc", chkHideLargo
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(0) Then crgLstClaves opSoloCLI, "clave", chkHideLargo
    If Option1(1) Then crgLstClaves opSoloCLI, "cliente", chkHideLargo
    If Option1(2) Then crgLstClaves opSoloCLI, "variacion", chkHideLargo
    If Option1(3) Then crgLstClaves opSoloCLI, "codigopc", chkHideLargo
End Sub

Private Sub Option2_Click()
    If Option1(0) Then crgLstClaves opSoloCLI, "clave", chkHideLargo
    If Option1(1) Then crgLstClaves opSoloCLI, "cliente", chkHideLargo
    If Option1(2) Then crgLstClaves opSoloCLI, "variacion", chkHideLargo
    If Option1(3) Then crgLstClaves opSoloCLI, "codigopc", chkHideLargo
End Sub

Private Sub Text1_Change()
    'al cambiar que se vea a que numero de la clave anterior corresponde
    'Text6 = K.UniquePCOLD
    'siempre mayusculas!!!!!!!!!!
    'ya tuve quilombo con tomas porque copiaba de una PC a otra el codigo
    'no tenia internet en la pcx de la fonola
    Text6 = K.GetOldFromNew(UCase(Text1.Text))
    txtCodPcLargo.Text = Text1.Text
End Sub

Private Sub Text6_Change()
    txtCodPcCorto.Text = Text6
End Sub
