VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00004080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inicio de 3PM"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      BackColor       =   &H00004080&
      Caption         =   "Iniciar con explorer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   500
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00004080&
      Caption         =   "Iniciar con progman"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   500
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   2445
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AP As String
Dim WinFolder As String
Dim FSO As New Scripting.FileSystemObject

Private Sub Command1_Click()

    Dim TE As TextStream
    'leer el system.ini y ver si estamos con PROGMAN o EXPLORER
    'copiarlo para no echar moco
    If FSO.FileExists(AP + "system.ini") Then FSO.DeleteFile AP + "system.ini", True
    FSO.CopyFile WinFolder + "system.ini", AP + "system.ini", True
    Set TE = FSO.OpenTextFile(AP + "system.ini")
    Dim TodoSystem() As String
    Dim ActualShell As String, UbicShell As Long
    c = 1
    Do While Not TE.AtEndOfStream
        ReDim Preserve TodoSystem(c)
        TodoSystem(c) = TE.ReadLine
        If LCase(txtInLista(TodoSystem(c), 0, "=")) = "shell" Then
            UbicShell = c
            ActualShell = txtInLista(TodoSystem(c), 1, "=")
            'no salir para que se copie todo
        End If
        c = c + 1
    Loop
    TE.Close
    If Option2 Then TodoSystem(UbicShell) = "Shell=explorer.exe"
    If Option1 Then TodoSystem(UbicShell) = "Shell=progman.exe"
    'volver a escribir el archivo
    If FSO.FileExists(AP + "system.ini") Then FSO.DeleteFile AP + "system.ini", True
    Set TE = FSO.CreateTextFile(AP + "system.ini", True)
    For A = 1 To UBound(TodoSystem)
        TE.WriteLine TodoSystem(A)
    Next
    TE.Close
    If FSO.FileExists(WinFolder + "OLDsystem.ini") Then FSO.DeleteFile WinFolder + "OLDsystem.ini", True
    If FSO.FileExists(WinFolder + "system.ini") Then FSO.MoveFile WinFolder + "system.ini", WinFolder + "OLDsystem.ini"
    FSO.MoveFile AP + "system.ini", WinFolder + "system.ini"
    
    MsgBox "El cambio se realizo correctamente"
    
    Set FSO = Nothing
    
    Unload Me
    
End Sub

Public Function txtInLista(lista As String, Orden As Long, Separador As String) As String
    'devuelve "OUT LISTA" si se solicita un orden no existente
    'separador es la "," o "-"
    'si pongo 99999 en orden saco el ultimo
    Dim lAct As String, lOrden As Integer
    Dim palabra(40) As String
    Dim c As Integer
    c = 1: lOrden = 0
    Do While c <= Len(lista)
        lAct = Mid(lista, c, 1)
        If lAct = Separador Then
            lOrden = lOrden + 1
        Else
            palabra(lOrden) = palabra(lOrden) + lAct
            If lOrden > Orden Then Exit Do
        End If
        c = c + 1
    Loop
    'si oreden solicitado>ultimo oreden de la lista...
    If Orden > lOrden Then
        If Orden = 99999 Then
            'tengo el ultimo. JOYA para ultima carpeta de path
            txtInLista = palabra(lOrden): Exit Function
        End If
        If Orden = 99998 Then
            'tengo el ultimo. JOYA para ultima carpeta de path
            txtInLista = palabra(lOrden - 1): Exit Function
        End If
        If Orden <> 99999 And Orden <> 99998 Then
            txtInLista = "OUT LISTA": Exit Function
        End If
    End If
    txtInLista = palabra(Orden)
End Function

Private Sub Form_Load()
    AP = App.Path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    WinFolder = FSO.GetSpecialFolder(WindowsFolder)
    If Right(WinFolder, 1) <> "\" Then WinFolder = WinFolder + "\"
End Sub
