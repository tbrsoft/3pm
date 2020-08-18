VERSION 5.00
Begin VB.Form frmERRORES 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SI SUCEDEN ERRORES EN 3PM"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Grabar Detalle de error"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   120
      TabIndex        =   10
      Top             =   4260
      Width           =   4995
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3300
      Width           =   5000
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   5220
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3300
      Width           =   5000
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   5220
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   870
      Width           =   5000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar descripcion de la falla e Instantánea del sistema"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   5250
      TabIndex        =   2
      Top             =   4260
      Width           =   4995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3630
      TabIndex        =   1
      Top             =   4920
      Width           =   3075
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   870
      Width           =   5000
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion del error:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3060
      Width           =   2970
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "En caso de que su sistema presente fallas notorias pero sin mensaje de error siga las siguinetes instrucciones:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   795
      Left            =   5220
      TabIndex        =   7
      Top             =   30
      Width           =   4950
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "En caso de que su sistema cierre bruscamente por algún mensaje de error"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   765
      Left            =   150
      TabIndex        =   6
      Top             =   60
      Width           =   4950
   End
   Begin VB.Label lblREP 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion de la Falla:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5220
      TabIndex        =   5
      Top             =   3030
      Width           =   2970
   End
End
Attribute VB_Name = "frmERRORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If Text3 = "" Then
        MsgBox "No hay texto para grabar!" + vbCrLf + "Escriba el detalle de la falla y pruebe de nuevo"
        Exit Sub
    End If
    
    Dim nFSO As New Scripting.FileSystemObject
    Dim nTE As TextStream
    
    Set nTE = FSO.CreateTextFile(AP + "REG_DESCR_FALLA.W15", True)
        nTE.WriteLine "DESCRIPCION FALLA:"
        nTE.WriteLine Text3.Text
        nTE.WriteLine "CAMINO"
        nTE.WriteLine tERR.LogAcumulado
    nTE.Close
    
    Set nTE = Nothing
    Set nFSO = Nothing
    
    MsgBox "El detalle se grabo OK con formato W15. Envie ahora los archivos mencionados"
End Sub

Private Sub Command3_Click()
    
    If Text4 = "" Then
        MsgBox "No hay texto para grabar!" + vbCrLf + "Escriba el mensaje de error que cierra 3PM bruscamente y pruebe de nuevo"
        Exit Sub
    End If
    
    Dim nFSO As New Scripting.FileSystemObject
    Dim nTE As TextStream
    
    Set nTE = FSO.CreateTextFile(AP + "REG_DESCR_ERR.W15", True)
        nTE.Write Text4.Text
    nTE.Close
    
    Set nTE = Nothing
    Set nFSO = Nothing
    
    MsgBox "El detalle se grabo OK con formato W15. Envie ahora los archivos mencionados"
    
    'empaquetar todo para mandarlo por el servidor
    'son todos los w15
    'el reg3pm.log
    'el archivo de configuracion (que deberá incluir la versión)
    
End Sub

Private Sub Form_Load()
    MostrarCursor True
    
    Text1 = "  1- Ingrese a la configuración de 3PM" + vbCrLf + _
        "  2- Ingrese a la seccion 'OTRAS OPCIONES'" + vbCrLf + _
        "  3- Active la casilla 'ACTIVAR REGISTRO DE ERROR PERMANENETE'" + vbCrLf + _
        "  4- Presione 'GRABAR'" + vbCrLf + _
        "  5- Cierre 3PM y vuelva a ejecutarlo" + vbCrLf + _
        "  6- Ponga al sistema en las mismas circunstancias en que se dio el error" + _
        " para tratar de que se genere de nuevo" + vbCrLf + _
        "  7- Al aparecer el error y luego de que 3PM se cierre bruscamente" + _
        " escriba el texto de error en la casilla 'Descripcion del error' y luego presione el boton 'Grabar Detalle de error'" + vbCrLf + _
        "  8- Envie por email a tbrsoft (info@tbrsoft.com - tbrsoft@cpcipc.org)" + _
        " los siguientes archivos:" + vbCrLf + _
        "  De la carpeta de 3PM" + vbCrLf + _
        "   * Todos los archivos REG****.W15" + vbCrLf + _
        "   * El archivo reg3PM.log" + vbCrLf + _
        "  De la carpeta de sistema (C:\Windows\System en W98 o Me ó C:\Windows\" + _
        "System32 en WXP):" + vbCrLf + _
        "   * El archivo 3PM.CFG" + vbCrLf + vbCrLf + _
        "Con esta información tbrSoft le enviará en un lapso de 24-72 Hs una " + _
        "respuesta concreta a su problema."

    Text2 = "  1- Ponga al sistema en las mismas circunstancias en que se presenta la falla" + vbCrLf + _
        "  2- Describa la falla en la casilla 'Descripcion de la Falla' y luego presione el boton 'Grabar descripcion de la falla e Instantánea del sistema'" + vbCrLf + _
        "  3- Envie por email a tbrsoft (info@tbrsoft.com - tbrsoft@cpcipc.org)" + _
        " los siguientes archivos:" + vbCrLf + _
        "  De la carpeta de 3PM" + vbCrLf + _
        "   * Todos los archivos REG****.W15" + vbCrLf + _
        "   * El archivo reg3PM.log" + vbCrLf + _
        "  De la carpeta de sistema (C:\Windows\System en W98 o Me ó C:\Windows\" + _
        "System32 en WXP):" + vbCrLf + _
        "   * El archivo 3PM.CFG" + vbCrLf + vbCrLf + _
        "Con esta información tbrSoft le enviará en un lapso de 24-72 Hs una " + _
        "respuesta concreta a su problema."
End Sub
