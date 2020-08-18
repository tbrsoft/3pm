VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKEYS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private FSO As New Scripting.FileSystemObject
Private nFOt As New tbrDATA.clsTODO

'tipos de licencia que se pueden cargar
Public Enum TypeLic
    a_SinLeer = -1
    aSinCargar = 0 'todavia no puso nada
    BErronea = 1 'cuando ya erro 3 o mas veces
    CGratuita = 2 'primera que me puede pedir
    DMinima = 3
    EComun = 4
    FPremium = 5
    GFull = 6
    HSuperLicencia = 7
End Enum

Private mLicencia As TypeLic

Private Sub Class_Initialize()
    'GPF("cd3pm") 'DatosClaveIngresada.DLL
    'asegurarme que exista "dciLib22.dll"
    mLicencia = a_SinLeer
    
    'generar el archivo de esta PC !
    nFOt.DoNow GPF("cd4pm")
    
    'validarlo con la clave si hubiera !
    LICENCIA
End Sub

''sacar despues!
'Public Sub CrearClave()
'    Dim S6 As String 'es la clave de la licencia 2004 2006
'    nFOt.CK GPF("cd4pm"), 6, "c:\de.L37", S6
'End Sub

Public Sub SetFileLic(pARCH As String)
    'recibo el archivo de respuesta
    'ponerlo en la ubicacion deseada!
    FSO.CopyFile pARCH, GPF("cd5pm")
    'le hago copia de seguridad
    FSO.CopyFile GPF("cd5pm"), GPF("cd6pm")
End Sub

Public Function LICENCIA() As TypeLic
    'lee lo grabado para que sea mas rapido,
    'si es la primera vez leo todo
    If mLicencia = a_SinLeer Then IngresaClave
    LICENCIA = mLicencia
End Function

Public Sub IngresaClave()
    'ingresa archivo de licencia y devuelve el tipo esperado
    
    Dim TmpLic As TypeLic
    mLicencia = aSinCargar
    
    'si no cargo el archivo no leer nada
    If FSO.FileExists(GPF("cd5pm")) = False Then
        'a ver si esta la copia!!
        If FSO.FileExists(GPF("cd6pm")) = False Then
            Exit Sub
        Else 'restaurar al original la copia
            FSO.CopyFile GPF("cd6pm"), GPF("cd5pm"), True
        End If
    End If
    Dim TEZ As String 'lista "FINAL:" original
    Dim mNR As Long
    mNR = nFOt.GetNR(GPF("cd5pm"), TEZ)
    Select Case mNR
        Case 2
            TmpLic = CGratuita
        Case 3
            TmpLic = DMinima
        Case 4
            TmpLic = EComun
        Case 5
            TmpLic = FPremium
        Case 6
            TmpLic = GFull
        Case 7
            TmpLic = HSuperLicencia
        Case Else
            TmpLic = BErronea
    End Select
                
    '**********************
    mLicencia = TmpLic
    '**********************
End Sub
