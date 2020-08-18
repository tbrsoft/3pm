VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsbDEV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'todo sobre el dispositivo USB conectado

Private CantUSBs As Long

Private VolumeNames() As String
Private TotalSizesMB() As Single
Private FreeSpacesMB() As Single
Private Letters() As String
Private mDevSel As Long 'de los que hay el elegido

Public Property Get DevSel() As Long
    DevSel = mDevSel
End Property

Public Property Let DevSel(NewDevSel As Long)
    mDevSel = NewDevSel 'sera el valor predeterminado cuando no se indique
End Property

Public Function GetCantidadUSB() As Long
    GetCantidadUSB = CantUSBs
End Function

Public Function CanSave(cantMB As Long, iDev As Long) As Boolean
    'me dice si puedo o no grabar una determinada cantidad de MB
    If iDev = -1 Then
        CanSave = (cantMB < GetFreeMB(mDevSel))
    Else
        CanSave = (cantMB < GetFreeMB(iDev))
    End If
    
End Function

'iniciacion de todo, buysca y guarda los datos de todos
Public Function FindUsb()
    CantUSBs = 0
    usbLetter = ""
    Dim S As String
    Dim D As Drive
    For Each D In fso.Drives
        If D.IsReady Then 'por mas que no tenga disketera y la bios diga que si busca los diskettes
            If D.DriveType = Removable Then
                CantUSBs = CantUSBs + 1
                ReDim Preserve VolumeNames(CantUSBs)
                ReDim Preserve TotalSizesMB(CantUSBs)
                ReDim Preserve FreeSpacesMB(CantUSBs)
                ReDim Preserve Letters(CantUSBs)
                
                VolumeNames(CantUSBs) = D.VolumeName
                TotalSizesMB(CantUSBs) = Round((D.TotalSize / 1048576), 2)
                FreeSpacesMB(CantUSBs) = Round((D.FreeSpace / 1048576), 2)
                Letters(CantUSBs) = D.DriveLetter
            End If
        End If
    Next
End Function

Public Function RefreshValues(iDev As Long) As Long
    
    'me dice si puedo o no grabar una determinada cantidad de MB
    Dim i As Long
    If iDev = -1 Then
        i = mDevSel
    Else
        i = iDev
    End If
    
    RefreshValues = -2 'puede ser que ya no este !!!!
    Dim D As Drive
    For Each D In fso.Drives
        'buscar el mismo que estoy usando
        If D.DriveLetter = Letters(i) Then
            If D.IsReady Then 'por mas que no tenga disketera y la bios diga que si busca los diskettes
                VolumeNames(i) = D.VolumeName
                TotalSizesMB(i) = Round((D.TotalSize / 1048576), 2)
                FreeSpacesMB(i) = Round((D.FreeSpace / 1048576), 2)
                'Letters(I) = D.DriveLetter
                Exit Function
            Else
                RefreshValues = -1
                Exit Function
            End If
        End If
    Next
    
End Function

Public Function GetLetterUSB(i As Long) As String
    If i > CantUSBs Then Exit Function
    If i = -1 Then
        GetLetterUSB = Letters(mDevSel)
    Else
        GetLetterUSB = Letters(i)
    End If
End Function

Public Function GetNameUSB(i As Long) As String
    If i > CantUSBs Then Exit Function
    If i = -1 Then
        GetNameUSB = VolumeNames(mDevSel)
    Else
        GetNameUSB = VolumeNames(i)
    End If
End Function

Public Function GetTotalMB(i As Long) As Single
    If i > CantUSBs Then Exit Function
    If i = -1 Then
        GetTotalMB = TotalSizesMB(mDevSel)
    Else
        GetTotalMB = TotalSizesMB(i)
    End If
End Function

Public Function GetFreeMB(i As Long) As Single
    If i > CantUSBs Then Exit Function
    If i = -1 Then
        GetFreeMB = FreeSpacesMB(mDevSel)
    Else
        GetFreeMB = FreeSpacesMB(i)
    End If
End Function

Private Sub Class_Initialize()
    mDevSel = 0
End Sub
