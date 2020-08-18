VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXT 
      Height          =   7000
      Left            =   270
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   480
      Width           =   9735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar Texto"
      Height          =   405
      Left            =   210
      TabIndex        =   0
      Top             =   90
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    'funcion GetMem1 para ver de la mother
    TXT = "BIOS DATE" + vbCrLf + GetBIOSDate + vbCrLf + vbCrLf
    'WBem
    Dim ObjSet As SWbemObjectSet
    Dim SERV As SWbemServices
    
    Set SERV = GetObject("WinMgmts:")
    
    Set ObjSet = Nothing
    'datos del ventilador
    TXT = TXT + "WBem" + vbCrLf
    TXT = TXT + "INFORMACION DE LA BIOS" + vbCrLf
    Set ObjSet = SERV.InstancesOf("Win32_Bios")
    If ObjSet.Count > 0 Then
        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
        'Dim BIOS As SWbemObject
        For Each BIOS In ObjSet
            'la primera propiedad es una matriz
            Dim LastNum As Long
            LastNum = -1
            For AAA = 0 To 40
            'ver toidas las cosas que soporta. Muestra solo los numeros que puede usar.
            'La matriz es tan grande como funcoines tenga la mother. El maximo es 39 funciones
                Dim nCH As Long
                nCH = BIOS.BiosCharacteristics(AAA)
                Select Case nCH
                    Case 0: strf = "Reserved"
                    Case 1: strf = "Reserved"
                    Case 2: strf = "Unknown"
                    Case 3: strf = "BIOS Characteristics Not Supported"
                    Case 4: strf = "ISA is supported"
                    Case 5: strf = "MCA is supported"
                    Case 6: strf = "EISA is supported"
                    Case 7: strf = "PCI is supported"
                    Case 8: strf = "PC Card (PCMCIA) is supported"
                    Case 9: strf = "Plug and Play is supported"
                    Case 10: strf = "APM is supported"
                    Case 11: strf = "BIOS is Upgradeable (Flash)"
                    Case 12: strf = "BIOS shadowing is allowed"
                    Case 13: strf = "VL-VESA is supported"
                    Case 14: strf = "ESCD support is available"
                    Case 15: strf = "Boot from CD is supported"
                    Case 16: strf = "Selectable Boot is supported"
                    Case 17: strf = "BIOS ROM is socketed"
                    Case 18: strf = "Boot From PC Card (PCMCIA) is supported"
                    Case 19: strf = "EDD (Enhanced Disk Drive) Specification is supported"
                    Case 20: strf = "Int 13h - Japanese Floppy for NEC 9800 1.2mb (3.5, 1k Bytes/Sector, 360 RPM) is supported"
                    Case 21: strf = "Int 13h - Japanese Floppy for Toshiba 1.2mb (3.5, 360 RPM) is supported"
                    Case 22: strf = "Int 13h - 5.25 / 360 KB Floppy Services are supported"
                    Case 23: strf = "Int 13h - 5.25 /1.2MB Floppy Services are supported"
                    Case 24: strf = "Int 13h - 3.5 / 720 KB Floppy Services are supported"
                    Case 25: strf = "Int 13h - 3.5 / 2.88 MB Floppy Services are supported"
                    Case 26: strf = "Int 5h, Print Screen Service is supported"
                    Case 27: strf = "Int 9h, 8042 Keyboard services are supported"
                    Case 28: strf = "Int 14h, Serial Services are supported"
                    Case 29: strf = "Int 17h, printer services are supported"
                    Case 30: strf = "Int 10h, CGA/Mono Video Services are supported"
                    Case 31: strf = "NEC PC-98"
                    Case 32: strf = "ACPI supported"
                    Case 33: strf = "USB Legacy is supported"
                    Case 34: strf = "AGP is supported"
                    Case 35: strf = "I2O boot is supported"
                    Case 36: strf = "LS-120 boot is supported"
                    Case 37: strf = "ATAPI ZIP Drive boot is supported"
                    Case 38: strf = "1394 boot is supported"
                    Case 39: strf = "Smart Battery supported"
                End Select
                If LastNum = nCH Or nCH > 39 Then
                    Exit For
                Else
                    LastNum = nCH
                End If
                TXT = TXT + "BiosCharacteristics (" + strf + ")" + vbCrLf
            Next
            
            TXT = TXT + "BuildNumber: " + CStr(NN(BIOS.BuildNumber)) + vbCrLf
            TXT = TXT + "Caption: " + CStr(NN(BIOS.Caption)) + vbCrLf
            TXT = TXT + "CodeSet: " + CStr(NN(BIOS.CodeSet)) + vbCrLf
            TXT = TXT + "CurrentLanguage: " + CStr(NN(BIOS.CurrentLanguage)) + vbCrLf
            TXT = TXT + "Description: " + CStr(NN(BIOS.Description)) + vbCrLf
            TXT = TXT + "IdentificationCode: " + CStr(NN(BIOS.IdentificationCode)) + vbCrLf
            TXT = TXT + "InstallableLanguages: " + CStr(NN(BIOS.InstallableLanguages)) + vbCrLf
            TXT = TXT + "InstallDate: " + CStr(NN(BIOS.InstallDate)) + vbCrLf
            TXT = TXT + "LanguageEdition: " + CStr(NN(BIOS.LanguageEdition)) + vbCrLf
            TXT = TXT + "ListOfLanguages: " + CStr(NN(BIOS.ListOfLanguages(0))) + vbCrLf
            TXT = TXT + "ListOfLanguages: " + CStr(NN(BIOS.ListOfLanguages(1))) + vbCrLf
            TXT = TXT + "ListOfLanguages: " + CStr(NN(BIOS.ListOfLanguages(2))) + vbCrLf
            TXT = TXT + "ListOfLanguages: " + CStr(NN(BIOS.ListOfLanguages(3))) + vbCrLf
            TXT = TXT + "ListOfLanguages: " + CStr(NN(BIOS.ListOfLanguages(4))) + vbCrLf
            TXT = TXT + "Manufacturer: " + CStr(NN(BIOS.Manufacturer)) + vbCrLf
            TXT = TXT + "Name: " + CStr(NN(BIOS.Name)) + vbCrLf
            TXT = TXT + "OtherTargetOS: " + CStr(NN(BIOS.OtherTargetOS)) + vbCrLf
            TXT = TXT + "PrimaryBIOS: " + CStr(NN(BIOS.PrimaryBIOS)) + vbCrLf
            TXT = TXT + "ReleaseDate: " + CStr(NN(BIOS.ReleaseDate)) + vbCrLf
            TXT = TXT + "SerialNumber: " + CStr(NN(BIOS.SerialNumber)) + vbCrLf
            TXT = TXT + "BuildNumber: " + CStr(NN(BIOS.BuildNumber)) + vbCrLf
            TXT = TXT + "SMBIOSBIOSVersion: " + CStr(NN(BIOS.SMBIOSBIOSVersion)) + vbCrLf
            TXT = TXT + "SMBIOSMajorVersion: " + CStr(NN(BIOS.SMBIOSMajorVersion)) + vbCrLf
            TXT = TXT + "SMBIOSMinorVersion: " + CStr(NN(BIOS.SMBIOSMinorVersion)) + vbCrLf
            TXT = TXT + "SMBIOSPresent: " + CStr(NN(BIOS.SMBIOSPresent)) + vbCrLf
            TXT = TXT + "SoftwareElementID: " + CStr(NN(BIOS.SoftwareElementID)) + vbCrLf
            TXT = TXT + "SoftwareElementState: " + CStr(NN(BIOS.SoftwareElementState)) + vbCrLf
            TXT = TXT + "Status: " + CStr(NN(BIOS.Status)) + vbCrLf
            TXT = TXT + "TargetOperatingSystem: " + CStr(NN(BIOS.TargetOperatingSystem)) + vbCrLf
            TXT = TXT + "Version: " + CStr(NN(BIOS.Version)) + vbCrLf
        Next
    Else
        TXT = TXT + "No se encontraron"
    End If
    
    Set ObjSet = Nothing
    'Win32_Processor Represents a device capable of interpreting a sequence of machine instructions on a Win32 computer system.
    'datos del ventilador
    TXT = TXT + vbCrLf
    TXT = TXT + "INFORMACION DEL PROCESADOR" + vbCrLf
    Set ObjSet = SERV.InstancesOf("Win32_Processor")
    If ObjSet.Count > 0 Then
        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
        For Each MICRO In ObjSet
            TXT = TXT + "Availability: " + CStr(NN(MICRO.Availability)) + vbCrLf
            TXT = TXT + "AddressWidth: " + CStr(NN(MICRO.AddressWidth)) + vbCrLf
            TXT = TXT + "Architecture: " + CStr(NN(MICRO.Architecture)) + vbCrLf
            TXT = TXT + "CpuStatus: " + CStr(NN(MICRO.CpuStatus)) + vbCrLf
            TXT = TXT + "CreationClassName: " + CStr(NN(MICRO.CreationClassName)) + vbCrLf
            TXT = TXT + "CurrentClockSpeed: " + CStr(NN(MICRO.CurrentClockSpeed)) + vbCrLf
            TXT = TXT + "CurrentVoltage: " + CStr(NN(MICRO.CurrentVoltage)) + vbCrLf
            TXT = TXT + "DataWidth: " + CStr(NN(MICRO.DataWidth)) + vbCrLf
            TXT = TXT + "Description: " + CStr(NN(MICRO.Description)) + vbCrLf
            TXT = TXT + "DeviceID: " + CStr(NN(MICRO.DeviceID)) + vbCrLf
            TXT = TXT + "ExtClock: " + CStr(NN(MICRO.ExtClock)) + vbCrLf
            TXT = TXT + "Family: " + CStr(NN(MICRO.Family)) + vbCrLf
            TXT = TXT + "L2CacheSize: " + CStr(NN(MICRO.L2CacheSize)) + vbCrLf
            TXT = TXT + "L2CacheSpeed: " + CStr(NN(MICRO.L2CacheSpeed)) + vbCrLf
            TXT = TXT + "Level: " + CStr(NN(MICRO.Level)) + vbCrLf
            TXT = TXT + "LoadPercentage: " + CStr(NN(MICRO.LoadPercentage)) + vbCrLf
            TXT = TXT + "Manufacturer: " + CStr(NN(MICRO.Manufacturer)) + vbCrLf
            TXT = TXT + "MaxClockSpeed: " + CStr(NN(MICRO.MaxClockSpeed)) + vbCrLf
            TXT = TXT + "Name: " + CStr(NN(MICRO.Name)) + vbCrLf
            TXT = TXT + "OtherFamilyDescription: " + CStr(NN(MICRO.OtherFamilyDescription)) + vbCrLf
            TXT = TXT + "PNPDeviceID: " + CStr(NN(MICRO.PNPDeviceID)) + vbCrLf
            TXT = TXT + "ProcessorId: " + CStr(NN(MICRO.ProcessorId)) + vbCrLf
            TXT = TXT + "ProcessorType: " + CStr(NN(MICRO.ProcessorType)) + vbCrLf
            TXT = TXT + "Revision: " + CStr(NN(MICRO.Revision)) + vbCrLf
            TXT = TXT + "Role: " + CStr(NN(MICRO.Role)) + vbCrLf
            TXT = TXT + "SocketDesignation: " + CStr(NN(MICRO.SocketDesignation)) + vbCrLf
            TXT = TXT + "Status: " + CStr(NN(MICRO.Status)) + vbCrLf
            TXT = TXT + "StatusInfo: " + CStr(NN(MICRO.StatusInfo)) + vbCrLf
            TXT = TXT + "Stepping: " + CStr(NN(MICRO.Stepping)) + vbCrLf
            TXT = TXT + "SystemCreationClassName: " + CStr(NN(MICRO.SystemCreationClassName)) + vbCrLf
            TXT = TXT + "SystemName: " + CStr(NN(MICRO.SystemName)) + vbCrLf
            TXT = TXT + "UniqueId: " + CStr(NN(MICRO.UniqueId)) + vbCrLf
            TXT = TXT + "UpgradeMethod: " + CStr(NN(MICRO.UpgradeMethod)) + vbCrLf
            TXT = TXT + "Version: " + CStr(NN(MICRO.Version)) + vbCrLf
            TXT = TXT + "VoltageCaps: " + CStr(NN(MICRO.VoltageCaps)) + vbCrLf
        Next
    Else
        TXT = TXT + "No se encontraron" + vbCrLf
    End If
    Set ObjSet = Nothing
    
    
    'viejo y glorioso System INFO
    Dim INFO As SYSTEM_INFO
    GetSystemInfo INFO
    TXT = TXT + vbCrLf
    TXT = TXT + "dwActiveProcessorMask: " + CStr(NN(INFO.dwActiveProcessorMask)) + vbCrLf
    TXT = TXT + "dwAllocationGranularity: " + CStr(NN(INFO.dwAllocationGranularity)) + vbCrLf
    TXT = TXT + "dwNumberOrfProcessors: " + CStr(NN(INFO.dwNumberOrfProcessors)) + vbCrLf
    TXT = TXT + "dwOemID: " + CStr(NN(INFO.dwOemID)) + vbCrLf
    TXT = TXT + "dwPageSize: " + CStr(NN(INFO.dwPageSize)) + vbCrLf
    TXT = TXT + "dwProcessorType: " + CStr(NN(INFO.dwProcessorType)) + vbCrLf
    'joya este es unico en cada PC (todos los amd son iguales!!!)
    TXT = TXT + "dwReserved: " + Str(NN(INFO.dwReserved)) + vbCrLf
    'joya este es unico en cada PC
    TXT = TXT + "lpMaximumApplicationAddress: " + CStr(NN(INFO.lpMaximumApplicationAddress)) + vbCrLf
    TXT = TXT + "lpMinimumApplicationAddress: " + CStr(NN(INFO.lpMinimumApplicationAddress)) + vbCrLf
    
End Sub

Private Sub Form_Resize()
    TXT.Width = Me.Width - TXT.Left - 50
    TXT.Height = Me.Height - (Command1.Top + Command1.Height + 50) - 500
End Sub

Private Function NN(Val, Optional DEfault = "NULO")
    'No Nulo
    If IsNull(Val) Then
        NN = DEfault
    Else
        NN = Val
    End If
End Function
