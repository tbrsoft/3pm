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
   WindowState     =   2  'Maximized
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
        
    On Local Error GoTo ErrWBMem
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
            TXT = TXT + "--------------------------------------" + vbCrLf
            'la primera propiedad es una matriz
            Dim LastNum As Long
            LastNum = -1
            For AAA = 0 To UBound(BIOS.BiosCharacteristics)
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
'                If LastNum = nCH Or nCH > 39 Then
'                    Exit For
'                Else
'                    LastNum = nCH
'                End If
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
    TXT = TXT + UCase("INFORMACION DEL PROCESADOR") + vbCrLf
    Set ObjSet = SERV.InstancesOf("Win32_Processor")
    If ObjSet.Count > 0 Then
        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
        For Each MICRO In ObjSet
            TXT = TXT + "--------------------------------------" + vbCrLf
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

'    Set ObjSet = Nothing
'    'The Win32_Fan WMI class represents the properties of a fan device in _
'        the computer system. For example, the CPU cooling fan.
'
'    TXT = TXT + vbCrLf
'    TXT = TXT + ucase("INFORMACION DEL FAN" + vbCrLf
'    Set ObjSet = SERV.InstancesOf("Win32_Fan")
'    If ObjSet.Count > 0 Then
'        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
'        For Each FAN In ObjSet
'            TXT = TXT + "--------------------------------------" + vbCrLf
'            TXT = TXT + "ActiveCooling: " + CStr(NN(FAN.ActiveCooling)) + vbCrLf
'            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
'            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
'            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
'            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
'            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
'            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
'            TXT = TXT + "DesiredSpeed: " + CStr(NN(FAN.DesiredSpeed)) + vbCrLf
'            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
'            TXT = TXT + "ErrorCleared: " + CStr(NN(FAN.ErrorCleared)) + vbCrLf
'            '...................SEGUIR
'        Next
'    Else
'        TXT = TXT + "No se encontraron" + vbCrLf
'    End If

'    Set ObjSet = Nothing
'    'The Win32_HeatPipe WMI class represents the properties of _
'        a heat pipe cooling device.
'
'    TXT = TXT + vbCrLf
'    TXT = TXT + ucase("INFORMACION DEL Win32_HeatPipe" + vbCrLf
'    Set ObjSet = SERV.InstancesOf("Win32_HeatPipe")
'    If ObjSet.Count > 0 Then
'        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
'        For Each FAN In ObjSet
'            TXT = TXT + "--------------------------------------" + vbCrLf
'            TXT = TXT + "ActiveCooling: " + CStr(NN(FAN.ActiveCooling)) + vbCrLf
'            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
'            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
'            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
'            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
'            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
'            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
'            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
'            TXT = TXT + "ErrorCleared: " + CStr(NN(FAN.ErrorCleared)) + vbCrLf
'            '...................SEGUIR
'        Next
'    Else
'        TXT = TXT + "No se encontraron" + vbCrLf
'    End If


'    Set ObjSet = Nothing
'    'The Win32_Refrigeration WMI class represents the properties of a refrigeration _
'    device
'
'    TXT = TXT + vbCrLf
'    TXT = TXT + ucase("INFORMACION DEL Win32_Refrigeration" + vbCrLf
'    Set ObjSet = SERV.InstancesOf("Win32_Refrigeration")
'    If ObjSet.Count > 0 Then
'        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
'        For Each FAN In ObjSet
'            TXT = TXT + "--------------------------------------" + vbCrLf
'            TXT = TXT + "ActiveCooling: " + CStr(NN(FAN.ActiveCooling)) + vbCrLf
'            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
'            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
'            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
'            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
'            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
'            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
'            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
'            TXT = TXT + "ErrorCleared: " + CStr(NN(FAN.ErrorCleared)) + vbCrLf
'            '...................SEGUIR
'        Next
'    Else
'        TXT = TXT + "No se encontraron" + vbCrLf
'    End If

'    Set ObjSet = Nothing
'    'The Win32_TemperatureProbe WMI class represents the properties of _
'        a temperature sensor (electronic thermometer).
'
'    TXT = TXT + vbCrLf
'    TXT = TXT + ucase("INFORMACION DEL Win32_TemperatureProbe" + vbCrLf
'    Set ObjSet = SERV.InstancesOf("Win32_TemperatureProbe")
'    If ObjSet.Count > 0 Then
'        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
'        For Each FAN In ObjSet
'            TXT = TXT + "--------------------------------------" + vbCrLf
'            TXT = TXT + "Accuracy: " + CStr(NN(FAN.ActiveCooling)) + vbCrLf
'            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
'            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
'            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
'            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
'            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
'            TXT = TXT + "CurrentReading: " + CStr(NN(FAN.CurrentReading)) + vbCrLf
'            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
'            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
'            TXT = TXT + "ErrorCleared: " + CStr(NN(FAN.ErrorCleared)) + vbCrLf
'            '...................SEGUIR
'        Next
'    Else
'        TXT = TXT + "No se encontraron" + vbCrLf
'    End If


    Set ObjSet = Nothing
    'The Win32_Keyboard WMI class represents a keyboard installed on a Win32 system.

    TXT = TXT + vbCrLf
    TXT = TXT + UCase("INFORMACION DEL Win32_Keyboard") + vbCrLf
    Set ObjSet = SERV.InstancesOf("Win32_Keyboard")
    If ObjSet.Count > 0 Then
        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
        For Each FAN In ObjSet
            TXT = TXT + "--------------------------------------" + vbCrLf
            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
            TXT = TXT + "ErrorCleared: " + CStr(NN(FAN.ErrorCleared)) + vbCrLf
            TXT = TXT + "ErrorDescription: " + CStr(NN(FAN.ErrorDescription)) + vbCrLf
            TXT = TXT + "InstallDate: " + CStr(NN(FAN.InstallDate)) + vbCrLf
            TXT = TXT + "IsLocked: " + CStr(NN(FAN.IsLocked)) + vbCrLf
            TXT = TXT + "LastErrorCode: " + CStr(NN(FAN.LastErrorCode)) + vbCrLf
            TXT = TXT + "Layout: " + CStr(NN(FAN.Layout)) + vbCrLf
            TXT = TXT + "Name: " + CStr(NN(FAN.Name)) + vbCrLf
            TXT = TXT + "NumberOfFunctionKeys: " + CStr(NN(FAN.NumberOfFunctionKeys)) + vbCrLf
            TXT = TXT + "Password: " + CStr(NN(FAN.Password)) + vbCrLf
            TXT = TXT + "PNPDeviceID: " + CStr(NN(FAN.PNPDeviceID)) + vbCrLf
            nCH = 0
            nCH = UBound(FAN.PowerManagementCapabilities)
            For A = 0 To nCH
                TXT = TXT + "PowerManagementCapabilities " + CStr(A) + ": " + CStr(NN(FAN.PowerManagementCapabilities(A))) + vbCrLf
            Next A
            TXT = TXT + "PowerManagementSupported: " + CStr(NN(FAN.PowerManagementSupported)) + vbCrLf
            TXT = TXT + "Status: " + CStr(NN(FAN.Status)) + vbCrLf
            TXT = TXT + "StatusInfo: " + CStr(NN(FAN.StatusInfo)) + vbCrLf
            TXT = TXT + "SystemCreationClassName: " + CStr(NN(FAN.SystemCreationClassName)) + vbCrLf
            TXT = TXT + "SystemName: " + CStr(NN(FAN.SystemName)) + vbCrLf
        Next
    Else
        TXT = TXT + "No se encontraron" + vbCrLf
    End If

    Set ObjSet = Nothing
    'The Win32_PointingDevice WMI class represents an input device used to point _
    to and select regions on the display of a Win32 computer system. Any device _
    used to manipulate a pointer, or point to the display on a Win32 computer _
    system is a member of this class.

    TXT = TXT + vbCrLf
    TXT = TXT + UCase("INFORMACION DEL Win32_PointingDevice") + vbCrLf
    Set ObjSet = SERV.InstancesOf("Win32_PointingDevice")
    If ObjSet.Count > 0 Then
        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
        For Each FAN In ObjSet
            TXT = TXT + "--------------------------------------" + vbCrLf
            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
            TXT = TXT + "DeviceInterface: " + CStr(NN(FAN.DeviceInterface)) + vbCrLf
            TXT = TXT + "DoubleSpeedThreshold: " + CStr(NN(FAN.DoubleSpeedThreshold)) + vbCrLf
            TXT = TXT + "ErrorCleared: " + CStr(NN(FAN.ErrorCleared)) + vbCrLf
            TXT = TXT + "ErrorDescription: " + CStr(NN(FAN.ErrorDescription)) + vbCrLf
            TXT = TXT + "Handedness: " + CStr(NN(FAN.Handedness)) + vbCrLf
            TXT = TXT + "HardwareType: " + CStr(NN(FAN.HardwareType)) + vbCrLf
            TXT = TXT + "InfFileName: " + CStr(NN(FAN.InfFileName)) + vbCrLf
            TXT = TXT + "InfSection: " + CStr(NN(FAN.InfSection)) + vbCrLf
            TXT = TXT + "InstallDate: " + CStr(NN(FAN.InstallDate)) + vbCrLf
            TXT = TXT + "IsLocked: " + CStr(NN(FAN.IsLocked)) + vbCrLf
            TXT = TXT + "LastErrorCode: " + CStr(NN(FAN.LastErrorCode)) + vbCrLf
            TXT = TXT + "Manufacturer: " + CStr(NN(FAN.Manufacturer)) + vbCrLf
            TXT = TXT + "Name: " + CStr(NN(FAN.Name)) + vbCrLf
            TXT = TXT + "NumberOfButtons: " + CStr(NN(FAN.NumberOfButtons)) + vbCrLf
            TXT = TXT + "PNPDeviceID: " + CStr(NN(FAN.PNPDeviceID)) + vbCrLf
            TXT = TXT + "PointingType: " + CStr(NN(FAN.PointingType)) + vbCrLf
            nCH = 0
            nCH = UBound(FAN.PowerManagementCapabilities)
            For A = 0 To nCH
                TXT = TXT + "PowerManagementCapabilities " + CStr(A) + ": " + CStr(NN(FAN.PowerManagementCapabilities(A))) + vbCrLf
            Next A
            TXT = TXT + "PowerManagementSupported: " + CStr(NN(FAN.PowerManagementSupported)) + vbCrLf
            TXT = TXT + "QuadSpeedThreshold: " + CStr(NN(FAN.QuadSpeedThreshold)) + vbCrLf
            TXT = TXT + "Resolution: " + CStr(NN(FAN.Resolution)) + vbCrLf
            TXT = TXT + "SampleRate: " + CStr(NN(FAN.SampleRate)) + vbCrLf
            TXT = TXT + "Status: " + CStr(NN(FAN.Status)) + vbCrLf
            TXT = TXT + "StatusInfo: " + CStr(NN(FAN.StatusInfo)) + vbCrLf
            TXT = TXT + "Synch: " + CStr(NN(FAN.Synch)) + vbCrLf
            TXT = TXT + "SystemCreationClassName: " + CStr(NN(FAN.SystemCreationClassName)) + vbCrLf
            TXT = TXT + "SystemName: " + CStr(NN(FAN.SystemName)) + vbCrLf
        Next
    Else
        TXT = TXT + "No se encontraron" + vbCrLf
    End If

'    Set ObjSet = Nothing
'    'The Win32_CDROMDrive WMI class represents a CD-ROM drive on a Win32 computer system. Note that the name of the drive does not correspond to the logical drive letter assigned to device.
'
'    TXT = TXT + vbCrLf
'    TXT = TXT + ucase("INFORMACION DEL Win32_CDROMDrive" + vbCrLf
'    Set ObjSet = serv.InstancesOf("Win32_CDROMDrive")
'    If ObjSet.Count > 0 Then
'        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
'        For Each FAN In ObjSet
'            TXT = TXT + "--------------------------------------"+vbcrlf
'            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
'            For AA = 0 To 9
'                TXT = TXT + "Capabilities " + CStr(AA) + ": " + CStr(NN(FAN.Capabilities(AA))) + vbCrLf
'                TXT = TXT + "CapabilityDescriptions " + CStr(AA) + ": " + CStr(NN(FAN.CapabilityDescriptions(AA))) + vbCrLf
'            Next AA
'            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
'            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
'            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
'            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
'            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
'            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
'            '...................SEGUIR
'        Next
'    Else
'        TXT = TXT + "No se encontraron" + vbCrLf
'    End If

'    Set ObjSet = Nothing
'    'The Win32_DiskDrive WMI class represents a physical disk drive as seen by a _
'    computer running the Win32 operating system. Any interface to a Win32 _
'    physical disk drive is a descendent (or member) of this class. The features _
'    of the disk drive seen through this object correspond to the logical and _
'    management characteristics of the drive. In some cases, this may not reflect _
'    the actual physical characteristics of the device. Any object based on another _
'    logical device would not be a member of this class.
'
'
'    TXT = TXT + vbCrLf
'    TXT = TXT + ucase("INFORMACION DEL Win32_DiskDrive" + vbCrLf
'    Set ObjSet = SERV.InstancesOf("Win32_DiskDrive")
'    If ObjSet.Count > 0 Then
'        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
'        For Each FAN In ObjSet
'            TXT = TXT + "--------------------------------------" + vbCrLf
'            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
'            TXT = TXT + "BytesPerSector: " + CStr(NN(FAN.BytesPerSector)) + vbCrLf
'            For AA = 0 To 9
'                TXT = TXT + "Capabilities " + CStr(AA) + ": " + CStr(NN(FAN.Capabilities(AA))) + vbCrLf
'                TXT = TXT + "CapabilityDescriptions " + CStr(AA) + ": " + CStr(NN(FAN.CapabilityDescriptions(AA))) + vbCrLf
'            Next AA
'            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
'            TXT = TXT + "CompressionMethod: " + CStr(NN(FAN.CompressionMethod)) + vbCrLf
'            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
'            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
'            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
'            TXT = TXT + "DefaultBlockSize: " + CStr(NN(FAN.DefaultBlockSize)) + vbCrLf
'            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
'            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
'            '...................SEGUIR
'        Next
'    Else
'        TXT = TXT + "No se encontraron" + vbCrLf
'    End If

'    Set ObjSet = Nothing
'    'The Win32_FloppyDrive WMI class manages the capabilities of a floppy disk drive.
'
'    TXT = TXT + vbCrLf
'    TXT = TXT + ucase("INFORMACION DEL Win32_FloppyDrive" + vbCrLf
'    Set ObjSet = SERV.InstancesOf("Win32_FloppyDrive")
'    If ObjSet.Count > 0 Then
'        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
'        For Each FAN In ObjSet
'            TXT = TXT + "--------------------------------------" + vbCrLf
'            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
'            TXT = TXT + "Binary: " + CStr(NN(FAN.Binary)) + vbCrLf
'            For AA = 0 To 9
'                TXT = TXT + "Capabilities " + CStr(AA) + ": " + CStr(NN(FAN.Capabilities(AA))) + vbCrLf
'                TXT = TXT + "CapabilityDescriptions " + CStr(AA) + ": " + CStr(NN(FAN.CapabilityDescriptions(AA))) + vbCrLf
'            Next AA
'            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
'            TXT = TXT + "CompressionMethod: " + CStr(NN(FAN.CompressionMethod)) + vbCrLf
'            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
'            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
'            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
'            TXT = TXT + "DefaultBlockSize: " + CStr(NN(FAN.DefaultBlockSize)) + vbCrLf
'            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
'            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
'            '...................SEGUIR
'        Next
'    Else
'        TXT = TXT + "No se encontraron" + vbCrLf
'    End If
'


'    Set ObjSet = Nothing
'    'The Win32_LogicalDisk WMI class represents a data source that resolves to an actual local storage device on a Win32 system.
'
'    TXT = TXT + vbCrLf
'    TXT = TXT + ucase("INFORMACION DEL Win32_LogicalDisk" + vbCrLf
'    Set ObjSet = SERV.InstancesOf("Win32_LogicalDisk")
'    If ObjSet.Count > 0 Then
'        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
'        For Each FAN In ObjSet
'            TXT = TXT + "--------------------------------------" + vbCrLf
'            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
'            TXT = TXT + "Access: " + CStr(NN(FAN.Access)) + vbCrLf
'            TXT = TXT + "BlockSize: " + CStr(NN(FAN.BlockSize)) + vbCrLf
'            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
'            TXT = TXT + "Compressed: " + CStr(NN(FAN.Compressed)) + vbCrLf
'            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
'            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
'            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
'            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
'            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
'            '...................SEGUIR
'        Next
'    Else
'        TXT = TXT + "No se encontraron" + vbCrLf
'    End If


'    Set ObjSet = Nothing
'    'The Win32_TapeDrive WMI class represents a tape drive on a Win32 computer. Tape drives are primarily distinguished by the fact that they can be accessed only sequentially.
'
'    TXT = TXT + vbCrLf
'    TXT = TXT + ucase("INFORMACION DEL Win32_TapeDrive" + vbCrLf
'    Set ObjSet = SERV.InstancesOf("Win32_TapeDrive")
'    If ObjSet.Count > 0 Then
'        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
'        For Each FAN In ObjSet
'            TXT = TXT + "--------------------------------------" + vbCrLf
'            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
'            For AA = 0 To 9
'                TXT = TXT + "Capabilities " + CStr(AA) + ": " + CStr(NN(FAN.Capabilities(AA))) + vbCrLf
'                TXT = TXT + "CapabilityDescriptions " + CStr(AA) + ": " + CStr(NN(FAN.CapabilityDescriptions(AA))) + vbCrLf
'            Next AA
'            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
'            TXT = TXT + "Compression: " + CStr(NN(FAN.Compression)) + vbCrLf
'            TXT = TXT + "CompressionMethod: " + CStr(NN(FAN.CompressionMethod)) + vbCrLf
'            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
'            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
'            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
'            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
'            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
'            '...................SEGUIR
'        Next
'    Else
'        TXT = TXT + "No se encontraron" + vbCrLf
'    End If

'    Set ObjSet = Nothing
'    'The Win32_1394Controller WMI class represents the capabilities and management of a 1394 controller. IEEE 1394 is a specification for a high speed serial bus.
'
'
'    TXT = TXT + vbCrLf
'    TXT = TXT + ucase("INFORMACION DEL Win32_1394Controller" + vbCrLf
'    Set ObjSet = SERV.InstancesOf("Win32_1394Controller")
'    If ObjSet.Count > 0 Then
'        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
'        For Each FAN In ObjSet
'            TXT = TXT + "--------------------------------------" + vbCrLf
'            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
'            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
'            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
'            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
'            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
'            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
'            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
'            '...................SEGUIR
'        Next
'    Else
'        TXT = TXT + "No se encontraron" + vbCrLf
'    End If

'    Set ObjSet = Nothing
'    'The Win32_1394ControllerDevice association WMI class relates the high-speed serial bus (IEEE 1394 Firewire) Controller and the CIM_LogicalDevice instance connected to it. This serial bus provides enhanced connectivity for a wide range of devices, including consumer audio/video components, storage peripherals, other computers, and portable devices. IEEE 1394 has been adopted by the consumer electronics industry and provides a Plug and Play-compatible expansion interface.
'
'
'    TXT = TXT + vbCrLf
'    TXT = TXT + ucase("INFORMACION DEL Win32_1394ControllerDevice" + vbCrLf
'    Set ObjSet = SERV.InstancesOf("Win32_1394ControllerDevice")
'    If ObjSet.Count > 0 Then
'        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
'        For Each FAN In ObjSet
'            TXT = TXT + "--------------------------------------" + vbCrLf
'            TXT = TXT + "AccessState: " + CStr(NN(FAN.AccessState)) + vbCrLf
'            TXT = TXT + "Antecedent: " + CStr(NN(FAN.Antecedent)) + vbCrLf
'            TXT = TXT + "Dependent: " + CStr(NN(FAN.Dependent)) + vbCrLf
'            TXT = TXT + "NegotiatedDataWidth: " + CStr(NN(FAN.NegotiatedDataWidth)) + vbCrLf
'            TXT = TXT + "NegotiatedSpeed: " + CStr(NN(FAN.NegotiatedSpeed)) + vbCrLf
'            TXT = TXT + "NumberOfHardResets: " + CStr(NN(FAN.NumberOfHardResets)) + vbCrLf
'            TXT = TXT + "NumberOfSoftResets: " + CStr(NN(FAN.NumberOfSoftResets)) + vbCrLf
'        Next
'    Else
'        TXT = TXT + "No se encontraron" + vbCrLf
'    End If


'    Set ObjSet = Nothing
'    'The Win32_AllocatedResource association WMI class relates a logical device _
'    to a system resource. The class is used to discover which resources, such as _
'    IRQs or DMA channels, are in-use by a specific device. This class has been _
'    deprecated in favor of the Win32_PNPAllocatedResource class.
'
'    TXT = TXT + vbCrLf
'    TXT = TXT + ucase("INFORMACION DEL Win32_AllocatedResource" + vbCrLf
'    Set ObjSet = SERV.InstancesOf("Win32_AllocatedResource")
'    If ObjSet.Count > 0 Then
'        TXT = TXT + "--------------------------------------" + vbCrLf
'        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
'        For Each FAN In ObjSet
'            TXT = TXT + "Antecedent: " + CStr(NN(FAN.Antecedent)) + vbCrLf
'            TXT = TXT + "Dependent: " + CStr(NN(FAN.Dependent)) + vbCrLf
'        Next
'    Else
'        TXT = TXT + "No se encontraron" + vbCrLf
'    End If

    Set ObjSet = Nothing
    'The Win32_AssociatedProcessorMemory association WMI class relates a _
    processor and its cache memory.

    TXT = TXT + vbCrLf
    TXT = TXT + UCase("INFORMACION DEL Win32_AssociatedProcessorMemory") + vbCrLf
    Set ObjSet = SERV.InstancesOf("Win32_AssociatedProcessorMemory")
    If ObjSet.Count > 0 Then
        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
        For Each FAN In ObjSet
            TXT = TXT + "--------------------------------------" + vbCrLf
            TXT = TXT + "Antecedent: " + CStr(NN(FAN.Antecedent)) + vbCrLf
            TXT = TXT + "BusSpeed: " + CStr(NN(FAN.BusSpeed)) + vbCrLf
            TXT = TXT + "Dependent: " + CStr(NN(FAN.Dependent)) + vbCrLf
        Next
    Else
        TXT = TXT + "No se encontraron" + vbCrLf
    End If

    Set ObjSet = Nothing
    'The Win32_BaseBoard WMI class represents a baseboard (also known as a motherboard or system board).

    TXT = TXT + vbCrLf
    TXT = TXT + UCase("INFORMACION DEL Win32_BaseBoard") + vbCrLf
    Set ObjSet = SERV.InstancesOf("Win32_BaseBoard")
    If ObjSet.Count > 0 Then
        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
        For Each FAN In ObjSet
            TXT = TXT + "--------------------------------------" + vbCrLf
            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
            nCH = 0
            nCH = UBound(FAN.ConfigOptions)
            For A = 0 To nCH
                TXT = TXT + "ConfigOptions " + CStr(A) + ": " + CStr(NN(FAN.ConfigOptions(A))) + vbCrLf
            Next A
            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
            TXT = TXT + "Depth: " + CStr(NN(FAN.Depth)) + vbCrLf
            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
            TXT = TXT + "Height: " + CStr(NN(FAN.Height)) + vbCrLf
            TXT = TXT + "HostingBoard: " + CStr(NN(FAN.HostingBoard)) + vbCrLf
            TXT = TXT + "HotSwappable: " + CStr(NN(FAN.HotSwappable)) + vbCrLf
            TXT = TXT + "InstallDate: " + CStr(NN(FAN.InstallDate)) + vbCrLf
            TXT = TXT + "Manufacturer: " + CStr(NN(FAN.Manufacturer)) + vbCrLf
            TXT = TXT + "Model: " + CStr(NN(FAN.Model)) + vbCrLf
            TXT = TXT + "Name: " + CStr(NN(FAN.Name)) + vbCrLf
            TXT = TXT + "OtherIdentifyingInfo: " + CStr(NN(FAN.OtherIdentifyingInfo)) + vbCrLf
            TXT = TXT + "PartNumber: " + CStr(NN(FAN.PartNumber)) + vbCrLf
            TXT = TXT + "PoweredOn: " + CStr(NN(FAN.PoweredOn)) + vbCrLf
            TXT = TXT + "Product: " + CStr(NN(FAN.Product)) + vbCrLf
            TXT = TXT + "Removable: " + CStr(NN(FAN.Removable)) + vbCrLf
            TXT = TXT + "Replaceable: " + CStr(NN(FAN.Replaceable)) + vbCrLf
            TXT = TXT + "RequirementsDescription: " + CStr(NN(FAN.RequirementsDescription)) + vbCrLf
            TXT = TXT + "RequiresDaughterBoard: " + CStr(NN(FAN.RequiresDaughterBoard)) + vbCrLf
            TXT = TXT + "SerialNumber: " + CStr(NN(FAN.SerialNumber)) + vbCrLf
            TXT = TXT + "SKU: " + CStr(NN(FAN.SKU)) + vbCrLf
            TXT = TXT + "SlotLayout: " + CStr(NN(FAN.SlotLayout)) + vbCrLf
            TXT = TXT + "SpecialRequirements: " + CStr(NN(FAN.SpecialRequirements)) + vbCrLf
            TXT = TXT + "Status: " + CStr(NN(FAN.SlotLayout)) + vbCrLf
            TXT = TXT + "Tag: " + CStr(NN(FAN.Tag)) + vbCrLf
            TXT = TXT + "Version: " + CStr(NN(FAN.Version)) + vbCrLf
            TXT = TXT + "Weight: " + CStr(NN(FAN.Weight)) + vbCrLf
            TXT = TXT + "Width: " + CStr(NN(FAN.Width)) + vbCrLf
        Next
    Else
        TXT = TXT + "No se encontraron" + vbCrLf
    End If


    Set ObjSet = Nothing
    'The Win32_Bus WMI class represents a physical bus as seen by a Win32 operating system. Any instance of a Win32 bus is a descendent (or member) of this class.

    TXT = TXT + vbCrLf
    TXT = TXT + UCase("INFORMACION DEL Win32_Bus") + vbCrLf
    Set ObjSet = SERV.InstancesOf("Win32_Bus")
    If ObjSet.Count > 0 Then
        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
        For Each FAN In ObjSet
            TXT = TXT + "--------------------------------------" + vbCrLf
            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
            TXT = TXT + "BusNum: " + CStr(NN(FAN.BusNum)) + vbCrLf
            TXT = TXT + "BusType: " + CStr(NN(FAN.BusType)) + vbCrLf
            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
            TXT = TXT + "ErrorCleared: " + CStr(NN(FAN.ErrorCleared)) + vbCrLf
            TXT = TXT + "ErrorDescription: " + CStr(NN(FAN.ErrorDescription)) + vbCrLf
            TXT = TXT + "InstallDate: " + CStr(NN(FAN.InstallDate)) + vbCrLf
            TXT = TXT + "LastErrorCode: " + CStr(NN(FAN.LastErrorCode)) + vbCrLf
            TXT = TXT + "Name: " + CStr(NN(FAN.Name)) + vbCrLf
            TXT = TXT + "PNPDeviceID: " + CStr(NN(FAN.PNPDeviceID)) + vbCrLf
            nCH = 0
            nCH = UBound(FAN.PowerManagementCapabilities)
            For A = 0 To nCH
                TXT = TXT + "PowerManagementCapabilities " + CStr(A) + ": " + CStr(NN(FAN.PowerManagementCapabilities(A))) + vbCrLf
            Next A
            TXT = TXT + "PowerManagementSupported: " + CStr(NN(FAN.PowerManagementSupported)) + vbCrLf
            TXT = TXT + "Status: " + CStr(NN(FAN.Status)) + vbCrLf
            TXT = TXT + "StatusInfo: " + CStr(NN(FAN.StatusInfo)) + vbCrLf
            TXT = TXT + "SystemCreationClassName: " + CStr(NN(FAN.SystemCreationClassName)) + vbCrLf
            TXT = TXT + "SystemName: " + CStr(NN(FAN.SystemName)) + vbCrLf
        Next
    Else
        TXT = TXT + "No se encontraron" + vbCrLf
    End If
    
    Set ObjSet = Nothing
    'The Win32_MotherboardDevice WMI class represents a device that contains the central components of the Win32 computer system

    TXT = TXT + vbCrLf
    TXT = TXT + UCase("INFORMACION DEL Win32_MotherboardDevice") + vbCrLf
    Set ObjSet = SERV.InstancesOf("Win32_MotherboardDevice")
    If ObjSet.Count > 0 Then
        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
        For Each FAN In ObjSet
            TXT = TXT + "--------------------------------------" + vbCrLf
            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
            TXT = TXT + "ErrorCleared: " + CStr(NN(FAN.ErrorCleared)) + vbCrLf
            TXT = TXT + "ErrorDescription: " + CStr(NN(FAN.ErrorDescription)) + vbCrLf
            TXT = TXT + "InstallDate: " + CStr(NN(FAN.InstallDate)) + vbCrLf
            TXT = TXT + "LastErrorCode: " + CStr(NN(FAN.LastErrorCode)) + vbCrLf
            TXT = TXT + "Name: " + CStr(NN(FAN.Name)) + vbCrLf
            TXT = TXT + "PNPDeviceID: " + CStr(NN(FAN.PNPDeviceID)) + vbCrLf
            nCH = 0
            nCH = UBound(FAN.PowerManagementCapabilities)
            For A = 0 To nCH
                TXT = TXT + "PowerManagementCapabilities " + CStr(A) + ": " + CStr(NN(FAN.PowerManagementCapabilities(A))) + vbCrLf
            Next A
            TXT = TXT + "PowerManagementSupported: " + CStr(NN(FAN.PowerManagementSupported)) + vbCrLf
            TXT = TXT + "PrimaryBusType: " + CStr(NN(FAN.PrimaryBusType)) + vbCrLf
            TXT = TXT + "RevisionNumber: " + CStr(NN(FAN.RevisionNumber)) + vbCrLf
            TXT = TXT + "SecondaryBusType: " + CStr(NN(FAN.SecondaryBusType)) + vbCrLf
            TXT = TXT + "Status: " + CStr(NN(FAN.Status)) + vbCrLf
            TXT = TXT + "StatusInfo: " + CStr(NN(FAN.StatusInfo)) + vbCrLf
            TXT = TXT + "SystemCreationClassName: " + CStr(NN(FAN.SystemCreationClassName)) + vbCrLf
            TXT = TXT + "SystemName: " + CStr(NN(FAN.SystemName)) + vbCrLf
        Next
    Else
        TXT = TXT + "No se encontraron" + vbCrLf
    End If
    
    Set ObjSet = Nothing
    'The Win32_ParallelPort WMI class represents the properties of a parallel port on a Win32 computer system

    TXT = TXT + vbCrLf
    TXT = TXT + UCase("INFORMACION DEL Win32_ParallelPort") + vbCrLf
    Set ObjSet = SERV.InstancesOf("Win32_ParallelPort")
    If ObjSet.Count > 0 Then
        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
        For Each FAN In ObjSet
            TXT = TXT + "--------------------------------------" + vbCrLf
            TXT = TXT + "Availability: " + CStr(NN(FAN.Availability)) + vbCrLf
            nCH = 0
            nCH = UBound(FAN.Capabilities)
            For A = 0 To nCH
                TXT = TXT + "Capabilities " + CStr(A) + ": " + CStr(NN(FAN.Capabilities(A))) + vbCrLf
                TXT = TXT + "CapabilityDescriptions " + CStr(A) + ": " + CStr(NN(FAN.CapabilityDescriptions(A))) + vbCrLf
            Next A
            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
            TXT = TXT + "ConfigManagerErrorCode: " + CStr(NN(FAN.ConfigManagerErrorCode)) + vbCrLf
            TXT = TXT + "ConfigManagerUserConfig: " + CStr(NN(FAN.ConfigManagerUserConfig)) + vbCrLf
            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
            TXT = TXT + "DeviceID: " + CStr(NN(FAN.DeviceID)) + vbCrLf
            TXT = TXT + "DMASupport: " + CStr(NN(FAN.DMASupport)) + vbCrLf
            TXT = TXT + "ErrorCleared: " + CStr(NN(FAN.ErrorCleared)) + vbCrLf
            TXT = TXT + "ErrorDescription: " + CStr(NN(FAN.ErrorDescription)) + vbCrLf
            TXT = TXT + "InstallDate: " + CStr(NN(FAN.InstallDate)) + vbCrLf
            TXT = TXT + "LastErrorCode: " + CStr(NN(FAN.LastErrorCode)) + vbCrLf
            TXT = TXT + "MaxNumberControlled: " + CStr(NN(FAN.MaxNumberControlled)) + vbCrLf
            TXT = TXT + "Name: " + CStr(NN(FAN.Name)) + vbCrLf
            TXT = TXT + "OSAutoDiscovered: " + CStr(NN(FAN.OSAutoDiscovered)) + vbCrLf
            TXT = TXT + "PNPDeviceID: " + CStr(NN(FAN.PNPDeviceID)) + vbCrLf
            nCH = 0
            nCH = UBound(FAN.PowerManagementCapabilities)
            For A = 0 To nCH
                TXT = TXT + "PowerManagementCapabilities " + CStr(A) + ": " + CStr(NN(FAN.PowerManagementCapabilities(A))) + vbCrLf
            Next A
            TXT = TXT + "PowerManagementSupported: " + CStr(NN(FAN.PowerManagementSupported)) + vbCrLf
            TXT = TXT + "ProtocolSupported: " + CStr(NN(FAN.ProtocolSupported)) + vbCrLf
            TXT = TXT + "Status: " + CStr(NN(FAN.Status)) + vbCrLf
            TXT = TXT + "StatusInfo: " + CStr(NN(FAN.StatusInfo)) + vbCrLf
            TXT = TXT + "SystemCreationClassName: " + CStr(NN(FAN.SystemCreationClassName)) + vbCrLf
            TXT = TXT + "SystemName: " + CStr(NN(FAN.SystemName)) + vbCrLf
            TXT = TXT + "TimeOfLastReset: " + CStr(NN(FAN.TimeOfLastReset)) + vbCrLf
        Next
    Else
        TXT = TXT + "No se encontraron" + vbCrLf
    End If
    
    Set ObjSet = Nothing
    'The Win32_PhysicalMemory WMI class represents a physical memory device located on a computer system as available to the operating system

    TXT = TXT + vbCrLf
    TXT = TXT + UCase("INFORMACION DEL Win32_PhysicalMemory") + vbCrLf
    Set ObjSet = SERV.InstancesOf("Win32_PhysicalMemory")
    If ObjSet.Count > 0 Then
        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
        For Each FAN In ObjSet
            TXT = TXT + "--------------------------------------" + vbCrLf
            TXT = TXT + "BankLabel: " + CStr(NN(FAN.Availability)) + vbCrLf
            TXT = TXT + "Capacity: " + CStr(NN(FAN.Capacity)) + vbCrLf
            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
            TXT = TXT + "DataWidth: " + CStr(NN(FAN.DataWidth)) + vbCrLf
            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
            TXT = TXT + "DeviceLocator: " + CStr(NN(FAN.DeviceLocator)) + vbCrLf
            TXT = TXT + "FormFactor: " + CStr(NN(FAN.FormFactor)) + vbCrLf
            TXT = TXT + "HotSwappable: " + CStr(NN(FAN.HotSwappable)) + vbCrLf
            TXT = TXT + "InstallDate: " + CStr(NN(FAN.InstallDate)) + vbCrLf
            TXT = TXT + "InterleaveDataDepth: " + CStr(NN(FAN.InterleaveDataDepth)) + vbCrLf
            TXT = TXT + "InterleavePosition: " + CStr(NN(FAN.InterleavePosition)) + vbCrLf
            TXT = TXT + "Manufacturer: " + CStr(NN(FAN.Manufacturer)) + vbCrLf
            TXT = TXT + "Model: " + CStr(NN(FAN.Model)) + vbCrLf
            TXT = TXT + "Name: " + CStr(NN(FAN.Name)) + vbCrLf
            TXT = TXT + "OtherIdentifyingInfo: " + CStr(NN(FAN.OtherIdentifyingInfo)) + vbCrLf
            TXT = TXT + "PartNumber: " + CStr(NN(FAN.PartNumber)) + vbCrLf
            TXT = TXT + "PositionInRow: " + CStr(NN(FAN.PositionInRow)) + vbCrLf
            TXT = TXT + "PoweredOn: " + CStr(NN(FAN.PoweredOn)) + vbCrLf
            TXT = TXT + "Removable: " + CStr(NN(FAN.Removable)) + vbCrLf
            TXT = TXT + "Replaceable: " + CStr(NN(FAN.Replaceable)) + vbCrLf
            TXT = TXT + "SerialNumber: " + CStr(NN(FAN.SerialNumber)) + vbCrLf
            TXT = TXT + "SKU: " + CStr(NN(FAN.SKU)) + vbCrLf
            TXT = TXT + "Speed: " + CStr(NN(FAN.Speed)) + vbCrLf
            TXT = TXT + "Status: " + CStr(NN(FAN.Status)) + vbCrLf
            TXT = TXT + "Tag: " + CStr(NN(FAN.Tag)) + vbCrLf
            TXT = TXT + "TotalWidth: " + CStr(NN(FAN.TotalWidth)) + vbCrLf
            TXT = TXT + "TypeDetail: " + CStr(NN(FAN.TypeDetail)) + vbCrLf
            TXT = TXT + "Version: " + CStr(NN(FAN.Version)) + vbCrLf
        Next
    Else
        TXT = TXT + "No se encontraron" + vbCrLf
    End If
    
    Set ObjSet = Nothing
    'The Win32_PortConnector WMI class represents physical connection ports, such as DB-25 pin male, Centronics, and PS/2

    TXT = TXT + vbCrLf
    TXT = TXT + UCase("INFORMACION DEL Win32_PortConnector") + vbCrLf
    Set ObjSet = SERV.InstancesOf("Win32_PortConnector")
    If ObjSet.Count > 0 Then
        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
        For Each FAN In ObjSet
            TXT = TXT + "--------------------------------------" + vbCrLf
            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
            TXT = TXT + "ConnectorPinout: " + CStr(NN(FAN.ConnectorPinout)) + vbCrLf
            nCH = 0
            nCH = UBound(FAN.ConnectorType)
            For A = 0 To nCH
                TXT = TXT + "ConnectorType " + CStr(A) + ": " + CStr(NN(FAN.ConnectorType(A))) + vbCrLf
            Next A
            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
            TXT = TXT + "ExternalReferenceDesignator: " + CStr(NN(FAN.ExternalReferenceDesignator)) + vbCrLf
            TXT = TXT + "InstallDate: " + CStr(NN(FAN.InstallDate)) + vbCrLf
            TXT = TXT + "InternalReferenceDesignator: " + CStr(NN(FAN.InternalReferenceDesignator)) + vbCrLf
            TXT = TXT + "Manufacturer: " + CStr(NN(FAN.Manufacturer)) + vbCrLf
            TXT = TXT + "Model: " + CStr(NN(FAN.Model)) + vbCrLf
            TXT = TXT + "Name: " + CStr(NN(FAN.Name)) + vbCrLf
            TXT = TXT + "OtherIdentifyingInfo: " + CStr(NN(FAN.OtherIdentifyingInfo)) + vbCrLf
            TXT = TXT + "PartNumber: " + CStr(NN(FAN.PartNumber)) + vbCrLf
            TXT = TXT + "PortType: " + CStr(NN(FAN.PortType)) + vbCrLf
            TXT = TXT + "PoweredOn: " + CStr(NN(FAN.PoweredOn)) + vbCrLf
            TXT = TXT + "SerialNumber: " + CStr(NN(FAN.SerialNumber)) + vbCrLf
            TXT = TXT + "SKU: " + CStr(NN(FAN.SKU)) + vbCrLf
            TXT = TXT + "Status: " + CStr(NN(FAN.Status)) + vbCrLf
            TXT = TXT + "Tag: " + CStr(NN(FAN.Tag)) + vbCrLf
            TXT = TXT + "Version: " + CStr(NN(FAN.Version)) + vbCrLf
        Next
    Else
        TXT = TXT + "No se encontraron" + vbCrLf
    End If
    
    Set ObjSet = Nothing
    'The Win32_OnBoardDevice WMI class represents common adapter devices built into the motherboard (system board

    TXT = TXT + vbCrLf
    TXT = TXT + UCase("INFORMACION DEL Win32_OnBoardDevice") + vbCrLf
    Set ObjSet = SERV.InstancesOf("Win32_OnBoardDevice")
    If ObjSet.Count > 0 Then
        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
        For Each FAN In ObjSet
            TXT = TXT + "--------------------------------------" + vbCrLf
            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
            TXT = TXT + "CreationClassName: " + CStr(NN(FAN.CreationClassName)) + vbCrLf
            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
            nCH = 0
            nCH = FAN.DeviceType
            Select Case nCH
                Case 1: TXT = TXT + "DeviceType: 1 = Other" + vbCrLf
                Case 2: TXT = TXT + "DeviceType: 2 = Unknown" + vbCrLf
                Case 3: TXT = TXT + "DeviceType: 3 = Video" + vbCrLf
                Case 4: TXT = TXT + "DeviceType: 4 = SCSI Controller" + vbCrLf
                Case 5: TXT = TXT + "DeviceType: 5 = Ethernet" + vbCrLf
                Case 6: TXT = TXT + "DeviceType: 6 = Token Ring" + vbCrLf
                Case 7: TXT = TXT + "DeviceType: 7 = Sound" + vbCrLf
                Case Else: TXT = TXT + "DeviceType: " + CStr(nCH) + vbCrLf
            End Select
            TXT = TXT + "Enabled: " + CStr(NN(FAN.Enabled)) + vbCrLf
            TXT = TXT + "HotSwappable: " + CStr(NN(FAN.HotSwappable)) + vbCrLf
            TXT = TXT + "InstallDate: " + CStr(NN(FAN.InstallDate)) + vbCrLf
            TXT = TXT + "Manufacturer: " + CStr(NN(FAN.Manufacturer)) + vbCrLf
            TXT = TXT + "Model: " + CStr(NN(FAN.Model)) + vbCrLf
            TXT = TXT + "Name: " + CStr(NN(FAN.Name)) + vbCrLf
            TXT = TXT + "OtherIdentifyingInfo: " + CStr(NN(FAN.OtherIdentifyingInfo)) + vbCrLf
            TXT = TXT + "PartNumber: " + CStr(NN(FAN.PartNumber)) + vbCrLf
            TXT = TXT + "PoweredOn: " + CStr(NN(FAN.PoweredOn)) + vbCrLf
            TXT = TXT + "Removable: " + CStr(NN(FAN.Removable)) + vbCrLf
            TXT = TXT + "Replaceable: " + CStr(NN(FAN.Replaceable)) + vbCrLf
            TXT = TXT + "SerialNumber: " + CStr(NN(FAN.SerialNumber)) + vbCrLf
            TXT = TXT + "SKU: " + CStr(NN(FAN.SKU)) + vbCrLf
            TXT = TXT + "Status: " + CStr(NN(FAN.Status)) + vbCrLf
            TXT = TXT + "Tag: " + CStr(NN(FAN.Tag)) + vbCrLf
            TXT = TXT + "Version: " + CStr(NN(FAN.Version)) + vbCrLf
        Next
    Else
        TXT = TXT + "No se encontraron" + vbCrLf
    End If
    
    
'    Set ObjSet = Nothing
'    'The Win32_ClassicCOMClass WMI class represents the properties of a COM component
'
'    TXT = TXT + vbCrLf
'    TXT = TXT + ucase("INFORMACION DEL Win32_ClassicCOMClass" + vbCrLf
'    Set ObjSet = serv.InstancesOf("Win32_ClassicCOMClass")
'    If ObjSet.Count > 0 Then
'        TXT = TXT + "Se encontaron: " + CStr(ObjSet.Count) + vbCrLf
'        For Each FAN In ObjSet
'            A = A + 1
'            If A = 100 Then Exit For'son MUUUUUUUUCHOS
'            TXT = TXT + "--------------------------------------" + vbCrLf
'            TXT = TXT + "Caption: " + CStr(NN(FAN.Caption)) + vbCrLf
'            TXT = TXT + "ComponentId: " + CStr(NN(FAN.ComponentId)) + vbCrLf
'            TXT = TXT + "Description: " + CStr(NN(FAN.Description)) + vbCrLf
'            TXT = TXT + "InstallDate: " + CStr(NN(FAN.InstallDate)) + vbCrLf
'            TXT = TXT + "Name: " + CStr(NN(FAN.Name)) + vbCrLf
'            TXT = TXT + "Status: " + CStr(NN(FAN.Status)) + vbCrLf
'        Next
'    Else
'        TXT = TXT + "No se encontraron" + vbCrLf
'    End If
    
    Set ObjSet = Nothing
    
    'viejo y glorioso System INFO
    Dim INFO As SYSTEM_INFO
    GetSystemInfo INFO
    TXT = TXT + vbCrLf + "System_Info viejo NOMAS" + vbCrLf
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
    
    Exit Sub
ErrWBMem:
    TXT = TXT + "--------------------------------------ERROR------------------------" + vbCrLf
    TXT = TXT + UCase(Err.Description + "(" + CStr(Err.Number) + ") " + Err.Source + vbCrLf)
    TXT = TXT + "--------------------------------------ERROR------------------------" + vbCrLf + vbCrLf
    Resume Next
End Sub

Private Sub Form_Resize()
    TXT.Width = Me.Width - TXT.Left - 250
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
